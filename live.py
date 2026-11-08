"""GitHub Actions'da jonli (headful) browser — VNC orqali kuzatish uchun.

Faqat token inject qilinadi; ERI kalit hech qachon yuborilmaydi.
Logga faqat sonlar chiqadi.
"""
import os
import re
import time

from playwright.sync_api import sync_playwright

TOKEN = os.environ["AUTH_TOKEN"]
MINUTES = float(os.environ.get("LIVE_MINUTES", "15"))
SITE = "https://online-mahalla.uz"
URL = (SITE + "/tables/survey_homes_street?obl_id=35&area_id=3509"
       "&district_id=3509040&street_id=350901168&_mid=3548")

PLATON_SHIM = r"""
(() => {
  if (window.PlatonPluginCore) return;
  let real = null;
  const queue = [];
  const shim = {
    preloadPlugins() { queue.push(Array.prototype.slice.call(arguments)); },
    triggerEvent() { return Promise.resolve(); }
  };
  try {
    Object.defineProperty(window, 'PlatonPluginCore', {
      configurable: true,
      get() { return real || shim; },
      set(v) {
        real = v;
        const q = queue.splice(0);
        for (const args of q) {
          try { v.preloadPlugins && v.preloadPlugins.apply(v, args); } catch (e) {}
        }
      }
    });
  } catch (e) {}
})();
"""

ROWS_JS = """() => {
  const t = [...document.querySelectorAll('table')]
    .find(x => (x.innerText || '').includes('кадастр'));
  return t ? t.querySelectorAll('tbody tr').length : 0;
}"""


def main():
    with sync_playwright() as p:
        b = p.chromium.launch(
            headless=False,
            args=["--no-sandbox", "--window-size=1440,900",
                  "--window-position=0,0", "--disable-infobars"],
        )
        ctx = b.new_context(viewport={"width": 1440, "height": 820})
        ctx.add_init_script(
            "try { localStorage.setItem('auth_token', %r);"
            " localStorage.setItem('refresh_token', %r);"
            " localStorage.setItem('auth_user', %r); } catch (e) {}"
            % (TOKEN, TOKEN, '{"id":1}')
        )
        ctx.add_init_script(PLATON_SHIM)
        pg = ctx.new_page()

        def _route(route):
            r = route.request
            if r.resource_type == "other" and re.search(r"/(?:js|css)/chunk-", r.url):
                route.abort()
            else:
                route.continue_()

        pg.route("**/*", _route)

        pg.goto(SITE + "/", wait_until="domcontentloaded", timeout=90000)
        pg.wait_for_timeout(3000)
        pg.goto(URL, wait_until="domcontentloaded", timeout=90000)

        rows = 0
        for _ in range(90):
            pg.wait_for_timeout(1000)
            if "/login" in pg.url:
                break
            rows = pg.evaluate(ROWS_JS)
            if rows:
                break
        print("LIVE_READY rows=%d path=%s" % (rows, pg.url.split("?")[0]), flush=True)

        end = time.time() + MINUTES * 60
        while time.time() < end:
            pg.wait_for_timeout(30000)
            try:
                rows = pg.evaluate(ROWS_JS)
                print("heartbeat rows=%d path=%s" % (rows, pg.url.split("?")[0]),
                      flush=True)
            except Exception as e:
                print("heartbeat error:", str(e)[:120], flush=True)
        b.close()


if __name__ == "__main__":
    main()

"""GitHub Actions smoke test: sayt GitHub IP'dan ochilishini tekshiradi.

Faqat SONLAR logga chiqadi: qatorlar soni, HTTP kod. Shaxsiy ma'lumot yo'q.
"""
import os
import re

from playwright.sync_api import sync_playwright

TOKEN = os.environ["AUTH_TOKEN"]
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


def main():
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        ctx = b.new_context(viewport={"width": 1440, "height": 900})
        ctx.add_init_script(
            "try { localStorage.setItem('auth_token', %r); } catch (e) {}" % TOKEN
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
        status = pg.evaluate(
            """async (t) => {
                 try {
                   const r = await fetch('https://api.online-mahalla.uz/auth/user_info',
                                         {headers: {authorization: 'Bearer ' + t}});
                   return r.status;
                 } catch (e) { return -1; }
               }""",
            TOKEN,
        )
        print("user_info:", status, flush=True)

        pg.goto(URL, wait_until="domcontentloaded", timeout=90000)
        rows = 0
        for _ in range(90):
            pg.wait_for_timeout(1000)
            if "/login" in pg.url:
                break
            rows = pg.evaluate(
                """() => {
                     const t = [...document.querySelectorAll('table')]
                       .find(x => (x.innerText || '').includes('кадастр'));
                     return t ? t.querySelectorAll('tbody tr').length : 0;
                   }"""
            )
            if rows:
                break

        path = pg.url.split("?")[0]
        print("final_path:", path, flush=True)
        print("table_rows:", rows, flush=True)
        ok = rows > 0 and "/login" not in pg.url and status == 200
        print("SMOKE OK" if ok else "SMOKE FAIL", flush=True)
        b.close()
        raise SystemExit(0 if ok else 1)


if __name__ == "__main__":
    main()

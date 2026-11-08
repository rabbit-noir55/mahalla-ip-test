"""Yangi baza (2026) bo'yicha xonadon a'zolarini qo'shish.

Ish tartibi (har bir xonadon uchun):
    1. Sayt ko'cha jadvali -> xonadon raqami, kadastr, EGASI F.I.Sh. + JSHSHIR
    2. Egasi bo'yicha bazadan xonadon bloki topiladi (JSHSHIR -> F.I.Sh.)
    3. Saytdagi a'zolar bilan solishtiriladi -> yetishmayotganlar qo'shiladi
    4. Har biri: guvohnoma (bola) yoki pasport (katta) -> қидириш -> Сақлаш
    5. Natija Excel hisobotga yoziladi

Ishga tushirish:
    python mahalla_new.py --street 350901168 --slot 1
    python mahalla_new.py --street 350901145 --slot 2 --limit 100
"""
import argparse
import json
import os
import re
import sys
import time
import traceback
from datetime import date, datetime

from playwright.sync_api import sync_playwright

import base_new
import report_log as rl

try:                       # Windows konsolida Қ/Ў kabi harflar yiqilmasin
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SITE = "https://online-mahalla.uz"
Q = "obl_id=35&area_id=3509&district_id=3509040"
MID = "&_mid=3548"

DOC_PASS = "фуқаролик паспорти / ID карта"
DOC_CERT = "туғилганлик ҳақида гувоҳнома"
RELATION = "Бошқа"
EDUCATION = "Ўрта"

SEARCH_DOC = "button[title='Ҳужжат рақами ва туғилган санаси бўйича қидириш']"
SEARCH_PINFL = "button.pa-inline-btn"

LOG_FILE = None
DRY = False

# Sayt app'i sekin tarmoqda yiqiladi: notification.js PlatonPluginCore ni
# core yaratilishidan oldin chaqiradi -> "PlatonPluginCore is not defined".
# Bu shim chaqiruvlarni navbatga qo'yib, core paydo bo'lganda qayta yuboradi.
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


def install_page_fixes(pg):
    """App boot fix + ortiqcha preload chunk'larni bloklash."""
    pg.add_init_script(PLATON_SHIM)

    def _route(route):
        r = route.request
        if r.resource_type == "other" and re.search(r"/(?:js|css)/chunk-", r.url):
            route.abort()          # preload hint — kerak bo'lsa webpack o'zi yuklaydi
        else:
            route.continue_()

    pg.route("**/*", _route)


def log(msg=""):
    line = msg if not msg else f"[{datetime.now():%H:%M:%S}] {msg}"
    try:
        print(line, flush=True)
    except Exception:
        try:
            sys.stdout.write(line.encode("ascii", "replace").decode("ascii") + "\n")
            sys.stdout.flush()
        except Exception:
            pass
    if LOG_FILE:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except OSError:
            pass


def random_phone():
    return "+998(91)%03d-%02d-%02d" % (
        __import__("random").randint(100, 999),
        __import__("random").randint(10, 99),
        __import__("random").randint(10, 99),
    )


def norm_phone(raw):
    """Bazadagi 9 xonali raqamni sayt formatiga keltiradi.

    '993347642' -> '+998(99)334-76-42'; noto'g'ri bo'lsa random.
    """
    d = re.sub(r"\D", "", str(raw or ""))
    if len(d) == 9:
        d = "998" + d
    if len(d) == 12 and d.startswith("998"):
        return "+998(%s)%s-%s-%s" % (d[3:5], d[5:8], d[8:10], d[10:12])
    return random_phone()


def norm(s):
    s = (s or "").upper().replace("Ў", "У").replace("Қ", "Қ").replace("Ғ", "Ғ")
    s = re.sub(r"[^A-ZА-ЯЁ\s]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# ------------------------------------------------------------------ browser
def ensure_pages(ctx, n, anchor):
    """Kamida n ta online-mahalla oynasi bo'lishini ta'minlaydi."""
    def pages():
        return [x for x in ctx.pages if "online-mahalla.uz" in x.url]

    while len(pages()) < n:
        try:
            sess = ctx.new_cdp_session(anchor)
            sess.send("Target.createTarget", {"url": SITE + "/", "newWindow": True})
        except Exception:
            ctx.new_page()
        # eventlarni haydash
        for _ in range(10):
            anchor.wait_for_timeout(400)
    return pages()


def street_url(street_id):
    return f"{SITE}/tables/survey_homes_street?{Q}&street_id={street_id}{MID}"


def open_street(pg, street_id, attempts=3):
    for i in range(attempts):
        try:
            pg.goto(street_url(street_id), wait_until="commit", timeout=60000)
        except Exception as e:
            log(f"  goto xatosi: {e}")
        pg.wait_for_timeout(2000)
        if "/login" in pg.url:
            return -1          # login kerak — qayta urinmaymiz
        for _ in range(60):    # 30 s
            pg.wait_for_timeout(500)
            n = pg.evaluate("""() => {
                const t = [...document.querySelectorAll('table')]
                    .find(x => (x.innerText||'').includes('кадастр'));
                return t ? t.querySelectorAll('tbody tr').length : 0;
            }""")
            if n:
                return n
        log(f"  ko'cha jadvali ochilmadi ({i + 1}/{attempts})")
    return 0


def go_page(pg, street_id, page):
    """Ko'cha jadvalining N-sahifasiga o'tish."""
    open_street(pg, street_id)
    for _ in range(page - 1):
        ok = pg.evaluate("""() => {
            const b = [...document.querySelectorAll('button')].filter(x => x.offsetWidth > 0 && !x.disabled)
                .find(x => {
                    const t = x.innerText.trim();
                    const al = (x.getAttribute('aria-label') || '').toLowerCase();
                    return t === '›' || t === '>' || al.includes('next');
                });
            if (!b) return false;
            b.click();
            return true;
        }""")
        if not ok:
            return False
        pg.wait_for_timeout(2500)
    return True


def read_houses(pg):
    """Ko'cha jadvali: xonadonlar ro'yxati."""
    return pg.evaluate("""() => {
      const t = [...document.querySelectorAll('table')]
          .find(x => (x.innerText||'').includes('кадастр'));
      if (!t) return [];
      const out = [];
      for (const tr of t.querySelectorAll('tbody tr')) {
        const td = [...tr.querySelectorAll('td')].map(x => (x.innerText||'').trim().replace(/\\s+/g,' '));
        if (!td.length) continue;
        const a = tr.querySelector("a[href*='/forms/survey_homes/']");
        out.push({
          n: td[0] || '', no: td[1] || '', cad: td[2] || '', fio: td[3] || '',
          phone: td[4] || '', pinfl: td[5] || '', href: a ? a.getAttribute('href') : ''
        });
      }
      return out;
    }""")


def open_house(pg, href, attempts=3):
    """Xonadon sahifasini ochadi.

    «Хонадон аъзолари» MATNI/tab'i paydo bo'lishini kutadi (faqat
    [role=tab] emas — ba'zi sahifalarda tablar kech chiziladi).
    Ochilmasa 2 marta qayta yuklaydi, timeout 120 s.
    """
    url = href if href.startswith("http") else SITE + href
    for i in range(attempts):
        try:
            if i == 0 or not pg.url.startswith(url[:60]):
                pg.goto(url, wait_until="commit", timeout=90000)
            else:
                pg.reload(wait_until="commit", timeout=90000)
        except Exception as e:
            log(f"      ... yuklash xatosi ({i + 1}/{attempts}): {e}")
        if "/login" in pg.url:
            return False

        for _ in range(60):                 # 60 * 0.5 = 30 s
            pg.wait_for_timeout(500)
            try:
                ok = pg.evaluate("""() => {
                    const t = document.body ? (document.body.innerText || '') : '';
                    if (t.includes('Хонадон аъзолари')) return true;
                    return [...document.querySelectorAll('[role=tab]')]
                        .some(x => (x.innerText || '').includes('аъзолари'));
                }""")
            except Exception:
                ok = False
            if ok:
                pg.wait_for_timeout(900)
                return True
        log(f"      ... xonadon yuklanmadi, qayta urinish ({i + 1}/{attempts})")
    return False


def read_members(pg, attempts=3):
    """«Хонадон аъзолари» jadvali. O'qilmasa None (dublikat qo'shmaslik uchun)."""
    for attempt in range(attempts):
        try:
            pg.evaluate("""() => {
              const t = [...document.querySelectorAll('[role=tab]')].find(x => x.innerText.includes('Хонадон аъзолари'));
              if (t) t.click();
            }""")
        except Exception:
            pass
        for _ in range(10):
            pg.wait_for_timeout(500)
            res = pg.evaluate("""() => {
              const t = [...document.querySelectorAll('table')].filter(t => t.offsetWidth > 0)
                  .find(t => (t.innerText || '').includes('Қариндошлиги'));
              if (!t) return null;
              const out = [];
              for (const tr of t.querySelectorAll('tbody tr')) {
                const td = [...tr.querySelectorAll('td')].map(x => (x.innerText||'').trim().replace(/\\s+/g,' '));
                if (td.length >= 5) out.push({rel: td[1]||'', fio: td[2]||'', birth: td[3]||'', pinfl: td[4]||''});
              }
              return out;
            }""")
            if res:
                return res
        log(f"      ... a'zolar jadvali o'qilmadi ({attempt + 1}/{attempts})")
    return None


# ------------------------------------------------------------------- modal
def modal(pg):
    return pg.locator(".vm--modal.form-modal").last


def fld(m, name):
    return m.locator(f"fieldset[data-name='{name}']").first


def inp(m, name):
    return fld(m, name).locator("input").first


def pick(m, pg, name, text):
    sel = fld(m, name).locator("select").first
    box = sel.locator("xpath=following-sibling::span[contains(@class,'select2')]").first
    box.click(force=True)
    opts = pg.locator(".select2-container--open .select2-results__option")
    texts = []
    # variantlar AJAX bilan keladi — bo'sh ro'yxatga darhol hukm qilmaymiz
    for _ in range(24):
        pg.wait_for_timeout(150)
        texts = [re.sub(r"\s+", " ", opts.nth(i).inner_text()).strip()
                 for i in range(opts.count())]
        if texts:
            break
    target = re.sub(r"\s+", " ", text).strip()
    idx = texts.index(target) if target in texts else -1
    if idx < 0:
        starts = [i for i, t in enumerate(texts) if t.startswith(target)]
        if len(starts) == 1:
            idx = starts[0]
    if idx < 0:
        pg.keyboard.press("Escape")
        raise RuntimeError(f"variant yo'q: {target!r} / {texts}")
    opts.nth(idx).click()
    pg.wait_for_timeout(200)


def set_value(pg, name, value):
    """Maydonni JS orqali to'ldiradi (Vue model yangilanadi)."""
    got = pg.evaluate("""([name, val]) => {
        const m = [...document.querySelectorAll('.vm--modal.form-modal')].filter(x => x.offsetWidth).pop();
        if (!m) return null;
        const fs = m.querySelector(`fieldset[data-name='${name}']`);
        if (!fs) return null;
        const el = fs.querySelector('input');
        if (!el) return null;
        el.focus();
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(el, val);
        el.dispatchEvent(new Event('input', {bubbles: true}));
        el.dispatchEvent(new Event('change', {bubbles: true}));
        el.dispatchEvent(new Event('blur', {bubbles: true}));
        return el.value;
    }""", [name, str(value)])
    pg.wait_for_timeout(250)
    return got


def type_into(pg, field, value, delay=25):
    field.click()
    pg.keyboard.press("Control+A")
    pg.keyboard.press("Backspace")
    pg.wait_for_timeout(90)
    field.type(str(value), delay=delay)
    pg.wait_for_timeout(150)
    return field.input_value().strip()


def modal_open(pg):
    try:
        return bool(pg.evaluate("""() => [...document.querySelectorAll(
            '.vm--modal.form-modal')].some(x => x.offsetWidth > 0)"""))
    except Exception:
        return False


def set_phone(pg, m, raw):
    """Telefonni to'g'ri formatda yozadi va qiymatni tekshiradi."""
    val = norm_phone(raw)
    try:
        set_value(pg, "phone", val)
        got = inp(m, "phone").input_value().strip()
    except Exception:
        got = ""
    if len(re.sub(r"\D", "", got or "")) != 12:
        try:
            got = type_into(pg, inp(m, "phone"), val)
        except Exception:
            pass
    return got


def validation_errors(pg):
    """Saqlashdan keyingi modal validatsiyasi -> {maydon: matn}.

    Sayt xato matnini .invalid-feedback emas, fieldset ichiga oddiy
    matn qilib qo'yadi: «Майдонни тўлдириш керак», «... мавжуд ...».
    """
    try:
        return pg.evaluate("""() => {
          const m = [...document.querySelectorAll('.vm--modal.form-modal')]
              .filter(x => x.offsetWidth).pop();
          if (!m) return {};
          const out = {};
          for (const fs of m.querySelectorAll('fieldset[data-name]')) {
            const t = (fs.innerText || '').replace(/\\s+/g, ' ').trim();
            if (/тўлдириш керак|тўғри тўлдирил|тўлиқлигини текширинг|хатолик|мавжуд|киритиб бўлмайди|мувофиқ эмас/i.test(t)) {
              out[fs.getAttribute('data-name')] = t.slice(0, 300);
            }
          }
          return out;
        }""") or {}
    except Exception:
        return {}


def save_error_text(resp):
    """POST javobidan (400) o'qiladigan sabab."""
    body = (resp or {}).get("body") or ""
    try:
        j = json.loads(body)
    except Exception:
        return (body or "Saqlanmadi")[:180]
    errs = j.get("errors") or {}
    if isinstance(errs, dict) and errs:
        return " | ".join(str(v) for v in errs.values())
    return str(j.get("message") or j.get("error") or body)[:180]


def human_reason(msg):
    low = (msg or "").lower()
    if "мавжуд" in low and ("киритиб" in low or "қайта" in low):
        detail = re.sub(r"^Ушбу фуқаро\s*", "", (msg or "").strip())
        return "Boshqa xonadonda mavjud: " + detail[:200]
    if "тўлиқлигини" in low:
        return "JSHSHIR to'liq emas"
    if "мувофиқ эмас" in low or "мувофик эмас" in low:
        return "Ma'lumotlar mos kelmadi"
    if "топилмади" in low or "not found" in low:
        return "Ma'lumot topilmadi"
    return (msg or "Saqlanmadi").replace("\n", " ")[:180]


def fix_field(pg, m, name, p):
    """Validatsiya ko'rsatgan maydonni bazadagi ma'lumot bilan to'ldiradi."""
    birth = p["birth"].strftime("%d.%m.%Y") if p["birth"] else ""
    try:
        if name == "relationship":
            pick(m, pg, name, RELATION)
        elif name == "document_type":
            is_cert = p["doc_type"] == "cert" and p["series"] and p["number"]
            pick(m, pg, name, DOC_CERT if is_cert else DOC_PASS)
        elif name == "doc_serial" and p["series"]:
            set_value(pg, name, p["series"])
        elif name == "doc_number" and p["number"]:
            set_value(pg, name, p["number"])
        elif name == "birth_date" and birth:
            set_value(pg, name, birth)
        elif name == "pinfl" and p["pinfl"]:
            set_value(pg, name, p["pinfl"])
        elif name == "phone":
            set_phone(pg, m, p["phone"])
        elif name == "study_level_id":
            pick(m, pg, name, EDUCATION)
        else:
            return False
        return True
    except Exception as e:
        log(f"      ! {name} tuzatilmadi: {e}")
        return False


def close_error(pg):
    """Xato oynasini yopadi. «Фамилия бўйича қидирмоқчимисиз?» да — «Нет/Йўқ»."""
    for _ in range(3):
        info = pg.evaluate("""() => {
          const vis = e => e && e.offsetWidth > 0;
          const m = [...document.querySelectorAll('.vm--modal,.modal,[role=dialog]')]
            .filter(e => vis(e) && /Фамилия|қидирмоқчи|қидирмокчи/i.test(e.innerText || ''))[0];
          if (m) return {kind: 'name'};
          const e = document.getElementById('pa-error-modal');
          return (e && vis(e)) ? {kind: 'err'} : null;
        }""")
        if not info:
            return
        if info["kind"] == "name":
            clicked = pg.evaluate("""() => {
              const vis = e => e && e.offsetWidth > 0;
              const m = [...document.querySelectorAll('.vm--modal,.modal,[role=dialog]')]
                .filter(e => vis(e) && /Фамилия|қидирмоқчи|қидирмокчи/i.test(e.innerText || ''))[0];
              if (!m) return false;
              const bs = [...m.querySelectorAll('button')].filter(vis);
              const no = bs.find(b => /^(нет|йўқ|йук|no|cancel|бекор|ёпиш)$/i.test((b.innerText || '').trim()));
              if (no) { no.click(); return true; }
              return false;
            }""")
            if not clicked:
                pg.keyboard.press("Escape")
            pg.wait_for_timeout(500)
            continue
        pg.evaluate("""() => { const e = document.getElementById('pa-error-modal');
            const b = e.querySelector('button'); if (b) b.click(); }""")
        pg.wait_for_timeout(500)


def name_search_prompt(pg):
    """«Фамилия бўйича қидирмоқчимисиз?» oynasi ko'rinib turibdimi."""
    try:
        return bool(pg.evaluate("""() => {
          const vis = e => e && e.offsetWidth > 0;
          return [...document.querySelectorAll('.vm--modal,.modal,[role=dialog]')]
            .some(e => vis(e) && /Фамилия|қидирмоқчи|қидирмокчи/i.test(e.innerText || ''));
        }"""))
    except Exception:
        return False


def answer_no(pg):
    """O'sha oynada «Йўқ / Нет» ni bosadi."""
    try:
        pg.evaluate("""() => {
          const vis = e => e && e.offsetWidth > 0;
          const m = [...document.querySelectorAll('.vm--modal,.modal,[role=dialog]')]
            .filter(e => vis(e) && /Фамилия|қидирмоқчи|қидирмокчи/i.test(e.innerText || ''))[0];
          if (!m) return;
          const bs = [...m.querySelectorAll('button')].filter(vis);
          const no = bs.find(b => /^(нет|йўқ|йук|no|cancel|бекор|ёпиш)$/i.test((b.innerText || '').trim()));
          if (no) no.click();
        }""")
    except Exception:
        pass
    pg.wait_for_timeout(400)


def error_text(pg):
    return pg.evaluate("""() => { const e = document.getElementById('pa-error-modal');
        return (e && e.offsetWidth) ? e.innerText.replace(/\\s+/g,' ').trim() : ''; }""")


def rate_wait(pg):
    """«So'rovlar ko'payib ketdi. N soniya kuting» -> N soniya."""
    try:
        return int(__import__("mahalla_bot").rate_limit(pg) or 0)
    except Exception:
        return 0


def close_rate(pg):
    try:
        __import__("mahalla_bot").close_rate(pg)
    except Exception:
        pg.keyboard.press("Escape")


def net_error(pg):
    try:
        return bool(__import__("mahalla_bot").network_error(pg))
    except Exception:
        return False


def close_net(pg):
    try:
        __import__("mahalla_bot").close_net(pg)
    except Exception:
        pass


def close_modals(pg):
    for _ in range(4):
        pg.keyboard.press("Escape")
        pg.wait_for_timeout(300)
    close_error(pg)


def open_add(pg):
    close_error(pg)
    if modal_open(pg):          # eski/ochiq forma qolgan bo'lsa
        close_modals(pg)
        pg.wait_for_timeout(300)
    pg.evaluate("""() => {
      const b = [...document.querySelectorAll('button')].find(x => x.offsetWidth>0 && x.innerText.trim() === 'Қўшиш');
      if (b) b.click();
    }""")
    for _ in range(25):
        pg.wait_for_timeout(400)
        if modal_open(pg):
            pg.wait_for_timeout(400)
            return True
    return False


def do_search(pg, m, btn, max_rate=5, max_net=3):
    """Қидириш bosadi; rate-limit / network xatolarini qayta uradi."""
    rate_t = net_t = 0
    while True:
        for _ in range(30):
            if not btn.is_disabled():
                break
            pg.wait_for_timeout(100)
        # overlay (col-md-2) klikni to'sadi -> JS click
        try:
            btn.evaluate("b => b.click()")
        except Exception:
            btn.click(force=True)

        end = time.time() + 30
        fio = ""
        outcome, data = "timeout", ""
        while time.time() < end:
            pg.wait_for_timeout(300)
            w = rate_wait(pg)
            if w:
                close_rate(pg)
                outcome, data = "rate", w
                break
            if net_error(pg):
                close_net(pg)
                outcome, data = "net", 0
                break
            if name_search_prompt(pg):
                answer_no(pg)
                outcome, data = "err", "Ma'lumot topilmadi"
                break
            txt = error_text(pg)
            if txt:
                close_error(pg)
                outcome, data = "err", txt
                break
            try:
                fio = inp(m, "full_name").input_value().strip()
            except Exception:
                fio = ""
                if not modal(pg).count():
                    outcome, data = "closed", ""
                    break
            if fio:
                outcome, data = "ok", fio
                break

        if outcome == "ok":
            return True, data
        if outcome == "rate":
            rate_t += 1
            if rate_t > max_rate:
                return False, "So'rovlar ko'paydi — urinishlar tugadi"
            log(f"    ! so'rovlar ko'paydi — {data}s kutiladi ({rate_t}/{max_rate})")
            pg.wait_for_timeout(int(data) * 1000 + 500)
            continue
        if outcome == "net":
            net_t += 1
            if net_t > max_net:
                return False, "Network error — urinishlar tugadi"
            log(f"    ! Network error — 5s, qayta ({net_t}/{max_net})")
            pg.wait_for_timeout(5000)
            continue
        if outcome == "err":
            return False, data
        return False, "Ma'lumot topilmadi"


def person_search(pg, m, p, is_cert):
    """Qidiruvni bajaradi. Guvohnoma o'tmasa -> pasport + JSHSHIR."""
    birth = p["birth"].strftime("%d.%m.%Y") if p["birth"] else ""

    def doc_search(p_series, p_number):
        set_value(pg, "doc_serial", p_series)
        set_value(pg, "doc_number", p_number)
        if birth:
            set_value(pg, "birth_date", birth)
        return do_search(pg, m, m.locator(SEARCH_DOC))

    def pinfl_search():
        cb = m.locator("input[type=checkbox]").first
        if not cb.is_checked():
            lbl = m.locator("label", has_text="ЖШШИР билан қидириш")
            (lbl.first.click() if lbl.count() else cb.check(force=True))
            pg.wait_for_timeout(500)
        if birth:
            set_value(pg, "birth_date", birth)
        got = set_value(pg, "pinfl", p["pinfl"])
        if not got or re.sub(r"\D", "", got) != p["pinfl"]:
            pg.wait_for_timeout(400)
            set_value(pg, "pinfl", p["pinfl"])
        return do_search(pg, m, m.locator(SEARCH_PINFL))

    use_doc = is_cert or (not p["pinfl"] and p["series"] and p["number"])
    if use_doc:
        ok, res = doc_search(p["series"], p["number"])
        if not ok and is_cert and p["pinfl"]:
            # guvohnoma qabul qilinmadi -> pasport + JSHSHIR bilan qayta
            log(f"      ! guvohnoma qidiruvi o'tmadi ({res}) — pasport+JSHSHIR")
            close_modals(pg)
            if not open_add(pg):
                return False, "Forma ochilmadi (qayta)", ""
            m2 = modal(pg)
            try:
                pick(m2, pg, "relationship", RELATION)
                pick(m2, pg, "document_type", DOC_PASS)
            except Exception as e:
                close_modals(pg)
                return False, f"Ҳужжат тури танланмади: {e}", ""
            m = m2
            ok, res = pinfl_search()
        if not ok:
            return False, res, ""
        return True, res, m

    if not p["pinfl"]:
        return False, "JSHSHIR yo'q (guvohnoma ham emas)", ""
    ok, res = pinfl_search()
    if not ok and p["series"] and p["number"] and birth:
        # JSHSHIR qidiruvi o'tmadi -> hujjat seriyasi/raqami bilan qayta
        log(f"      ! JSHSHIR qidiruvi o'tmadi ({res}) — hujjat bo'yicha qayta")
        try:
            cb = m.locator("input[type=checkbox]").first
            if cb.is_checked():
                lbl = m.locator("label", has_text="ЖШШИР билан қидириш")
                (lbl.first.click() if lbl.count() else cb.uncheck(force=True))
                pg.wait_for_timeout(500)
        except Exception:
            pass
        ok, res = doc_search(p["series"], p["number"])
    return (True, res, m) if ok else (False, res, m)


def add_person(pg, p, existing=None, relation=RELATION):
    """Bitta odamni qo'shadi. (ok, reason, fio_site, series, number)"""
    if not p["birth"]:
        return False, "Tug'ilgan sana bazada yo'q", "", "", ""
    if not open_add(pg):
        return False, "Forma ochilmadi", "", "", ""
    m = modal(pg)

    try:
        pick(m, pg, "relationship", relation)
    except Exception as e:
        close_modals(pg)
        return False, f"Қариндошлик танланмади: {e}", "", "", ""

    is_cert = bool(p["doc_type"] == "cert" and p["series"] and p["number"])
    try:
        pick(m, pg, "document_type", DOC_CERT if is_cert else DOC_PASS)
    except Exception as e:
        close_modals(pg)
        return False, f"Ҳужжат тури танланмади: {e}", "", "", ""

    ok, res, m = person_search(pg, m, p, is_cert)
    if not ok:
        close_modals(pg)
        return False, res, "", "", ""

    fio_site = res
    series = inp(m, "doc_serial").input_value().strip()
    number = inp(m, "doc_number").input_value().strip()

    # Qidiruv topgan JSHSHIR saytdagi a'zolarda bor bo'lsa -> dublikat emas
    try:
        found_pinfl = re.sub(r"\D", "", inp(m, "pinfl").input_value() or "")
    except Exception:
        found_pinfl = ""
    if found_pinfl and existing and found_pinfl in existing:
        close_modals(pg)
        return True, "MAVJUD", fio_site, series, number

    # telefon: bo'sh bo'lsa ham, 17+ uchun majburiy — har doim formatlab yozamiz
    try:
        ph = inp(m, "phone")
        if not ph.input_value().strip() or len(re.sub(r"\D", "", ph.input_value())) < 12:
            set_phone(pg, m, p["phone"])
    except Exception:
        pass

    # ma'lumoti (majburiy)
    try:
        pick(m, pg, "study_level_id", EDUCATION)
    except Exception as e:
        log(f"      ! Маълумоти танланмади: {e}")

    save_resp = {}

    def on_save_resp(r):
        try:
            if "/forms/survey_homes_family" in r.url and r.request.method in ("POST", "PUT", "PATCH"):
                save_resp["status"] = r.status
                try:
                    save_resp["body"] = r.text()
                except Exception:
                    save_resp["body"] = ""
        except Exception:
            pass

    pg.on("response", on_save_resp)

    def click_save():
        save_resp.clear()
        pg.evaluate("""() => {
          const m = [...document.querySelectorAll('.vm--modal.form-modal')].filter(x => x.offsetWidth).pop();
          if (!m) return;
          const b = [...m.querySelectorAll('button')].find(x => x.innerText.trim() === 'Сақлаш');
          if (b) b.click();
        }""")

    try:
        click_save()
        end = time.time() + 45
        fixes = 0
        last_err = {}
        while time.time() < end:
            pg.wait_for_timeout(600)

            st = save_resp.get("status")
            if st is not None:
                if st >= 400:
                    reason = human_reason(save_error_text(save_resp))
                    log(f"      ! sayt javobi {st}: {reason}")
                    close_modals(pg)
                    return False, reason, fio_site, series, number
                for _ in range(12):          # 2xx -> modal yopilishini kut
                    if not modal_open(pg):
                        break
                    pg.wait_for_timeout(300)
                return True, "OK", fio_site, series, number

            txt = error_text(pg)
            if txt:
                close_error(pg)
                close_modals(pg)
                return False, txt, fio_site, series, number
            if not modal_open(pg):
                return True, "OK", fio_site, series, number

            errs = validation_errors(pg)
            if errs:
                last_err = errs
                if fixes < 2:
                    fixes += 1
                    names = ", ".join(errs)
                    log(f"      ! validatsiya: {names} — tuzatilmoqda ({fixes}/2)")
                    fixed = False
                    for name in errs:
                        fixed = fix_field(pg, m, name, p) or fixed
                    if fixed:
                        pg.wait_for_timeout(500)
                        click_save()
                        end = time.time() + 30
                        continue
                close_modals(pg)
                return (False, human_reason(" ".join(last_err.values())),
                        fio_site, series, number)

        close_modals(pg)
        return False, "Saqlanmadi (javob kelmadi)", fio_site, series, number
    finally:
        try:
            pg.remove_listener("response", on_save_resp)
        except Exception:
            pass



# ---------------------------------------------------------------- xonadon
def site_has(members, p):
    """Odam saytdagi a'zolar orasida bormi — faqat JSHSHIR bo'yicha.

    F.I.Sh. va tug'ilgan sana ishlatilmaydi (lotin/kirill farqi, xatolar).
    JSHSHIR yo'q odamlar uchun qo'shish formasidagi qidiruvdan keyin
    chiqqan JSHSHIR solishtiriladi (add_person -> existing).
    """
    if not p["pinfl"]:
        return False
    for x in members:
        if re.sub(r"\D", "", x["pinfl"] or "") == p["pinfl"]:
            return True
    return False


def process_house(pg, h, by_pinfl, by_name, member_houses, stats):
    cad, owner_fio = h["cad"], h["fio"]
    owner_pinfl = re.sub(r"\D", "", h["pinfl"] or "")

    block = by_pinfl.get(owner_pinfl) if owner_pinfl else None
    found_as_member = bool(block is not None and owner_pinfl in member_houses
                           and member_houses[owner_pinfl] is block)
    if block is None:
        block = by_name.get(norm(owner_fio))
    if block is None:
        log(f"  XATO -> xonadon bazada topilmadi: {owner_fio} ({cad})")
        rl.log_fail(owner_fio, "", "", owner_pinfl, cad, "Bazada topilmadi (egasi)",
                    street=h["street_name"])
        stats["fail"] += 1
        return

    members = read_members(pg)
    if members is None or not members:
        # a'zolar o'qilmasa qo'shishga urinmaymiz — dublikat bo'lib ketadi
        log("  XATO -> saytdagi a'zolar o'qilmadi — xonadon o'tkazildi")
        rl.log_fail(owner_fio, "", "", owner_pinfl, cad, "Sayt a'zolari o'qilmadi",
                    street=h["street_name"])
        stats["fail"] += 1
        return

    existing = {re.sub(r"\D", "", x["pinfl"] or "") for x in members if x["pinfl"]}
    # sayt egasi bazada oddiy a'zo bo'lib chiqsa — bazadagi egasi ham nomzod
    candidates = list(block["members"])
    if found_as_member:
        candidates = [block["owner"]] + candidates
    missing = [p for p in candidates if not site_has(members, p)]

    if not missing:
        log(f"  = {cad} {owner_fio}: hamma a'zo joyida ({len(members)})")
        return

    log(f"  > {cad} {owner_fio}: {len(missing)} ta qo'shiladi "
        f"(saytda {len(members)}, bazada {1 + len(block['members'])})")

    added = []
    for p in missing:
        if stats["total"] >= stats["limit"]:
            return
        kind = "guvohnoma" if p["doc_type"] == "cert" else "pasport"
        log(f"    ? {p['fio']} ({p['pinfl'] or p['series'] + ' ' + p['number']}) [{kind}]")
        if DRY:
            log("      (dry-run)")
            stats["total"] += 1
            continue
        try:
            ok, reason, fio_site, series, number = add_person(pg, p, existing=existing)
        except Exception as e:
            ok, reason, fio_site, series, number = False, f"Exception: {e}", "", "", ""
            traceback.print_exc()

        if ok and reason == "MAVJUD":
            log("      = saytda mavjud (JSHSHIR bo'yicha) — o'tkazildi")
            continue

        # boshqa xonadonda ro'yxatda bo'lsa — limitni egallamaydi
        skip = (not ok and reason.startswith("Boshqa xonadonda mavjud"))
        if not skip:
            stats["total"] += 1
        if ok:
            added.append(p)
            if p["pinfl"]:
                existing.add(p["pinfl"])
            log(f"      OK -> {fio_site}")
            rl.log_success(p["fio"], series, number, p["pinfl"], cad,
                           street=h["street_name"], phone=p["phone"],
                           doc_type=kind)
            stats["ok"] += 1
        else:
            log(f"      XATO -> {reason}")
            rl.log_fail(p["fio"], p["series"], p["number"], p["pinfl"], cad, reason,
                        street=h["street_name"], doc_type=kind)
            stats["fail"] += 1
        rl.flush()

    if added:
        now = read_members(pg)
        for p in added:
            if now and not site_has(now, p):
                log(f"      ! tekshiruv: {p['fio']} saytda ko'rinmadi")


def street_label(street_id):
    """streets.json dan ko'cha nomi (kirill)."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "streets.json")
    try:
        with open(path, encoding="utf-8") as f:
            for s in json.load(f):
                if int(s["id"]) == int(street_id):
                    return s["name"].replace("ул. ", "").strip()
    except Exception:
        pass
    return ""


def all_index(houses):
    """Barcha odamlar bo'yicha indeks (egasi + a'zolar)."""
    by_pinfl, by_name = base_new.index(houses)
    by_pinfl = dict(by_pinfl)
    member_house = {}
    for h in houses:
        for p in h["members"]:
            if p["pinfl"]:
                by_pinfl.setdefault(p["pinfl"], h)
                member_house.setdefault(p["pinfl"], h)
    return by_pinfl, by_name, member_house


def main():
    global LOG_FILE, DRY

    ap = argparse.ArgumentParser()
    ap.add_argument("--street", required=True, help="street_id yoki nom")
    ap.add_argument("--slot", type=int, default=1)
    ap.add_argument("--limit", type=int, default=0, help="nechta odam (0 = hammasi)")
    ap.add_argument("--cad", default="", help="faqat shu kadastr (test uchun)")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--base", default=base_new.BASE_PATH)
    ap.add_argument("--cdp", default="http://127.0.0.1:9222")
    args = ap.parse_args()

    DRY = args.dry_run
    LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            f"run_new_{args.slot}.log")

    # ko'cha id/ nom
    houses, dead = base_new.load(args.base)
    by_pinfl, by_name, member_houses = all_index(houses)

    street_id, street_name = None, None
    q = str(args.street).strip()
    if q.isdigit():
        street_id = int(q)
        street_name = street_label(street_id)          # streets.json -> kirill
        if not street_name:
            street_name = q
    else:
        for h in houses:
            if q.lower() in (h["street"] or "").lower():
                street_name = h["street"]
                break

    log("=" * 60)
    log(f"OYNA (slot): {args.slot}")
    log(f"BAZA: {len(houses)} xonadon, o'lgan tashlandi: {dead}")
    log(f"KO'CHA: {args.street} -> {street_name or '?'}")
    log(f"LIMIT: {args.limit or 'hammasi'} | dry-run: {DRY}")

    rep_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "hisobot",
                           f"ул_{street_name or args.street}")
    rl.set_dir(rep_dir)
    log(f"hisobot: {rep_dir}")

    stats = {"ok": 0, "fail": 0, "total": 0, "limit": args.limit or 10 ** 9}

    with sync_playwright() as p:
        b = p.chromium.connect_over_cdp(args.cdp)
        ctx = b.contexts[0]
        anchor = next((x for x in ctx.pages if "online-mahalla.uz" in x.url), ctx.pages[0])
        pages = ensure_pages(ctx, args.slot, anchor)
        pg = pages[args.slot - 1]
        pg.bring_to_front()
        install_page_fixes(pg)

        if not street_id:
            log("street_id kerak (masalan --street 350901168)")
            return

        n = open_street(pg, street_id)
        if n == -1:
            log("!!! LOGIN KERAK — Chrome oynasida saytga kiring, keyin qayta ishga tushiring.")
            log("    Hech qanday sahifa yangilanmaydi, xotirjam kiring.")
            return
        log(f"ko'chada {n} xonadon topildi")
        if not n:
            return

        house_list = read_houses(pg)
        log(f"jadval: {len(house_list)} qator")

        page = 1
        done_cads = set()
        while stats["total"] < stats["limit"]:
            if page > 1:
                if not go_page(pg, street_id, page):
                    log(f"{page}-sahifa yo'q — tugadi")
                    break
            house_list = [h for h in read_houses(pg)
                          if h["cad"] not in done_cads
                          and (not args.cad or h["cad"] == args.cad)]
            log(f"--- {page}-sahifa: {len(house_list)} xonadon ---")
            if not house_list:
                break

            for i, h in enumerate(house_list, 1):
                if stats["total"] >= stats["limit"]:
                    break
                done_cads.add(h["cad"])
                h["street_name"] = street_name or ""
                log(f"[{page}.{i}/{len(house_list)}] {h['cad']} — {h['fio']}")
                try:
                    if not open_house(pg, h["href"]):
                        log("  XATO -> xonadon ochilmadi")
                        stats["fail"] += 1
                        continue
                    process_house(pg, h, by_pinfl, by_name, member_houses, stats)
                except Exception as e:
                    log(f"  XATO -> {e}")
                    traceback.print_exc()
                    stats["fail"] += 1
                rl.flush()
                pg.wait_for_timeout(500)
            page += 1

    rl.flush()
    log("-" * 60)
    log(f"TUGADI: qo'shildi={stats['ok']}, xato={stats['fail']}, jami={stats['total']}")
    log("=" * 60)


if __name__ == "__main__":
    main()

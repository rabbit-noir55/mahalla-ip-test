"""Lolazor MFY — YANGI BOT (v2).

Bu chatda o'rgatilgan mantiq + eski mahalla_bot.py dagi mustahkam qismlar.

Asosiy qoidalar:
  * Saytdagi a'zolar JSHSHIR bo'yicha solishtiriladi.
  * Egasi JSHSHIRi bo'yicha bazadagi xonadon bloki topiladi.
  * Guvohnoma: 2010+ -> I-QQ, II-QQ, III-QQ, IV-QQ, V-QQ (T bo'lsa faqat T).
    Raqam o'zgarmaydi, seriya ishonchsiz.
  * Guvohnoma 2010 dan oldin -> sayt qabul qilmaydi, o'tkaziladi.
  * Tug'ilgan sana yo'q bo'lsa -> IIB bazasidan (faqat sana) olinadi.
  * Pasport: JSHSHIR bo'lsa JSHSHIR bilan, bo'lmasa seriya+raqam+sana.
  * Saqlash 400 -> «Текшириш»; «Муваффақиятли сақланди» = qo'shildi.
  * «Фамилия бўйича қидирмоқчимисиз?» -> «Йўқ».
  * Kutilmagan holat -> to'xtaydi (--continue-on-new bilan davom etadi).

Ishga tushirish:
  python mahalla_v2.py --street 350901168 --slot 1
  python mahalla_v2.py --street 350901168 --slot 1 --limit 20 --houses 5
"""
import argparse
import json
import os
import random
import re
import sys
import time
import traceback
from datetime import date, datetime

from playwright.sync_api import sync_playwright

import base_new
import report_log as rl
import mahalla_new as m

SITE = "https://online-mahalla.uz"
Q = "obl_id=35&area_id=3509&district_id=3509040"
MID = "&_mid=3548"

BASE = base_new.BASE_PATH
IIB = os.environ.get("MAHALLA_IIB",
                     r"C:\Users\Admin\Desktop\Беруний Тумани ИИБ база.xlsx")
HERE = os.path.dirname(os.path.abspath(__file__))

CERT_SERIES = ["I-QQ", "II-QQ", "III-QQ", "IV-QQ", "V-QQ"]
PHONE_PREFIX = "97"

LOG_FILE = None
DRY = False
STRICT = True


# ---------------------------------------------------------------- log
def log(msg=""):
    line = msg if not msg else f"[{datetime.now():%H:%M:%S}] {msg}"
    try:
        print(line, flush=True)
    except Exception:
        try:
            sys.stdout.write(line.encode("ascii", "replace").decode("ascii") + "\n")
        except Exception:
            pass
    if LOG_FILE:
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except OSError:
            pass


def digits(s):
    return re.sub(r"\D", "", s or "")


def rand_phone():
    d = "998%s%03d%02d%02d" % (PHONE_PREFIX, random.randint(100, 999),
                               random.randint(10, 99), random.randint(10, 99))
    return "+998(%s)%s-%s-%s" % (d[3:5], d[5:8], d[8:10], d[10:12])


def is_known_notfound(res):
    t = (res or "").upper()
    return ("ТОПИЛМАДИ" in t or "TOPILMADI" in t or "ТУҒИЛГАН САНА" in t
            or "МАЙДОНЛАР ТЎҒРИ" in t)


def norm_name(s):
    return base_new.norm_name(s)


_TR = str.maketrans({
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Е": "E", "Ё": "YO",
    "Ж": "J", "З": "Z", "И": "I", "Й": "Y", "К": "K", "Л": "L", "М": "M",
    "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T", "У": "U",
    "Ф": "F", "Х": "X", "Ц": "TS", "Ч": "CH", "Ш": "SH", "Щ": "SH",
    "Ъ": "", "Ь": "", "Ы": "I", "Э": "E", "Ю": "YU", "Я": "YA",
    "Ў": "O", "Қ": "Q", "Ғ": "G", "Ҳ": "H",
})


def latin(s):
    return re.sub(r"[^A-Z ]", "", (s or "").upper().translate(_TR))


def name_like(a, b):
    from difflib import SequenceMatcher
    ta, tb = latin(a).split(), latin(b).split()
    if not ta or not tb:
        return False
    return SequenceMatcher(None, ta[0], tb[0]).ratio() >= 0.72


def read_employment(pg):
    """«Бандлик ва таълим» -> «Янгилаш» -> jadvaldagi F.I.Sh. lar."""
    try:
        pg.evaluate("""() => {
          const t = [...document.querySelectorAll('[role=tab]')]
            .find(x => (x.innerText||'').includes('Бандлик ва таълим'));
          if (t) t.click();
        }""")
        pg.wait_for_timeout(500)
        clicked = pg.evaluate("""() => {
          const panes = [...document.querySelectorAll('[role=tabpanel]')].filter(x => x.offsetWidth > 0);
          for (const pane of panes) {
            const b = [...pane.querySelectorAll('button')].find(x => (x.innerText||'').trim() === 'Янгилаш');
            if (b) { b.click(); return true; }
          }
          return false;
        }""")
        log(f"    Бандлик: Янгилаш={clicked}")
    except Exception as e:
        log(f"    ! Бандлик: {e}")
        return []
    for _ in range(6):
        pg.wait_for_timeout(600)
        names = pg.evaluate("""() => {
          const pane = [...document.querySelectorAll('[role=tabpanel]')].find(x => x.offsetWidth > 0);
          if (!pane) return [];
          const t = pane.querySelector('table');
          if (!t) return [];
          return [...t.querySelectorAll('tbody tr')].map(tr => {
            const td = [...tr.querySelectorAll('td')];
            return td.length >= 2 ? (td[1].innerText||'').trim().replace(/\\s+/g,' ') : '';
          }).filter(Boolean);
        }""")
        if names:
            return names
    return []


def verify_employment(pg, added):
    names = read_employment(pg)
    if not names:
        log("    ! Бандлик jadvali o'qilmadi")
        return
    log(f"    Бандлик ro'yxati: {len(names)}")
    for p in added:
        if any(name_like(p["fio"], n) for n in names):
            log(f"    OK -> bandlikda bor: {p['fio']}")
        else:
            log(f"    ! bandlikda ko'rinmadi: {p['fio']}")


# ---------------------------------------------------------------- ma'lumot
def load_base(path):
    houses, dead = base_new.load(path)
    by_pinfl, by_name = base_new.index(houses)
    by_pinfl = dict(by_pinfl)
    member_house = {}
    for h in houses:
        for p in h["members"]:
            if p["pinfl"]:
                by_pinfl.setdefault(p["pinfl"], h)
                member_house.setdefault(p["pinfl"], h)
    return houses, by_pinfl, by_name, member_house, dead


def load_iib_dates(path):
    """IIB (eski) bazadan faqat tug'ilgan sanani oladi: JSHSHIR -> sana."""
    from openpyxl import load_workbook
    if not os.path.exists(path):
        log(f"IIB baza yo'q: {path}")
        return {}
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Лист1"]
    out = {}
    for r in ws.iter_rows(values_only=True):
        pin = digits(str(r[0] or ""))
        if len(pin) != 14:
            continue
        b = r[3]
        d = None
        if isinstance(b, datetime):
            d = b.date()
        elif b:
            for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                try:
                    d = datetime.strptime(str(b).strip(), fmt).date()
                    break
                except ValueError:
                    pass
        if d:
            out[pin] = d
    wb.close()
    return out


def load_streets():
    with open(os.path.join(HERE, "streets.json"), encoding="utf-8") as f:
        return json.load(f)


def street_info(q):
    q = str(q).strip()
    streets = load_streets()
    if q.isdigit():
        for s in streets:
            if str(s["id"]) == q:
                return int(s["id"]), s["name"].replace("ул. ", "").strip()
        return int(q), q
    low = q.lower()
    for s in streets:
        if low == s["name"].lower() or low in s["name"].lower():
            return int(s["id"]), s["name"].replace("ул. ", "").strip()
    raise SystemExit(f"Ko'cha topilmadi: {q!r}")


def state_path(street_id):
    return os.path.join(HERE, f"state_v2_{street_id}.json")


def load_state(street_id):
    p = state_path(street_id)
    if os.path.exists(p):
        try:
            return json.load(open(p, encoding="utf-8"))
        except Exception:
            pass
    return {"processed": [], "added": [], "last": 0}


def save_state(street_id, st):
    with open(state_path(street_id), "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1)


# ---------------------------------------------------------------- ko'cha
def street_url(sid):
    return f"{SITE}/tables/survey_homes_street?{Q}&street_id={sid}{MID}"


def read_all_houses(pg, sid):
    """Barcha xonadonlar: _page=N (eski botdagidek), sekin tarmoq uchun 3 urinish."""
    rows, seen = [], set()
    for page in range(1, 21):
        url = street_url(sid) + (f"&_page={page}" if page > 1 else "")
        n = 0
        for attempt in range(3):
            try:
                pg.goto(url, wait_until="commit", timeout=90000)
            except Exception as e:
                log(f"  goto xatosi: {e}")
            n = 0
            for _ in range(120):                 # 60 s
                pg.wait_for_timeout(500)
                if "/login" in pg.url:
                    return -1
                n = pg.evaluate("""() => {
                  const t = [...document.querySelectorAll('table')].find(x => (x.innerText||'').includes('кадастр'));
                  return t ? t.querySelectorAll('tbody tr').length : 0;
                }""")
                if n:
                    break
            if n:
                break
            log(f"  {page}-sahifa yuklanmadi ({attempt + 1}/3) — qayta")
            pg.wait_for_timeout(3000)
        if not n:
            if page == 1:
                return []
            break
        batch = m.read_houses(pg)
        new = 0
        for h in batch:
            key = h["cad"] or h["href"]
            if not key or key in seen:
                continue
            seen.add(key)
            rows.append(h)
            new += 1
        if not new or len(batch) < 100:
            break
    return rows


# ---------------------------------------------------------------- qidiruv
def set_checkbox(pg, modal, on):
    cb = modal.locator("input[type=checkbox]").first
    if cb.is_checked() != on:
        cb.evaluate("c => c.click()")
        pg.wait_for_timeout(400)
    return cb.is_checked()


def search_person(pg, modal, p):
    """(ok, res, modal, reg) — reg = reyestrdagi haqiqiy ma'lumot."""
    reg = {}

    def on_reg(r):
        try:
            if "/fhdyo/" in r.url:
                j = r.json()
                d = (j.get("data") or {}).get("info", {}).get("data") or {}
                if isinstance(d, dict) and d.get("birth_date"):
                    reg.update(d)
        except Exception:
            pass

    pg.on("response", on_reg)

    def done(ok, res):
        try:
            pg.remove_listener("response", on_reg)
        except Exception:
            pass
        return ok, res, modal, reg

    birth = p["birth"].strftime("%d.%m.%Y") if p["birth"] else ""
    is_cert = bool(p["doc_type"] == "cert" and p["series"] and p["number"])

    if is_cert:
        set_checkbox(pg, modal, False)
        ser0 = (p["series"] or "").upper()
        variants = ["T"] if ser0.startswith("T") else CERT_SERIES
        last = ""
        for ser in variants:
            m.set_value(pg, "birth_date", birth)
            m.set_value(pg, "doc_serial", ser)
            m.set_value(pg, "doc_number", p["number"])
            m.close_error(pg)
            ok, res = m.do_search(pg, modal, modal.locator(m.SEARCH_DOC))
            log(f"      seriya {ser}: {ok} | {str(res)[:90]}")
            if ok:
                return done(True, res)
            last = res
            if not is_known_notfound(res):
                return done(False, res)
        return done(False, last)

    if p["pinfl"]:
        set_checkbox(pg, modal, True)
        m.set_value(pg, "birth_date", birth)
        m.set_value(pg, "pinfl", p["pinfl"])
        m.close_error(pg)
        ok, res = m.do_search(pg, modal, modal.locator(m.SEARCH_PINFL))
        log(f"      JSHSHIR: {ok} | {str(res)[:90]}")
        return done(ok, res)

    if not (p["series"] and p["number"]):
        return done(False, "Seriya/raqam yo'q (JSHSHIR ham yo'q)")
    set_checkbox(pg, modal, False)
    m.set_value(pg, "birth_date", birth)
    m.set_value(pg, "doc_serial", p["series"])
    m.set_value(pg, "doc_number", p["number"])
    m.close_error(pg)
    ok, res = m.do_search(pg, modal, modal.locator(m.SEARCH_DOC))
    log(f"      pasport: {ok} | {str(res)[:90]}")
    return done(ok, res)


# ---------------------------------------------------------------- saqlash
def click_save(pg):
    pg.evaluate("""() => {
      const mm = [...document.querySelectorAll('.vm--modal.form-modal')].filter(x=>x.offsetWidth).pop();
      if (!mm) return;
      const b = [...mm.querySelectorAll('button')].find(x => x.innerText.trim() === 'Сақлаш');
      if (b) b.click();
    }""")


def click_check(pg):
    return pg.evaluate("""() => {
      const mm = [...document.querySelectorAll('.vm--modal.form-modal')].filter(x=>x.offsetWidth).pop();
      if (!mm) return false;
      const b = [...mm.querySelectorAll('button.error-action-btn')]
        .find(x => (x.innerText||'').trim() === 'Текшириш');
      if (!b) return false;
      b.click(); return true;
    }""")


def save_person(pg, modal, p, site_pins):
    """('added'|'already'|'exists'|'mismatch'|'notfound'|'skip'|'new', izoh, phone)"""
    try:
        found_pin = digits(m.inp(modal, "pinfl").input_value())
    except Exception:
        found_pin = ""
    if found_pin and found_pin in site_pins:
        m.close_modals(pg)
        return "already", found_pin, ""

    # qidiruv F.I.Sh.ni to'ldirmagan bo'lsa — saqlashga urinmaymiz
    try:
        fio_now = m.inp(modal, "full_name").input_value().strip()
    except Exception:
        fio_now = ""
    if not fio_now:
        m.close_modals(pg)
        return "notfound", "Qidiruv F.I.Sh.ni to'ldirmadi", ""

    phone = rand_phone()
    try:
        m.pick(modal, pg, "study_level_id", m.EDUCATION)
        m.set_phone(pg, modal, phone)
    except Exception as e:
        m.close_modals(pg)
        return "skip", f"Маълумоти/телефон: {e}", phone

    resp = {}

    def on_resp(r):
        try:
            if "/forms/survey_homes_family" in r.url and r.request.method in ("POST", "PUT", "PATCH"):
                resp["status"] = r.status
                try:
                    resp["body"] = r.text()
                except Exception:
                    resp["body"] = ""
        except Exception:
            pass

    pg.on("response", on_resp)
    fixes = 0
    try:
        click_save(pg)
        end = time.time() + 45
        while time.time() < end:
            pg.wait_for_timeout(600)
            st = resp.get("status")
            if st is not None:
                break
            if m.net_error(pg):
                m.close_net(pg)
                m.close_modals(pg)
                return "notfound", "Network error", phone
            txt = m.error_text(pg)
            if txt:
                m.close_error(pg)
                m.close_modals(pg)
                low = txt.lower()
                if "сақланди" in low and "муваффақиятли" in low:
                    return "added", txt, phone
                if "5 ёки ундан кўп" in txt:
                    return "phone5", txt, phone
                return "notsaved", txt, phone
            if not m.modal_open(pg):
                return "added", "modal yopildi", phone
            # faqat haqiqiy validatsiya (тўлдириш керак) — cheklangan tuzatish
            errs = m.validation_errors(pg)
            if errs and any("тўлдириш керак" in v.lower() for v in errs.values()):
                if fixes >= 2:
                    m.close_modals(pg)
                    return "new", "Validatsiya takrorlanmoqda: " + ", ".join(errs), phone
                fixes += 1
                for name in errs:
                    m.fix_field(pg, modal, name, p)
                pg.wait_for_timeout(400)
                click_save(pg)
                end = time.time() + 25
                continue
        st = resp.get("status")
        if st and 200 <= st < 300:
            for _ in range(12):
                if not m.modal_open(pg):
                    break
                pg.wait_for_timeout(300)
            return "added", resp.get("body", "")[:100], phone
        if st and st >= 400:
            errs = {}
            try:
                errs = (json.loads(resp.get("body") or "{}").get("errors") or {})
            except Exception:
                errs = {}
            phone_err = str(errs.get("phone") or "")
            pinfl_err = str(errs.get("pinfl") or "")
            if "5 ёки ундан кўп" in phone_err:
                m.close_modals(pg)
                return "phone5", phone_err, phone
            if "мувофиқ" in pinfl_err.lower() or "мос келмади" in pinfl_err.lower():
                m.close_modals(pg)
                return "mismatch_reg", pinfl_err, phone

            pg.wait_for_timeout(800)
            clicked = click_check(pg)
            pg.wait_for_timeout(3000)
            err = pg.evaluate("""() => {
              const e = document.getElementById('pa-error-modal');
              return e && e.offsetWidth ? e.innerText.replace(/\\s+/g,' ').trim() : '';
            }""")
            m.close_error(pg)
            m.close_modals(pg)
            if clicked and "МУВАФФАҚИЯТЛИ" in (err or "").upper():
                return "added", err, phone
            if clicked and "Текширувдан ўтмади" in err:
                return "exists", err, phone
            reason = m.human_reason(m.save_error_text(resp))
            body = (resp.get("body") or "").lower()
            low_pinfl = pinfl_err.lower()
            if ("мавжуд" in low_pinfl or "қайта киритиб" in low_pinfl
                    or "мавжуд" in body or "қайта киритиб" in body):
                return "exists", (err or pinfl_err or reason), phone
            if "мос келмади" in reason.lower() or "мувофиқ" in body or "мувофик" in body:
                return "mismatch", reason, phone
            return "new", (err or reason), phone
        m.close_modals(pg)
        return "new", "saqlash javobi kelmadi", phone
    finally:
        try:
            pg.remove_listener("response", on_resp)
        except Exception:
            pass


# ---------------------------------------------------------------- xonadon
def add_candidate(pg, p, site_pins, street_name, cad, retry=0, phone_try=0):
    """Bitta nomzodni qo'shadi. ('added'|'already'|'exists'|'mismatch'|'notfound'|'skip'|'new', ...)"""
    if not p["birth"]:
        return "skip", "Tug'ilgan sana yo'q (IIB da ham topilmadi)"
    is_cert = bool(p["doc_type"] == "cert" and p["series"] and p["number"])
    if is_cert and p["birth"].year < 2010:
        return "skip", "Guvohnoma 2010 dan oldin — sayt qabul qilmaydi"
    if is_cert and not p["number"]:
        return "skip", "Guvohnoma raqami yo'q"
    if not p["pinfl"] and not (p["series"] and p["number"]):
        return "skip", "JSHSHIR ham, hujjat ham yo'q"

    if not m.open_add(pg):
        # tab/pahifa kechikishi bo'lishi mumkin — qayta urinamiz
        m.read_members(pg)
        m.close_modals(pg)
        if not m.open_add(pg):
            return "skip", "Qo'shish formasi ochilmadi"
    modal = m.modal(pg)
    try:
        m.pick(modal, pg, "relationship", "Бошқа")
        m.pick(modal, pg, "document_type", m.DOC_CERT if is_cert else m.DOC_PASS)
    except Exception as e:
        m.close_modals(pg)
        return "skip", f"Select2: {e}"

    ok, res, modal, reg = search_person(pg, modal, p)
    if not ok:
        m.close_modals(pg)
        if "Текширувдан ўтмади" in (res or ""):
            return "exists", res
        if is_known_notfound(res) or "network error" in (res or "").lower():
            return "notfound", res
        return "new", res

    # reyestrdagi haqiqiy ma'lumot (sana / JSHSHIR)
    reg_date = None
    if reg.get("birth_date"):
        try:
            reg_date = datetime.strptime(reg["birth_date"], "%d.%m.%Y").date()
        except ValueError:
            reg_date = None
    reg_pinfl = digits(str(reg.get("pinfl") or ""))

    if reg_date and p["birth"] and reg_date != p["birth"] and retry < 2:
        m.close_modals(pg)
        p2 = dict(p)
        p2["birth"] = reg_date
        if reg_pinfl:
            p2["pinfl"] = reg_pinfl
        if is_cert and reg_date.year >= 2010:
            p2["doc_type"] = "cert"
        else:
            p2["doc_type"] = "pass"
            p2["series"] = ""
            p2["number"] = ""
        log(f"      ! reyestr sanasi {reg_date} — qayta qidiruv")
        return add_candidate(pg, p2, site_pins, street_name, cad,
                             retry=retry + 1, phone_try=phone_try)
    if reg_date and p["birth"] and reg_date != p["birth"]:
        return "new", f"Reyestr sanasi barqarorlashmadi: {reg_date}"

    outcome, detail, phone = save_person(pg, modal, p, site_pins)

    if outcome == "phone5":
        m.close_modals(pg)
        if phone_try < 3:
            log(f"      ! telefon 5+ xonadonda — yangi raqam ({phone_try + 1}/3)")
            return add_candidate(pg, p, site_pins, street_name, cad,
                                 retry=retry, phone_try=phone_try + 1)
        return "skip", "Telefon 5+ xonadonda ishlatilgan"

    if outcome == "mismatch_reg" and retry < 1 and reg_date:
        m.close_modals(pg)
        p2 = dict(p)
        p2["birth"] = reg_date
        p2["pinfl"] = reg_pinfl or p["pinfl"]
        p2["doc_type"] = "pass"
        p2["series"] = ""
        p2["number"] = ""
        log(f"      ! mos kelmadi — reyestr sanasi {reg_date} bilan qayta")
        return add_candidate(pg, p2, site_pins, street_name, cad,
                             retry=retry + 1, phone_try=phone_try)

    if outcome == "mismatch_reg":
        outcome = "mismatch"
    return outcome, detail


def process_house(pg, h, by_pinfl, by_name, member_house, iib, site_pins_all, street_name, stats):
    """Bir xonadon. 'new' qaytarsa — to'xtash."""
    cad = h["cad"]
    if not m.open_house(pg, h["href"]):
        log("  XATO -> xonadon ochilmadi (o'tkazildi)")
        rl.log_fail(h["fio"], "", "", "", cad, "Xonadon ochilmadi", street=street_name)
        stats["fail"] += 1
        return "ok", "xonadon ochilmadi"
    members = m.read_members(pg)
    if not members:
        log("  XATO -> a'zolar o'qilmadi (o'tkazildi)")
        rl.log_fail(h["fio"], "", "", "", cad, "Sayt a'zolari o'qilmadi", street=street_name)
        stats["fail"] += 1
        return "ok", "a'zolar o'qilmadi"

    site_pins = {digits(x["pinfl"]) for x in members if x["pinfl"]}
    owner = next((x for x in members if "эгаси" in (x["rel"] or "").lower()), members[0])
    owner_pin = digits(owner["pinfl"])

    block = by_pinfl.get(owner_pin) if owner_pin else None
    found_as_member = bool(block is not None and owner_pin in member_house
                           and member_house[owner_pin] is block)
    if block is None:
        block = by_name.get(norm_name(owner["fio"]))
    if block is None:
        rl.log_fail(owner["fio"], "", "", owner_pin, cad, "Bazada topilmadi (egasi)",
                    street=street_name)
        stats["fail"] += 1
        return "ok", "baza bloki topilmadi"

    candidates = list(block["members"])
    if found_as_member:
        candidates = [block["owner"]] + candidates

    missing = [p for p in candidates if not (p["pinfl"] and p["pinfl"] in site_pins)]
    if not missing:
        log(f"  = {cad} {owner['fio']}: hamma a'zo joyida ({len(members)})")
        return "ok", ""

    log(f"  > {cad} {owner['fio']}: {len(missing)} ta nomzod "
        f"(saytda {len(members)})")

    added_here = []
    for p in missing:
        if stats["limit"] and stats["attempts"] >= stats["limit"]:
            return "ok", "limit"
        if not p["birth"] and p["pinfl"]:
            d = iib.get(p["pinfl"])
            if d:
                p["birth"] = d
                log(f"    (IIB: sana {d})")
        log(f"    ? {p['fio']} | {p['birth']} | {p['doc_raw'] or '-'} | {p['pinfl'] or '-'}")
        if DRY:
            log("      (dry-run)")
            stats["attempts"] += 1
            continue

        outcome, detail = add_candidate(pg, p, site_pins, street_name, cad)
        if outcome == "new":
            return "new", detail
        if outcome == "skip":
            log(f"      = o'tkazildi: {detail}")
            rl.log_fail(p["fio"], p["series"], p["number"], p["pinfl"], cad,
                        f"O'tkazildi: {detail}", street=street_name,
                        doc_type="guvohnoma" if p["doc_type"] == "cert" else "pasport")
            stats["skip"] += 1
            rl.flush()
            continue

        stats["attempts"] += 1
        if outcome == "added":
            if p["pinfl"]:
                site_pins.add(p["pinfl"])
                site_pins_all.add(p["pinfl"])
            added_here.append(p)
            log(f"      OK -> qo'shildi")
            rl.log_success(p["fio"], p["series"], p["number"], p["pinfl"], cad,
                           street=street_name, phone=p.get("phone", ""),
                           doc_type="guvohnoma" if p["doc_type"] == "cert" else "pasport")
            stats["ok"] += 1
        elif outcome == "already":
            log(f"      = saytda mavjud ({detail})")
            rl.log_fail(p["fio"], p["series"], p["number"], p["pinfl"], cad,
                        "Saytda mavjud", street=street_name)
            stats["skip"] += 1
        elif outcome == "exists":
            log("      = boshqa xonadonda — tekshiruvdan o'tmadi")
            rl.log_fail(p["fio"], p["series"], p["number"], p["pinfl"], cad,
                        "Boshqa xonadonda mavjud (tekshiruvdan o'tmadi)", street=street_name)
            stats["fail"] += 1
        elif outcome == "mismatch":
            log(f"      = mos kelmadi: {detail}")
            rl.log_fail(p["fio"], p["series"], p["number"], p["pinfl"], cad,
                        f"Ma'lumotlar mos kelmadi: {detail}", street=street_name)
            stats["fail"] += 1
        elif outcome == "notfound":
            log(f"      = topilmadi: {str(detail)[:80]}")
            rl.log_fail(p["fio"], p["series"], p["number"], p["pinfl"], cad,
                        f"Topilmadi: {detail}", street=street_name)
            stats["fail"] += 1
        rl.flush()

    if added_here:
        now = m.read_members(pg)
        if now is not None:
            now_pins = {digits(x["pinfl"]) for x in now if x["pinfl"]}
            for p in added_here:
                if p["pinfl"] and p["pinfl"] not in now_pins:
                    log(f"      ! tekshiruv: {p['fio']} ro'yxatda ko'rinmadi")
        verify_employment(pg, added_here)
    return "ok", ""


# ---------------------------------------------------------------- main
def main():
    global LOG_FILE, DRY, STRICT
    ap = argparse.ArgumentParser()
    ap.add_argument("--street", required=True)
    ap.add_argument("--slot", type=int, default=1)
    ap.add_argument("--limit", type=int, default=0, help="nechta urinish (0=hammasi)")
    ap.add_argument("--houses", type=int, default=0, help="nechta xonadon (0=hammasi)")
    ap.add_argument("--start", type=int, default=1, help="nechanchi xonadondan")
    ap.add_argument("--restart", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--cdp", default="http://127.0.0.1:9222")
    ap.add_argument("--continue-on-new", action="store_true")
    args = ap.parse_args()

    DRY = args.dry_run
    STRICT = not args.continue_on_new
    LOG_FILE = os.path.join(HERE, f"run_new_{args.slot}.log")

    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    street_id, street_name = street_info(args.street)
    log("=" * 64)
    log(f"OYNA (slot): {args.slot}")
    log(f"KO'CHA: {args.street} -> {street_name}")
    log(f"BOT v2 | slot {args.slot} | ko'cha: {street_name} ({street_id})")
    log(f"limit={args.limit or 'hammasi'} houses={args.houses or 'hammasi'} dry={DRY}")

    log("Baza yuklanmoqda...")
    houses, by_pinfl, by_name, member_house, dead = load_base(BASE)
    log(f"Baza: {len(houses)} xonadon, o'lgan: {dead}")
    log("IIB (sanalar) yuklanmoqda...")
    iib = load_iib_dates(IIB)
    log(f"IIB: {len(iib)} sana")

    rep_dir = os.path.join(HERE, "hisobot", f"ул_{street_name}")
    rl.set_dir(rep_dir)
    log(f"hisobot: {rep_dir}")

    st = load_state(street_id)
    if args.restart:
        st = {"processed": [], "added": [], "last": 0}
    processed = set(st.get("processed", []))
    site_pins_all = set()  # keyingi versiyada to'ldiriladi

    stats = {"ok": 0, "fail": 0, "skip": 0, "attempts": 0, "limit": args.limit or 0}

    with sync_playwright() as pw:
        b = pw.chromium.connect_over_cdp(args.cdp)
        ctx = b.contexts[0]
        anchor = next((x for x in ctx.pages if "online-mahalla.uz" in x.url), ctx.pages[0])
        pages = m.ensure_pages(ctx, args.slot, anchor)
        pg = pages[args.slot - 1]
        pg.bring_to_front()
        m.install_page_fixes(pg)

        log("Ko'cha jadvali o'qilmoqda...")
        all_houses = read_all_houses(pg, street_id)
        if all_houses == -1:
            log("!!! LOGIN KERAK — Chrome'da saytga kiring, keyin qayta ishga tushiring.")
            return
        log(f"Xonadonlar: {len(all_houses)}")
        if not all_houses:
            log("!!! Xonadon topilmadi — 10 s dan keyin qayta urinish")
            pg.wait_for_timeout(10000)
            all_houses = read_all_houses(pg, street_id)
            if all_houses == -1:
                log("!!! LOGIN KERAK — Chrome'da saytga kiring, keyin qayta ishga tushiring.")
                return
            log(f"Xonadonlar (qayta): {len(all_houses)}")
        if not all_houses:
            log("!!! Xonadon topilmadi — to'xtadim")
            return

        start = max(1, args.start or 1)
        done = 0
        for idx in range(start - 1, len(all_houses)):
            h = all_houses[idx]
            if h["cad"] in processed:
                continue
            if args.houses and done >= args.houses:
                break
            if stats["limit"] and stats["attempts"] >= stats["limit"]:
                break
            done += 1
            h["street_name"] = street_name
            log(f"[{idx + 1}/{len(all_houses)}] {h['cad']} — {h['fio']}")
            try:
                res, detail = process_house(pg, h, by_pinfl, by_name, member_house,
                                            iib, site_pins_all, street_name, stats)
            except Exception as e:
                log(f"  XATO -> {e}")
                traceback.print_exc()
                stats["fail"] += 1
                res, detail = "ok", ""
            if res == "new":
                if STRICT:
                    log("=" * 64)
                    log(f"!!! YANGI HOLAT: {detail}")
                    log("Bot to'xtadi. Tekshirib, --continue-on-new bilan davom etish mumkin.")
                    save_state(street_id, {"processed": sorted(processed),
                                           "added": sorted(site_pins_all),
                                           "last": idx + 1})
                    rl.flush()
                    return
                log(f"!!! YANGI HOLAT (davom etildi): {detail}")
                rl.log_fail(h["fio"], "", "", "", h["cad"],
                            f"Yangi holat: {detail}", street=street_name)
            processed.add(h["cad"])
            st = {"processed": sorted(processed), "added": sorted(site_pins_all),
                  "last": idx + 1}
            save_state(street_id, st)
            rl.flush()
            pg.wait_for_timeout(400)

    rl.flush()
    log("-" * 64)
    log(f"TUGADI: qo'shildi={stats['ok']}, xato={stats['fail']}, "
        f"o'tkazildi={stats['skip']}, urinish={stats['attempts']}")
    log("=" * 64)


if __name__ == "__main__":
    main()

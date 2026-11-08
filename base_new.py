"""Yangi baza (2026) — xonadonlarga ajratish.

Manba: C:\\Users\\Admin\\Desktop\\Yangi_Lolazor_baza27_07_2026.xlsx
Varaq: «Фуқаролар буйича руйхат»

Ustunlar:
    A № | B F.I.Sh. | C tug'ilgan sana | D ko'cha | E millat
    F guvohnoma/pasport | G JSHSHIR | H telefon | I ish joyi

Xonadon tuzilishi FORMATLASH bilan berilgan:
    * qalin shrift YOKI rangli fon (sariq/yashil/pushti) YOKI qora bo'lmagan
      matn rangi = YANGI XONADON EGASI
    * qizil fon yoki qizil shrift = O'LGAN -> butunlay tashlab ketiladi (50 ta)
    * oq fon / theme-rangli fon egasi belgisi EMAS
    * qolgan qatorlar = o'sha xonadon a'zolari (keyingi egasigacha)
"""
import os
import re
from datetime import datetime, date

import openpyxl

BASE_PATH = os.environ.get(
    "MAHALLA_BASE", r"C:\Users\Admin\Desktop\Yangi_Lolazor_baza27_07_2026.xlsx")
SHEET = "Фуқаролар буйича руйхат"
HEADER_ROWS = 3          # 1-3 qatorlar sarlavha

# --- ko'cha nomlari: bazadagi (lotin) -> saytdagi (kirill) -----------------
STREET_MAP = {
    "universitet": "Университет",
    "qoʻshbuloq": "Қушбулоқ", "qo'shbuloq": "Қушбулоқ", "qushbuloq": "Қушбулоқ",
    "tosh": "Тош",
    "fazogir": "Фазогир",
    "nakkosh": "Наққош", "naqqosh": "Наққош",
    "chukursoy": "Чуқурсой",
    "muynoq": "Мўйноқ",
    "karshi": "Қарши",
    "namangan": "Наманган",
    "karagay": "Қарағай",
    "temuriylar": "Темурий",
    "ellikkalʼa": "Элликқалъа", "ellikkalla": "Элликқалъа", "ellikkal'a": "Элликқалъа",
    "binafsha": "Бинафша",
    "registon": "Регистон",
    "j.neru": "Ж.Неру", "j. neru": "Ж. Неру", "jneru": "Ж.Неру",
    "lutfkor": "Лутфкор",
    "denov": "Денов",
    "uysoz": "Уйсоз",
    "korakul": "Қоракўл",
    "uchkuduk": "Учкудук",
    "paxlavonlar": "Паҳлавонлар",
    "kumushsimon": "Кумишсимон",
    "charxpalag": "Чархпалак",
    "obunachi": "Обуначи",
    "sh.burxanov": "Ш.Бурханов", "sh.burxonov": "Ш.Бурханов", "sh burxanov": "Ш.Бурханов",
    "shabnam": "Шабнам",
    "asr": "Аср",
}

ROMAN = {"I": "I", "II": "II", "III": "III", "IV": "IV", "V": "V", "VI": "VI",
         "L": "I", "LL": "II", "LLL": "III", "1": "I", "2": "II", "3": "III"}


def norm_street(raw):
    """'Obunachi 104' -> 'Обуначи'; 'Universitet' -> 'Университет'."""
    s = (raw or "").strip()
    if not s:
        return ""
    # uy raqamini olib tashlaymiz
    s2 = re.sub(r"\s+\d+[a-zA-Zа-яА-Я]?(\s*\(.*\))?$", "", s).strip()
    s2 = re.sub(r"\s+\d+$", "", s2).strip()
    key = re.sub(r"\s+", " ", s2).lower().replace("’", "ʼ")
    if key in STREET_MAP:
        return STREET_MAP[key]
    # bo'lakma-bo'lak urinish
    for k, v in STREET_MAP.items():
        if key.startswith(k):
            return v
    return s2


def parse_doc(raw):
    """Hujjat satrini (seriya, raqam, tur) ga ajratadi.

    Qaytaradi: (series, number, doc_type)
      doc_type = 'cert' (туғилганлик ҳақида гувоҳнома) yoki 'pass' (паспорт / ID)
    """
    s = re.sub(r"\s+", " ", (raw or "")).strip().upper()
    if not s:
        return "", "", ""

    # 1) guvohnoma: [rim raqam] [chiziqcha] QQ/KK/KQ/T ... raqamlar
    m = re.match(r"^(I{1,3}|IV|V|VI|L{1,3}|1|2|3)?\s*-?\s*([A-Z]{1,3})\s*-?\s*(\d{5,9})$", s)
    if m:
        roman, letters, num = m.group(1), m.group(2), m.group(3)
        if letters in ("QQ", "KK", "KQ", "QK", "QQQ"):
            roman = ROMAN.get(roman or "", "") or "I"
            series = roman + "-QQ"
            return series, num, "cert"
        if letters == "T" or (letters.startswith("T") and len(letters) <= 2):
            return letters, num, "cert"
        # pasport: KA/AD/AE...
        if len(letters) == 2:
            return letters, num, "pass"

    # 2) faqat raqam yoki boshqa holat — seriya yo'q
    m2 = re.match(r"^([A-Z]{2})\s*(\d{6,9})$", s)
    if m2:
        return m2.group(1), m2.group(2), "pass"
    nums = re.sub(r"\D", "", s)
    letters = re.sub(r"[^A-Z]", "", s)
    return letters[:3], nums, ("cert" if "Q" in letters else "pass")


def _fill_rgb(cell):
    fl = cell.fill
    if fl is None or fl.fill_type is None:
        return None
    fg = fl.fgColor
    if fg is None or fg.type != "rgb":
        return None
    return fg.rgb


def _is_red(rgb):
    if not rgb or len(rgb) != 8:
        return False
    try:
        r, g, b = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
    except ValueError:
        return False
    return r > 150 and g < 110 and b < 110


def _font_rgb(cell):
    fo = cell.font
    if fo is None or fo.color is None:
        return None
    if fo.color.type == "rgb":
        return fo.color.rgb
    return None


def _font_marks_owner(cell):
    """Qora/qizil bo'lmagan matn rangi ham xonadon egasini bildiradi."""
    rgb = _font_rgb(cell)
    return bool(rgb and rgb not in ("FF000000", "FFFF0000", "FF9C0006"))


def _as_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if v:
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})", str(v))
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", str(v))
        if m:
            try:
                return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                pass
    return None


def person(row_cells, row_no):
    """Qatordan odam yozuvi."""
    def val(i):
        v = row_cells[i].value
        return "" if v is None else (v.strip() if isinstance(v, str) else v)

    doc_raw = str(val(5) or "").strip()
    series, number, dtype = parse_doc(doc_raw)
    birth = _as_date(val(2))
    pinfl = re.sub(r"\D", "", str(val(6) or ""))
    if len(pinfl) != 14:
        pinfl = ""
    phone = re.sub(r"\D", "", str(val(7) or ""))
    return {
        "row": row_no,
        "fio": str(val(1) or "").strip(),
        "birth": birth,
        "street_raw": str(val(3) or "").strip(),
        "street": norm_street(str(val(3) or "")),
        "doc_raw": doc_raw,
        "series": series,
        "number": number,
        "doc_type": dtype,
        "pinfl": pinfl,
        "phone": phone,
        "work": str(val(8) or "").strip(),
    }


def load(path=BASE_PATH):
    """Bazani o'qiydi -> xonadonlar ro'yxati.

    Har bir xonadon:
        {"owner": person, "members": [person, ...], "street": "Университет"}
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[SHEET]

    households = []
    cur = None
    skipped_dead = 0

    for i in range(HEADER_ROWS + 1, ws.max_row + 1):
        row = [ws.cell(i, c) for c in range(1, 10)]
        if not row[1].value:
            continue

        fill = _fill_rgb(row[1])
        red = _font_rgb(row[1]) == "FFFF0000" or _is_red(fill)
        if red:
            skipped_dead += 1
            continue

        owner_mark = (bool(row[1].font.bold) or (fill is not None)
                      or _font_marks_owner(row[1]))
        p = person(row, i)

        if owner_mark or cur is None:
            cur = {"owner": p, "members": [], "street": p["street"]}
            households.append(cur)
        else:
            cur["members"].append(p)

    wb.close()
    return households, skipped_dead


def norm_name(s):
    """F.I.Sh. solishtirish uchun normallashtirish."""
    s = (s or "").upper()
    s = s.replace("Ў", "У").replace("Қ", "Ғ").replace("Ҳ", "Ҳ")
    s = re.sub(r"[^A-ZА-ЯЁ\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def index(households):
    """Egasi bo'yicha indeks: JSHSHIR -> xonadon, F.I.Sh. -> xonadon."""
    by_pinfl, by_name = {}, {}
    for h in households:
        o = h["owner"]
        if o["pinfl"]:
            by_pinfl.setdefault(o["pinfl"], h)
        n = norm_name(o["fio"])
        if n:
            by_name.setdefault(n, h)
    return by_pinfl, by_name


if __name__ == "__main__":
    import sys
    try:                       # Windows konsolida Қ/Ў yiqilmasin
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    hs, dead = load()
    print("xonadonlar:", len(hs), "| o'lgan (tashlandi):", dead)
    sizes = [1 + len(h["members"]) for h in hs]
    print("o'rtacha: %.1f | jami odam: %d" % (sum(sizes) / len(sizes), sum(sizes)))
    from collections import Counter
    c = Counter(h["street"] for h in hs)
    print("\nko'chalar:", len(c))
    for k, v in c.most_common():
        print("   %-14s %4d" % (k or "(noma'lum)", v))
    by_pinfl, by_name = index(hs)
    print("\nindeks: JSHSHIR=%d, F.I.Sh.=%d" % (len(by_pinfl), len(by_name)))
    cert = sum(1 for h in hs for p in [h["owner"]] + h["members"] if p["doc_type"] == "cert")
    nop = sum(1 for h in hs for p in [h["owner"]] + h["members"] if not p["pinfl"])
    print("guvohnoma: %d | JSHSHIRsiz: %d" % (cert, nop))

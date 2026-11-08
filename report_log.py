"""Hisobot fayllari — endi EXCEL (.xlsx).

Har bir ko'cha papkasida:
    muvaffaqiyatli.xlsx    — qo'shilganlar
    muvaffaqiyatsiz.xlsx   — qo'shilmaganlar (sabab bilan)

Ustunlar: №, F.I.Sh., Seriya, Raqam, JSHSHIR, Kadastr, Ko'cha, Sana,
          Telefon, Hujjat turi, Holat / Sabab

Funksiyalar avvalgi API bilan bir xil ishlaydi:
    set_dir(path), log_success(...), log_fail(...), log_minor(...)
"""
import json
import os
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

DIR = os.path.dirname(os.path.abspath(__file__))
OK_FILE = os.path.join(DIR, "muvaffaqiyatli.xlsx")
FAIL_FILE = os.path.join(DIR, "muvaffaqiyatsiz.xlsx")

HEADERS = ["№", "F.I.Sh.", "Seriya", "Raqam", "JSHSHIR", "Kadastr",
           "Ko'cha", "Sana", "Telefon", "Hujjat turi", "Holat / Sabab"]
WIDTHS = [5, 34, 9, 12, 16, 22, 20, 12, 19, 26, 44]

HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

_cache = {}   # path -> (workbook, worksheet, set(pinfl))


def set_dir(path):
    """Hisobot papkasini o'zgartiradi (har bir ko'cha uchun alohida)."""
    global DIR, OK_FILE, FAIL_FILE
    DIR = path if os.path.isabs(path) else os.path.join(os.path.dirname(os.path.abspath(__file__)), path)
    os.makedirs(DIR, exist_ok=True)
    OK_FILE = os.path.join(DIR, "muvaffaqiyatli.xlsx")
    FAIL_FILE = os.path.join(DIR, "muvaffaqiyatsiz.xlsx")
    _cache.clear()


def _sheet(path):
    """Fayl varsa ochadi, yo'qsa sarlavha bilan yaratadi."""
    if path in _cache:
        return _cache[path]

    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
        seen = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and len(row) > 4 and row[4]:
                seen.add(str(row[4]).strip())
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Hisobot"
        ws.append(HEADERS)
        for i, w in enumerate(WIDTHS, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for c in ws[1]:
            c.fill = HDR_FILL
            c.font = HDR_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER
        ws.freeze_panes = "A2"
        seen = set()

    _cache[path] = (wb, ws, seen)
    return _cache[path]


def _append(path, row, dedup_key):
    wb, ws, seen = _sheet(path)
    key = str(dedup_key or "").strip()
    if key and key in seen:
        return False
    n = ws.max_row  # sarlavha + mavjud qatorlar
    ws.append([n] + list(row))
    for c in ws[ws.max_row]:
        c.border = BORDER
        c.alignment = Alignment(vertical="top", wrap_text=(c.column == 11))
    if key:
        seen.add(key)
    try:
        wb.save(path)
    except PermissionError:
        # fayl Excel'da ochiq bo'lsa — keyingi safar saqlanadi
        pass
    return True


def flush():
    """Barcha ochiq ish kitoblarini diskka yozadi."""
    for path, (wb, ws, seen) in list(_cache.items()):
        try:
            wb.save(path)
        except PermissionError:
            pass


def _sink(ev, payload):
    """GitHub worker uchun JSONL oqim (REPORT_JSONL env bilan yoqiladi)."""
    path = os.environ.get("REPORT_JSONL")
    if not path:
        return
    try:
        payload = dict(payload)
        payload["ev"] = ev
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except OSError:
        pass


def log_success(fio, series, number, pinfl, cadastral, street="",
                phone="", note="Qo'shildi", doc_type="", **kw):
    added = _append(
        OK_FILE,
        [fio, series or "-", number or "-", pinfl or "-", cadastral or "-", street or "-",
         datetime.now().strftime("%d.%m.%Y"), phone or "-", doc_type or "-", note],
        dedup_key=pinfl or f"{fio}|{cadastral}",
    )
    if added:
        _sink("ok", {"fio": fio, "series": series, "number": number, "pinfl": pinfl,
                     "cad": cadastral, "street": street, "phone": phone,
                     "note": note, "doc_type": doc_type})
    return added


def log_fail(fio, series, number, pinfl, cadastral, reason, street="",
             doc_type="", **kw):
    added = _append(
        FAIL_FILE,
        [fio, series or "-", number or "-", pinfl or "-", cadastral or "-", street or "-",
         datetime.now().strftime("%d.%m.%Y"), "-", doc_type or "-", reason],
        dedup_key=f"{fio}|{pinfl}|{cadastral}",
    )
    if added:
        _sink("fail", {"fio": fio, "series": series, "number": number, "pinfl": pinfl,
                       "cad": cadastral, "reason": reason, "street": street,
                       "doc_type": doc_type})
    return added


def log_minor(fio, series, number, pinfl, cadastral, birth="", street="", **kw):
    """Endi voyaga yetmaganlar alohida faylga yozilmaydi — ular ham qo'shiladi.
    Moslik uchun funksiya qoldirilgan; muvaffaqiyatli faylga yozadi."""
    return log_success(fio, series, number, pinfl, cadastral, street=street,
                       note=f"Bola (tug'ilgan: {birth})", doc_type="гувоҳнома")

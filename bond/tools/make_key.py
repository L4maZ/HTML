"""Sinh Key_YYYYMMDD.xlsx tu File 02.

Doc 'Linked (1)' (khoi 2), 'Report' (nhan dinh) va 'Run Tool' (truc ngay),
xuat ra mot workbook phang co cot KeyID lam neo. Vi tri dong khong con y nghia
sau buoc nay: HTML tra cuu bang KeyID.

    python3 make_key.py <File02.xlsm> [thu_muc_dich]
"""

import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BOOKS = [
    ("TB_MSB", "2.1.", "TRADING BOOK NỘI BỘ", 5, 30),
    ("BB_MSB", "2.2.", "BANKING BOOK NỘI BỘ", 32, 53),
    ("TB_SBV", "2.3.", "TRADING BOOK SBV", 55, 60),
    ("BB_SBV", "2.4.", "BANKING BOOK SBV", 62, 67),
    ("OTHER", "2.5.", "KHÁC", 69, 69),
    ("FIBOND", "2.6.", "FI Bond & CD", 71, 74),
]

COL = dict(stt=1, cat=2, label=3, today=4, dtd=5, yest=6, lm=7, lq=8, ly=9,
           limit=10, used=11, light=12)

TEXTS = [
    ("txt.market", "Report", "B4", "Thông tin thị trường"),
    ("txt.assessment", "Report", "H4", "Tuân thủ hạn mức & đánh giá rủi ro"),
    ("txt.noteTB", "Report", "Q10", "Ghi chú Trading Book"),
    ("txt.noteTPCP", "Report", "Q59", "Ghi chú tỷ lệ đầu tư TPCP"),
    ("txt.noteFI", "Report", "Q66", "Ghi chú cơ cấu FI Bond"),
    ("txt.noteFV", "Report", "B110", "Ghi chú Fair value"),
]

HDR_FILL = PatternFill("solid", fgColor="7B2D3B")
HDR_FONT = Font(color="FFFFFF", bold=True, size=10)
EDIT_FONT = Font(color="0070C0")
SUB_FONT = Font(italic=True, color="6B665E")


def slug(text):
    s = unicodedata.normalize("NFD", str(text))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    s = re.sub(r"[^0-9a-zA-Z]+", " ", s).strip()
    parts = s.split()
    if not parts:
        return "unnamed"
    head = parts[0].lower()
    return head + "".join(p.capitalize() for p in parts[1:])


def vndate(v):
    return v.strftime("%d/%m/%Y") if isinstance(v, (date, datetime)) else v


def as_of(wb):
    rt = wb["Run Tool"]
    out = {}
    for key, cell in (("today", "B2"), ("yest", "B3"), ("lm", "B4"), ("ly", "B7")):
        v = rt[cell].value
        out[key] = v.date() if isinstance(v, datetime) else v
    lq = wb["Linked (1)"]["H3"].value
    out["lq"] = lq.date() if isinstance(lq, datetime) else lq
    return out


def read_block(ws, key, code, title, r0, r1):
    rows = []
    seen = {}
    for r in range(r0, r1 + 1):
        label = ws.cell(r, COL["label"]).value
        if label in (None, ""):
            continue
        label = str(label).replace("\n", " ").strip()
        sub = bool(re.match(r"^[ab]\)", label))
        base = slug(label)
        seen[base] = seen.get(base, 0) + 1
        if seen[base] > 1:
            base = "%s%d" % (base, seen[base])
        rec = {
            "keyid": "S2.%s.%s" % (key, base),
            "book": key,
            "code": code,
            "section": title,
            "stt": ws.cell(r, COL["stt"]).value,
            "cat": ws.cell(r, COL["cat"]).value,
            "label": label,
            "sub": sub,
        }
        for f in ("today", "dtd", "yest", "lm", "lq", "ly", "limit", "used", "light"):
            rec[f] = ws.cell(r, COL[f]).value
        rows.append(rec)
    return rows


def build(src, outdir):
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb["Linked (1)"]
    dates = as_of(wb)

    data = []
    for key, code, title, r0, r1 in BOOKS:
        data.extend(read_block(ws, key, code, title, r0, r1))

    out = openpyxl.Workbook()

    meta = out.active
    meta.title = "META"
    meta.append(["Key", "Value", "Ghi chu"])
    meta.append(["schema", "bond.key.v1", "Phien ban cau truc file key"])
    meta.append(["asOf", vndate(dates["today"]), "Ngay bao cao"])
    meta.append(["dateYest", vndate(dates["yest"]), "Cot Yesterday"])
    meta.append(["dateLastMonth", vndate(dates["lm"]), "Cot Last month"])
    meta.append(["dateLastQuarter", vndate(dates["lq"]), "Cot Last Quarter"])
    meta.append(["dateLastYear", vndate(dates["ly"]), "Cot Last Year"])
    meta.append(["rptTitle", "Báo cáo rủi ro thị trường", "Tieu de trang"])
    meta.append(["rptSubtitle", "Báo cáo Desk Bond", "Phu de"])
    meta.append(["unitNote", "Đơn vị: tỷ VND, trừ khi ghi khác. Giá trị âm là lỗ.", ""])
    meta.append(["generatedAt", datetime.now().strftime("%d/%m/%Y %H:%M"), "Thoi diem sinh file"])

    dsh = out.create_sheet("DATA_S2")
    headers = ["KeyID", "Book", "Code", "Section", "STT", "Nhom", "ChiTieu", "Sub",
               "Today", "DtD", "Yesterday", "LastMonth", "LastQuarter", "LastYear",
               "Limit", "Used", "Light"]
    dsh.append(headers)
    for rec in data:
        dsh.append([
            rec["keyid"], rec["book"], rec["code"], rec["section"], rec["stt"],
            rec["cat"], rec["label"], 1 if rec["sub"] else 0,
            rec["today"], rec["dtd"], rec["yest"], rec["lm"], rec["lq"], rec["ly"],
            rec["limit"], rec["used"], rec["light"],
        ])

    tsh = out.create_sheet("TEXT")
    tsh.append(["KeyID", "Mo ta", "Nguon", "Value"])
    for keyid, sheet, cell, desc in TEXTS:
        tsh.append([keyid, desc, "%s!%s" % (sheet, cell), wb[sheet][cell].value])

    style(meta, [16, 46, 40], value_col=2)
    style(dsh, [30, 10, 7, 26, 6, 20, 34, 5] + [14] * 6 + [12, 10, 10], value_col=None)
    style(tsh, [18, 40, 16, 90], value_col=4, wrap_col=4)

    for row in dsh.iter_rows(min_row=2, max_row=dsh.max_row):
        if row[7].value == 1:
            for c in row:
                c.font = SUB_FONT

    dsh.freeze_panes = "I2"
    tsh.freeze_panes = "D2"

    stamp = dates["today"]
    if isinstance(stamp, (date, datetime)):
        stamp = stamp.strftime("%Y%m%d")
    else:
        stamp = "unknown"
    dest = Path(outdir) / ("Key_%s.xlsx" % stamp)
    out.save(dest)
    return dest, len(data)


def style(ws, widths, value_col=None, wrap_col=None):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    if value_col:
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, value_col)
            cell.font = EDIT_FONT
            if wrap_col == value_col:
                cell.alignment = Alignment(wrap_text=True, vertical="top")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    target = sys.argv[2] if len(sys.argv) > 2 else "."
    path, n = build(sys.argv[1], target)
    print("%s — %d dong DATA_S2" % (path, n))

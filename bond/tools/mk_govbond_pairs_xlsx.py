#!/usr/bin/env python3
"""Xuất file Excel để Jak tự đối chiếu kết quả ghép cặp GovBond với file gốc.

    cd bond/tools && python3 mk_govbond_pairs_xlsx.py

Ba sheet, thiết kế để pivot đối chiếu HAI CHIỀU:
  1. `Cap deal`  — 198 cặp đã ghép, mỗi dòng một cặp.
  2. `Chan le`   — 374 chân lẻ nguồn (sau lọc folder), `GrossAmount` ĐÃ CHUẨN HOÁ.
                   File gốc có 16 dòng GrossAmount là CHUỖI đuôi 'K' (đơn vị nghìn);
                   SUM thẳng trong Excel sẽ bỏ qua chúng và ra 84.284 triệu thay vì
                   263.575 triệu. Sheet này quy đổi sẵn, kèm cờ đánh dấu 16 dòng đó.
  3. `Doi chieu` — CÔNG THỨC SỐNG tie hai sheet trên với nhau. Mở bằng Excel là tự
                   tính; nếu hai vế lệch nhau thì có gì đó sai.

Lưu ý khi dò tay: 24/198 cặp bị TÁCH CHÂN (một chân lớn ghép với nhiều chân nhỏ), cột
Gross của các dòng đó là phần PRO-RATA theo khối lượng, không bằng Gross của chân gốc
trong file nguồn. Cột `Tach chan` đánh dấu sẵn để khỏi tưởng sai.

Format trắng đen, không viền, không comment — theo yêu cầu của Jak.
"""
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from ghep_cap_deal_govbond import matched, byid, R, KEEP_FOLDER

OUT = '/home/user/HTML/bond/Doi_chieu_GD_Bond_20260608_20260904.xlsx'
FONT, DT, VND, TR, PC = 'Arial', 'DD/MM/YYYY', '#,##0', '#,##0.0', '0.00%'
HF, CF = Font(name=FONT, size=10, bold=True), Font(name=FONT, size=10)


def sheet(ws, cols, rows, widths=None):
    for i, t in enumerate(cols, 1):
        ws.cell(row=1, column=i, value=t).font = HF
    for ri, vals in enumerate(rows, 2):
        for ci, v in enumerate(vals, 1):
            ws.cell(row=ri, column=ci, value=v).font = CF
    for i in range(1, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(i, 14)
    ws.auto_filter.ref = f'A1:{get_column_letter(len(cols))}{1 + len(rows)}'
    ws.freeze_panes = 'A2'


wb = openpyxl.Workbook()

# ---------------------------------------------------------------- 1. Cặp deal
ws1 = wb.active
ws1.title = 'Cap deal'
C1 = ['Deal S', 'Deal B', 'Ma TP', 'Doi tac', 'Coupon (%)', 'Menh gia (ty)',
      'Chan thanh toan truoc', 'Ky han (ngay)', 'Settlement S', 'Settlement B',
      'CaptureDate S', 'CaptureDate B', 'Gross S (VND)', 'Gross B (VND)',
      'Lai/lo (VND)', 'Lai/lo (trieu)', '%/nam', 'Dau', 'Nhom', 'Tach chan',
      'Yield S', 'Yield B', 'Dyield (bp)', 'Folder S', 'Folder B']
rows1 = []
for p in sorted(matched, key=lambda x: (x['scap'], str(x['sid']))):
    s_int = isinstance(p['sid'], int)
    tach = 'Co' if s_int and byid[p['sid']]['Quantity'] != byid[p['bid']]['Quantity'] else ''
    s_cash = p['cash'] if p['dirn'] == 'A' else p['cash'] + p['pnl']
    b_cash = p['cash'] - p['pnl'] if p['dirn'] == 'A' else p['cash']
    rows1.append([
        p['sid'], p['bid'], p['paper'], p['cpty'], p['cpn'], p['face'],
        'Sell (repo)' if p['dirn'] == 'A' else 'Buy (reverse repo)', p['days'],
        p['sset'], p['bset'], p['scap'], p['bcap'],
        round(s_cash), round(b_cash), round(p['pnl']), p['pnl'] / 1e6,
        p['rate'] / 100, 'duong' if p['pnl'] > 0 else ('am' if p['pnl'] < 0 else 'bang 0'),
        p['loai'], tach, p['sy'], p['by'], p['dy'], p['sf'], p['bf'],
    ])
sheet(ws1, C1, rows1, {7: 20, 19: 15})
for r in range(2, 2 + len(rows1)):
    for c in (9, 10, 11, 12):
        ws1.cell(row=r, column=c).number_format = DT
    for c in (13, 14, 15):
        ws1.cell(row=r, column=c).number_format = VND
    ws1.cell(row=r, column=6).number_format = TR
    ws1.cell(row=r, column=16).number_format = TR
    ws1.cell(row=r, column=17).number_format = PC

# ---------------------------------------------------------------- 2. Chân lẻ
ws2 = wb.create_sheet('Chan le')
C2 = ['Deal id', 'Loai', 'Ma TP', 'Doi tac', 'Coupon (%)', 'Quantity', 'FaceValue',
      'Gross da chuan hoa (VND)', 'Gross goc la CHU?', 'Settlement', 'CaptureDate',
      'Folder', 'Yield']
raw = openpyxl.load_workbook(
    '/home/user/HTML/bond/source/RPBOD_B002_2026.08.27.xlsx', data_only=True)['Sheet1']
hdr = [raw.cell(row=1, column=c).value for c in range(1, 26)]
rows2 = []
for r in range(2, raw.max_row + 1):
    d = dict(zip(hdr, [raw.cell(row=r, column=c).value for c in range(1, 26)]))
    if d['BondsDeals_Id'] is None or d['Folders_ShortName'] not in KEEP_FOLDER:
        continue
    is_txt = isinstance(d['GrossAmount'], str)
    gross = float(d['GrossAmount'].strip().rstrip('K')) * 1000.0 if is_txt else float(d['GrossAmount'])
    rows2.append([d['BondsDeals_Id'], d['DealType'], d['Bonds_ShortName'], d['Cpty_ShortName'],
                  d['CouponRate'], d['Quantity'], d['FaceValue'], round(gross),
                  'CHU (x1000)' if is_txt else '', d['SettlementDate'], d['CaptureDate'],
                  d['Folders_ShortName'], d['Yield']])
sheet(ws2, C2, rows2, {8: 24, 9: 18})
for r in range(2, 2 + len(rows2)):
    for c in (10, 11):
        ws2.cell(row=r, column=c).number_format = DT
    for c in (6, 8):
        ws2.cell(row=r, column=c).number_format = VND

# ---------------------------------------------------------------- 3. Đối chiếu
ws3 = wb.create_sheet('Doi chieu')
n1, n2 = 1 + len(rows1), 1 + len(rows2)
BLOCK = [
    ('A. TONG QUAT', '', ''),
    ('So cap da ghep', f'=COUNTA(\'Cap deal\'!A2:A{n1})', '198'),
    ('So chan le nguon (sau loc folder)', f'=COUNTA(\'Chan le\'!A2:A{n2})', '374'),
    ('', '', ''),
    ('B. KHOI LUONG PHAI OFFSET HET', '', ''),
    ('Sum Quantity chan S', f'=SUMIF(\'Chan le\'!B2:B{n2},"S",\'Chan le\'!F2:F{n2})', '353.844.000'),
    ('Sum Quantity chan B', f'=SUMIF(\'Chan le\'!B2:B{n2},"B",\'Chan le\'!F2:F{n2})', 'phai BANG dong tren'),
    ('Lech', f'=B6-B7', '0'),
    ('', '', ''),
    ('C. LAI/LO - TIE HAI CHIEU', '', ''),
    ('Tu CHAN LE:  Sum Gross(S) - Sum Gross(B)',
     f'=SUMIF(\'Chan le\'!B2:B{n2},"S",\'Chan le\'!H2:H{n2})'
     f'-SUMIF(\'Chan le\'!B2:B{n2},"B",\'Chan le\'!H2:H{n2})', '263.574.510.000'),
    ('Tu CAP DEAL: Sum cot Lai/lo (VND)', f'=SUM(\'Cap deal\'!O2:O{n1})', 'phai BANG dong tren'),
    ('Lech', '=B11-B12', '0  <- khac 0 la co van de'),
    ('', '', ''),
    ('D. MENH GIA', '', ''),
    ('Tu CHAN LE:  Sum Quantity(S) x FaceValue / 1e9',
     f'=SUMIF(\'Chan le\'!B2:B{n2},"S",\'Chan le\'!F2:F{n2})*100000/1000000000', '35.384,4 ty'),
    ('Tu CAP DEAL: Sum cot Menh gia', f'=SUM(\'Cap deal\'!F2:F{n1})', 'phai BANG dong tren'),
    ('Lech', '=B16-B17', '0'),
    ('', '', ''),
    ('E. PHAN BO 4 NHOM (so cap)', '', ''),
    ('Vay tien', f'=COUNTIF(\'Cap deal\'!S2:S{n1},"Vay tien")', '102'),
    ('Cho vay bond', f'=COUNTIF(\'Cap deal\'!S2:S{n1},"Cho vay bond")', '59'),
    ('Cho vay tien', f'=COUNTIF(\'Cap deal\'!S2:S{n1},"Cho vay tien")', '37'),
    ('Vay bond', f'=COUNTIF(\'Cap deal\'!S2:S{n1},"Vay bond")', '0'),
    ('Tong', f'=SUM(B21:B24)', '198'),
    ('', '', ''),
    ('F. KIEM TRA LOGIC PHAN LOAI (dem dong SAI, phai = 0)', '', ''),
    ('Sell truoc + am  ma KHONG phai "Vay tien"',
     f'=SUMPRODUCT((\'Cap deal\'!G2:G{n1}="Sell (repo)")*(\'Cap deal\'!O2:O{n1}<0)'
     f'*(\'Cap deal\'!S2:S{n1}<>"Vay tien"))', '0'),
    ('Sell truoc + duong ma KHONG phai "Cho vay bond"',
     f'=SUMPRODUCT((\'Cap deal\'!G2:G{n1}="Sell (repo)")*(\'Cap deal\'!O2:O{n1}>0)'
     f'*(\'Cap deal\'!S2:S{n1}<>"Cho vay bond"))', '0'),
    ('Buy truoc + duong ma KHONG phai "Cho vay tien"',
     f'=SUMPRODUCT((\'Cap deal\'!G2:G{n1}="Buy (reverse repo)")*(\'Cap deal\'!O2:O{n1}>0)'
     f'*(\'Cap deal\'!S2:S{n1}<>"Cho vay tien"))', '0'),
    ('Buy truoc + am ma KHONG phai "Vay bond"',
     f'=SUMPRODUCT((\'Cap deal\'!G2:G{n1}="Buy (reverse repo)")*(\'Cap deal\'!O2:O{n1}<0)'
     f'*(\'Cap deal\'!S2:S{n1}<>"Vay bond"))', '0'),
    ('', '', ''),
    ('G. KY HAN = KHOANG CACH 2 SETTLEMENT (dem dong SAI, phai = 0)', '', ''),
    ('So dong lech',
     f'=SUMPRODUCT((ABS(\'Cap deal\'!J2:J{n1}-\'Cap deal\'!I2:I{n1})<>\'Cap deal\'!H2:H{n1})*1)', '0'),
]
ws3.cell(row=1, column=1, value='Noi dung').font = HF
ws3.cell(row=1, column=2, value='Cong thuc tinh song').font = HF
ws3.cell(row=1, column=3, value='Phai ra').font = HF
for i, (lbl, formula, expect) in enumerate(BLOCK, 2):
    a = ws3.cell(row=i, column=1, value=lbl)
    a.font = HF if formula == '' and lbl else CF
    if formula:
        c = ws3.cell(row=i, column=2, value=formula)
        c.font = CF
        c.number_format = VND if 'Gross' in lbl or 'Quantity' in lbl or 'Lai/lo' in lbl else 'General'
    ws3.cell(row=i, column=3, value=expect).font = CF
ws3.column_dimensions['A'].width = 52
ws3.column_dimensions['B'].width = 26
ws3.column_dimensions['C'].width = 26

wb.save(OUT)
print(f'saved {OUT}')
print(f'  Sheet "Cap deal" : {len(rows1)} dong')
print(f'  Sheet "Chan le"  : {len(rows2)} dong ({sum(1 for r in rows2 if r[8])} dong Gross goc la chu)')
print(f'  Sheet "Doi chieu": {len(BLOCK)} dong cong thuc song')

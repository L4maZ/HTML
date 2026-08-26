#!/usr/bin/env python3
"""Ghép cặp repo từ export deal FIBond (CD / trái phiếu doanh nghiệp FI).

Schema KHÁC HẲN file TPCP (MSB_RP_DM) — ba bẫy phải nhớ:
  1. QUANTITY vô dụng: 682/805 dòng = 1. Khối lượng thật nằm ở FACE_AMOUNT.
  2. ACCRUED không phải tiền: chỉ nhận 0-6, đó là SỐ NGÀY.
  3. PRICE không tái tạo GROSS: FACE x PRICE/100 + ACCRUED chỉ khớp 69/805 dòng.
     GROSS_AMOUNT đã gồm lãi dồn tích -> luôn lấy thẳng GROSS_AMOUNT.

Điều kiện phân biệt REPO với mua-bán thường: HAI CHÂN NHẬP CÙNG NGÀY (CAPTURE_DATE).
Thiếu điều kiện đó thì chỉ là mua rồi bán lại sau -> xếp vào nhóm `loose`.
"""
import openpyxl
from collections import defaultdict, Counter

F = '/root/.claude/uploads/739f5b4e-bb4f-5d4a-b668-c927f537f4c6/5464ab24-List_deal_ps_FIBond_24.08.2026.xlsx'
ws = openpyxl.load_workbook(F, data_only=True)['Sheet1']
H = [ws.cell(row=1, column=c).value for c in range(1, 20)]
R = [dict(zip(H, [ws.cell(row=r, column=c).value for c in range(1, 20)])) for r in range(2, ws.max_row + 1)]
R = [x for x in R if x['DEAL_ID'] is not None and x['TRANSACTION_STATUS'] != 'N']

# Khoá ghép, theo đúng quy tắc nghiệp vụ đã chốt ở file TPCP:
#   cùng mã giấy tờ + cùng đối tác + ngược chiều + cùng khối lượng
# Ở file FIBond, khối lượng là FACE_AMOUNT (QUANTITY = 1 ở 682/805 dòng).
# Điều kiện phân biệt REPO với mua-bán thường: hai chân được NHẬP CÙNG NGÀY.
g = defaultdict(lambda: {'B': [], 'S': []})
for x in R:
    g[(x['PAPER_CODE'], x['CPTY_CODE'], x['FACE_AMOUNT'])][x['DEAL_TYPE']].append(x)

repo, loose = [], []
for (paper, cpty, face), v in g.items():
    S = sorted(v['S'], key=lambda x: (x['VALUE_DATE'], x['DEAL_ID']))
    B = sorted(v['B'], key=lambda x: (x['VALUE_DATE'], x['DEAL_ID']))
    us, ub = set(), set()
    for same_cap in (True, False):          # lượt 1 chặt, lượt 2 nới
        cands = sorted((abs(s['DEAL_ID'] - b['DEAL_ID']), i, j)
                       for i, s in enumerate(S) for j, b in enumerate(B)
                       if (s['CAPTURE_DATE'] == b['CAPTURE_DATE']) == same_cap)
        for _, i, j in cands:
            if i in us or j in ub:
                continue
            us.add(i); ub.add(j)
            s, b = S[i], B[j]
            first, second = (s, b) if s['VALUE_DATE'] <= b['VALUE_DATE'] else (b, s)
            dirn = 'A' if first is s else 'B'
            gap_cap = abs((s['CAPTURE_DATE'] - b['CAPTURE_DATE']).days)
            rec = dict(
                paper=paper, cpty=cpty, prod=s['PRODUCT_CODE'], face=face,
                sid=s['DEAL_ID'], bid=b['DEAL_ID'], dirn=dirn,
                days=(second['VALUE_DATE'] - first['VALUE_DATE']).days,
                d1=first['VALUE_DATE'], d2=second['VALUE_DATE'],
                cash=first['GROSS_AMOUNT'], gap_cap=gap_cap,
                # dương = có lợi cho MSB, đúng quy ước file TPCP
                pnl=(first['GROSS_AMOUNT'] - second['GROSS_AMOUNT']) if dirn == 'A'
                    else (second['GROSS_AMOUNT'] - first['GROSS_AMOUNT']),
                sy=s['YIELD'], by=b['YIELD'], st=s['TRANSACTION_STATUS'] + b['TRANSACTION_STATUS'],
                scap=s['CAPTURE_DATE'], bcap=b['CAPTURE_DATE'])
            (repo if same_cap else loose).append(rec)

left = [x for x in R if not any(x['DEAL_ID'] in (p['sid'], p['bid']) for p in repo + loose)]

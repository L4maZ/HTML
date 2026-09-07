#!/usr/bin/env python3
"""Ghép cặp repo từ export deal FIBond (CD / trái phiếu doanh nghiệp FI).

Schema KHÁC HẲN file TPCP (MSB_RP_DM) — ba bẫy phải nhớ:
  1. QUANTITY vô dụng: 682/805 dòng = 1. Khối lượng thật nằm ở FACE_AMOUNT.
  2. ACCRUED không phải tiền: chỉ nhận 0-6, đó là SỐ NGÀY.
  3. PRICE không tái tạo GROSS: FACE x PRICE/100 + ACCRUED chỉ khớp 69/805 dòng.
     GROSS_AMOUNT đã gồm lãi dồn tích -> luôn lấy thẳng GROSS_AMOUNT.

Điều kiện ghép cặp = repo/reverse repo: cùng mã giấy tờ + cùng đối tác +
ngược chiều (S/B) + cùng FACE_AMOUNT. KHÔNG loại trừ theo CAPTURE_DATE —
xác nhận của Jak (08/2026): BO (back office) có thể nhập hai chân của cùng
một hợp đồng cách xa ngày nhau, việc đó không có nghĩa là không phải deal
repo. Toàn bộ cặp khớp đủ 4 điều kiện trên đều tính là Repo/Reverse Repo.

CAPTURE_DATE hai chân lệch xa nhau chỉ dùng để TÁCH RIÊNG một sheet cho dễ
soát BO (`gap_cap`), không dùng để loại trừ khỏi kết quả.

PHÂN LOẠI 4 NHÓM (chốt 09/2026, thay cho nhị phân đi vay/cho vay cũ):
theo (chân nào thanh toán TRƯỚC) x (Gross bên nào LỚN HƠN). pnl = Gross(S)
− Gross(B) luôn luôn, bất kể chân nào trước — xem chứng minh trong `emit()`.

    Buy trước, Sell sau, Gross(S) > Gross(B)  ->  Cho vay tien
    Buy trước, Sell sau, Gross(S) < Gross(B)  ->  Vay bond
    Sell trước, Buy sau, Gross(S) > Gross(B)  ->  Di vay tien
    Sell trước, Buy sau, Gross(S) < Gross(B)  ->  Cho vay bond

Đây không phải "lãi/lỗ của đi vay/cho vay tiền" — "Cho vay bond" (Sell
trước, trả nhiều hơn khi mua lại) là MSB cho mượn bond, trả rebate interest
trên tiền cọc nhận được — đó là chiều BÌNH THƯỜNG của repo bán trước, không
phải một khoản lỗ bất thường. Dữ liệu 24/08/2026: 270/271 cặp Sell-trước
rơi vào "Cho vay bond"; 59/59 cặp Buy-trước đều "Cho vay tien" (không cặp
nào "Vay bond" trong kỳ này).
"""
import openpyxl
from collections import defaultdict, Counter

F = '/root/.claude/uploads/739f5b4e-bb4f-5d4a-b668-c927f537f4c6/5464ab24-List_deal_ps_FIBond_24.08.2026.xlsx'
ws = openpyxl.load_workbook(F, data_only=True)['Sheet1']
H = [ws.cell(row=1, column=c).value for c in range(1, 20)]
R = [dict(zip(H, [ws.cell(row=r, column=c).value for c in range(1, 20)])) for r in range(2, ws.max_row + 1)]
R = [x for x in R if x['DEAL_ID'] is not None and x['TRANSACTION_STATUS'] != 'N']

g = defaultdict(lambda: {'B': [], 'S': []})
for x in R:
    g[(x['PAPER_CODE'], x['CPTY_CODE'], x['FACE_AMOUNT'])][x['DEAL_TYPE']].append(x)

matched = []
for (paper, cpty, face), v in g.items():
    S = sorted(v['S'], key=lambda x: (x['VALUE_DATE'], x['DEAL_ID']))
    B = sorted(v['B'], key=lambda x: (x['VALUE_DATE'], x['DEAL_ID']))
    us, ub = set(), set()

    def emit(s, b):
        first, second = (s, b) if s['VALUE_DATE'] <= b['VALUE_DATE'] else (b, s)
        dirn = 'A' if first is s else 'B'
        gap_cap = abs((s['CAPTURE_DATE'] - b['CAPTURE_DATE']).days)
        # pnl = Gross(S) - Gross(B) LUÔN LUÔN, bất kể chân nào thanh toán trước
        # (hai nhánh dirn rút gọn về cùng biểu thức này — xem chứng minh ở
        # docs/bond-deal-pairing.md).
        pnl = (first['GROSS_AMOUNT'] - second['GROSS_AMOUNT']) if dirn == 'A' \
            else (second['GROSS_AMOUNT'] - first['GROSS_AMOUNT'])
        # Phân loại 4 nhóm theo (chân nào thanh toán trước) x (Gross bên nào lớn hơn).
        # Không phải "lãi/lỗ của đi vay/cho vay" nữa — đây là 4 hình thái nghiệp vụ
        # khác nhau: Buy trước + Sell>Buy = cho vay tiền (reverse repo lãi thường thấy);
        # Buy trước + Sell<Buy = vay bond (trả phí để mượn bond); Sell trước +
        # Sell>Buy = đi vay tiền (repo lãi/hiếm gặp); Sell trước + Sell<Buy = cho vay
        # bond (cho mượn bond, trả rebate interest trên tiền cọc — đây là chiều
        # THƯỜNG GẶP của "repo bán trước", không phải chi phí vay bất thường).
        if dirn == 'B':
            loai = 'Cho vay tien' if pnl > 0 else 'Vay bond'
        else:
            loai = 'Di vay tien' if pnl > 0 else 'Cho vay bond'
        matched.append(dict(
            paper=paper, cpty=cpty, prod=s['PRODUCT_CODE'], face=face,
            sid=s['DEAL_ID'], bid=b['DEAL_ID'], dirn=dirn, loai=loai,
            days=(second['VALUE_DATE'] - first['VALUE_DATE']).days,
            d1=first['VALUE_DATE'], d2=second['VALUE_DATE'],
            cash=first['GROSS_AMOUNT'], gap_cap=gap_cap, pnl=pnl,
            sy=s['YIELD'], by=b['YIELD'], st=s['TRANSACTION_STATUS'] + b['TRANSACTION_STATUS'],
            scap=s['CAPTURE_DATE'], bcap=b['CAPTURE_DATE'])
        )

    # Lượt 1 — ưu tiên ghép hai chân CÙNG NGÀY NHẬP MÁY trước: khi một nhóm có
    # nhiều hơn 2 chân, đây là cách chọn ĐÚNG cặp trong số nhiều khả năng (dấu
    # vết book cùng lúc là bằng chứng mạnh nhất về việc hai chân thuộc cùng một
    # hợp đồng). Sắp theo |lệch mã deal| ở đây hợp lý vì hai chân một hợp đồng
    # được nhập kề số hiệu. Đây chỉ là THỨ TỰ ƯU TIÊN GHÉP — không phải điều
    # kiện loại trừ, nên vẫn tính là Repo dù capture date lệch xa ở lượt sau.
    cands = sorted((abs(s['DEAL_ID'] - b['DEAL_ID']), i, j)
                   for i, s in enumerate(S) for j, b in enumerate(B)
                   if s['CAPTURE_DATE'] == b['CAPTURE_DATE'])
    for _, i, j in cands:
        if i in us or j in ub:
            continue
        us.add(i); ub.add(j)
        emit(S[i], B[j])

    # Lượt 2 — phần dư (không cùng ngày nhập máy). Ghép theo THỨ TỰ NGÀY THANH
    # TOÁN (FIFO), không theo mã deal: sắp theo |lệch mã deal| ở đây từng đẩy
    # một chân xa hẳn về cuối hàng đợi khi nhóm có ≥3 chân mỗi bên, tạo ra
    # "repo" dài giả tạo (ví dụ VPBCD080427/TCBV-HO: S 29/05 lẽ ra ghép với
    # B 12/06 cách 14 ngày, bị đẩy đi ghép với B 21/08 — thành 84 ngày). Khớp
    # theo thứ tự thời gian mô phỏng đúng một chu kỳ repo xoay vòng.
    # VẪN LÀ REPO — chỉ là BO nhập hai chân khác ngày, không loại trừ.
    rs = [s for i, s in enumerate(S) if i not in us]
    rb = [b for j, b in enumerate(B) if j not in ub]
    for s, b in zip(rs, rb):
        emit(s, b)

# Tách riêng để soát BO, KHÔNG dùng để loại trừ khỏi kết quả:
repo_same_cap = [p for p in matched if p['gap_cap'] == 0]
repo_diff_cap = [p for p in matched if p['gap_cap'] > 0]

left = [x for x in R if not any(x['DEAL_ID'] in (p['sid'], p['bid']) for p in matched)]

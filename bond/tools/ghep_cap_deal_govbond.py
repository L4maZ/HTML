#!/usr/bin/env python3
"""Ghép cặp repo từ export deal GovBond (RPBOD_B002 — TPCP/tín phiếu KBNN).

Thay cho export MSB_RP_DM cũ (`ghep_cap_deal_bond.py`) — RPBOD_B002 đầy đủ hơn nhưng có
một bẫy schema quan trọng: cột `Cpty_ShortName` KHÔNG PHẢI lúc nào cũng là đối tác ngoài
thật. `Folders_ShortName` mới là sổ (book) MSB dùng để book chân giao dịch. Theo lọc và
xác nhận của Jak (09/2026), chỉ 5 sổ AFS-* xuất hiện trong export gốc là hợp lệ để xét
tiếp (loại các rail nội bộ SBV/VSD-AVBL/SBV-A/SBV-L-*/AFS-ITB/AFS-LOANFI), và trong 5 sổ
đó có 2 sổ nội bộ — `AFS-GOV`, `AFS-ALM` — không tách bạch rõ ràng được đối tác ngoài nên
cũng bị loại nốt. **Quyết định**: chỉ giữ lại dòng có `Folders_ShortName` thuộc
{AFS-DCM, AFS-HUONG, AFS-HUY} trước khi ghép cặp — ba sổ "sạch" còn lại, mỗi giao dịch có
đối tác ngoài rõ ràng ở `Cpty_ShortName`, ghép đúng trong phạm vi cùng đối tác không cần
suy luận thêm.

Ghép cặp = repo/reverse repo: cùng mã TP (`Bonds_ShortName`) + cùng đối tác (`Cpty_ShortName`)
+ cùng coupon (`CouponRate`) + ngược chiều (S/B). Coupon suy ra được từ mã TP trên thực tế,
giữ trong khoá để mã nào lỡ trả nhiều coupon thì tự tách, không ghép nhầm.

CROSS_OK — MỘT ngoại lệ chéo đối tác đã xác nhận nghiệp vụ (MSB đứng trung gian mua KBNN
rồi bán lại qua một tổ chức khác cùng ngày book, xem `docs/bond-deal-pairing.md`):
    deal 50120 (B, Cpty=KBNN) ghép với deal 50127+50128 (S, Cpty=PGBV-HO)
    mã TD2636023, cùng CaptureDate 24/06/2026, cùng tổng khối lượng 6,000,000.
Bất kỳ cặp chéo đối tác nào KHÁC không có trong đây phải hỏi lại nghiệp vụ trước khi tính.

QUAN TRỌNG — thứ tự trước/sau và kỳ hạn dùng CAPTURE_DATE (ngày BO ghi nhận giao dịch),
KHÔNG dùng SettlementDate. Đây là khác biệt so với `ghep_cap_deal_bond.py` (MSB_RP_DM) và
`ghep_cap_deal_fibond.py` (dùng VALUE_DATE) — xác nhận riêng của Jak cho pipeline này
(09/2026): "logic là lấy capture date, ngày ghi nhận giao dịch". KHÔNG áp dụng ngược lại
cho hai pipeline kia trừ khi có yêu cầu riêng.

PHÂN LOẠI 4 NHÓM — giống hệt quy tắc đã chốt cho FIBond, theo (chân nào CaptureDate
TRƯỚC) x (Gross bên nào LỚN HƠN). pnl = Gross(S) − Gross(B) luôn luôn, bất kể chân nào
trước — xem chứng minh trong `emit()`.

    Buy trước, Sell sau, Gross(S) > Gross(B)  ->  Cho vay tien
    Buy trước, Sell sau, Gross(S) < Gross(B)  ->  Vay bond
    Sell trước, Buy sau, Gross(S) > Gross(B)  ->  Di vay tien
    Sell trước, Buy sau, Gross(S) < Gross(B)  ->  Cho vay bond

Kết quả kỳ 24/08/2026 (601 dòng gốc -> 374 dòng sau lọc Folders_ShortName): 198 cặp
Repo/Reverse Repo (197 cùng đối tác + 1 CROSS_OK), 0 chân dư. Phân bổ: 104 Cho vay bond,
93 Di vay tien, 1 Cho vay tien, 0 Vay bond.
"""
import datetime
from collections import defaultdict, Counter

import openpyxl

NCOLS = 25
SHEET = 'Sheet1'
F = '/home/user/HTML/bond/source/RPBOD_B002_2026.08.27.xlsx'

# Chỉ giữ 3 sổ "sạch" — đối tác ngoài rõ ràng ở Cpty_ShortName. Loại SBV/VSD-AVBL/
# SBV-A/SBV-L-*/AFS-ITB/AFS-LOANFI (rail nội bộ, không phải đối tác) VÀ AFS-GOV/AFS-ALM
# (hai sổ nội bộ không tách bạch được đối tác ngoài — xác nhận của Jak, 09/2026).
KEEP_FOLDER = {'AFS-DCM', 'AFS-HUONG', 'AFS-HUY'}

# Ngoại lệ chéo đối tác đã xác nhận nghiệp vụ (xem docstring). Khoá theo (bid, sid).
CROSS_OK = {
    (50120, '50127+50128'),
}


def parse_gross(v):
    if isinstance(v, str):
        return float(v.strip().rstrip('K')) * 1000.0
    return float(v)


def read_deals(path):
    ws = openpyxl.load_workbook(path, data_only=True)[SHEET]
    headers = [ws.cell(row=1, column=c).value for c in range(1, NCOLS + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        d = dict(zip(headers, [ws.cell(row=r, column=c).value for c in range(1, NCOLS + 1)]))
        if d.get('BondsDeals_Id') is None:
            continue
        d['GrossAmount'] = parse_gross(d['GrossAmount'])
        rows.append(d)
    return rows


R = read_deals(F)
R = [x for x in R if x['Folders_ShortName'] in KEEP_FOLDER]

g = defaultdict(lambda: {'S': [], 'B': []})
for x in R:
    g[(x['Bonds_ShortName'], x['Cpty_ShortName'], x['CouponRate'])][x['DealType']].append(x)

matched = []


def emit(s, b, qty, out):
    """Dựng một cặp từ chân bán s và chân mua b, khối lượng qty (đơn vị TP).

    Thứ tự trước/sau và kỳ hạn theo CAPTURE_DATE (xem docstring đầu file).
    pnl = Gross(S) - Gross(B) LUÔN LUÔN, bất kể chân nào thanh toán trước — hai nhánh
    dirn rút gọn về cùng biểu thức này (đã chứng minh trong pipeline FIBond, xem
    docs/bond-deal-pairing.md).
    """
    first, second = (s, b) if s['CaptureDate'] <= b['CaptureDate'] else (b, s)
    dirn = 'A' if first is s else 'B'   # A = Sell truoc, B = Buy truoc
    days = abs((b['CaptureDate'] - s['CaptureDate']).days)
    s_cash = s['GrossAmount'] * (qty / s['Quantity'])
    b_cash = b['GrossAmount'] * (qty / b['Quantity'])
    first_cash = s_cash if first is s else b_cash
    pnl = s_cash - b_cash
    if dirn == 'B':
        loai = 'Cho vay tien' if pnl > 0 else 'Vay bond'
    else:
        loai = 'Di vay tien' if pnl > 0 else 'Cho vay bond'
    # rate %/nam: chenh lech tien / tien chan dau / so ngay * 365 * 100.
    rate = ((b_cash - s_cash) / first_cash * 365 / days * 100) if days > 0 and first_cash else 0.0
    out.append(dict(
        paper=s['Bonds_ShortName'], cpty=s['Cpty_ShortName'],
        face=round(qty * s['FaceValue'] / 1e9, 4),
        sid=s['BondsDeals_Id'], bid=b['BondsDeals_Id'], dirn=dirn, loai=loai,
        days=days, d1=first['CaptureDate'], d2=second['CaptureDate'],
        sset=s['SettlementDate'], bset=b['SettlementDate'],
        cash=first_cash, pnl=pnl, rate=rate,
        sy=s['Yield'], by=b['Yield'], cpn=s['CouponRate'], dy=(b['Yield'] - s['Yield']) * 100,
        mat=s['MaturityDate'],
        sf=s['Folders_ShortName'], bf=b['Folders_ShortName'],
        sclr=s['ClearingModes_ShortName'], bclr=b['ClearingModes_ShortName'],
        scap=s['CaptureDate'], bcap=b['CaptureDate'],
    ))


for (paper, cpty, cpn), v in g.items():
    S = sorted(v['S'], key=lambda x: (x['CaptureDate'], x['BondsDeals_Id']))
    B = sorted(v['B'], key=lambda x: (x['CaptureDate'], x['BondsDeals_Id']))
    sq = [x['Quantity'] for x in S]
    bq = [x['Quantity'] for x in B]

    # Lượt 1 — ưu tiên ghép cùng CaptureDate, |lệch mã deal| gần nhất (dấu vết book).
    cands = sorted((abs(s['BondsDeals_Id'] - b['BondsDeals_Id']), i, j)
                   for i, s in enumerate(S) for j, b in enumerate(B)
                   if s['CaptureDate'] == b['CaptureDate'])
    for _, i, j in cands:
        if sq[i] <= 1e-6 or bq[j] <= 1e-6:
            continue
        qty = min(sq[i], bq[j])
        sq[i] -= qty
        bq[j] -= qty
        emit(S[i], B[j], qty, matched)

    # Lượt 2 — FIFO phần dư theo thứ tự CaptureDate.
    si = bi = 0
    while si < len(S) and bi < len(B):
        if sq[si] <= 1e-6:
            si += 1
            continue
        if bq[bi] <= 1e-6:
            bi += 1
            continue
        qty = min(sq[si], bq[bi])
        sq[si] -= qty
        bq[bi] -= qty
        emit(S[si], B[bi], qty, matched)

byid = {x['BondsDeals_Id']: x for x in R}
_matched_ids = {p['sid'] for p in matched} | {p['bid'] for p in matched}

# Áp CROSS_OK: ghép nốt chân KBNN (B) với PGBV-HO (S) chưa dùng tới.
for bid, sid_group in CROSS_OK:
    sids = [int(x) for x in sid_group.split('+')]
    b = byid[bid]
    slegs = [byid[i] for i in sids]
    qty = sum(x['Quantity'] for x in slegs)
    if b['Quantity'] != qty:
        continue
    s0 = slegs[0]
    fake_s = dict(s0)
    fake_s['BondsDeals_Id'] = sid_group
    fake_s['Quantity'] = qty
    fake_s['GrossAmount'] = sum(x['GrossAmount'] for x in slegs)
    emit(fake_s, b, qty, matched)
    _matched_ids |= set(sids) | {bid}

left = [x for x in R if x['BondsDeals_Id'] not in _matched_ids]

if __name__ == '__main__':
    print(f'Tong dong sau loc Folders_ShortName: {len(R)}')
    print(f'So cap Repo/Reverse Repo: {len(matched)}')
    print(f'Phan bo Loai: {Counter(p["loai"] for p in matched)}')
    print(f'Chan du: {len(left)}')

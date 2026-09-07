#!/usr/bin/env python3
"""Tổng hợp D (JSON bơm vào HTML) cho GovBond, 4 nhóm phân loại thay 2 nhóm A/B cũ.

Logic tổng hợp (blk/group_by/tenor/month/cpty/bond/folder/outstanding) chuyển thể trực
tiếp từ `ghep_cap_deal_bond.py::build_D()` — chỉ đổi trục nhóm từ (A,B) sang 4 khoá
(CVT,VB,DVT,CVB) tương ứng loai = Cho vay tien / Vay bond / Di vay tien / Cho vay bond.

Đơn vị hiển thị: cash/face tỷ VND, pnl triệu VND — quy đổi tại chỗ từ VND thô do
`ghep_cap_deal_govbond.py` trả về (đã có cash=tiền chân đầu, pnl=Gross(S)-Gross(B), cả
hai đều VND thô).

folder(ps): dùng folder của CHÂN ĐẦU — sf nếu dirn=='A' (Sell trước), bf nếu dirn=='B'
(Buy trước) — cùng quy ước dirn với `ghep_cap_deal_bond.py` nên tái dùng y hệt.
"""
import datetime
import json
import statistics
from collections import defaultdict

from ghep_cap_deal_govbond import matched, left, R

# Thứ tự này quyết định thứ tự hiển thị: hai nhóm bán-trước (repo) rồi hai nhóm
# mua-trước (reverse repo).
GCODES = ['VT', 'CVB', 'CVT', 'VB']
GROUP_CODE = {'Vay tien': 'VT', 'Cho vay bond': 'CVB', 'Cho vay tien': 'CVT', 'Vay bond': 'VB'}
GROUP_LABEL = {v: k for k, v in GROUP_CODE.items()}


def scale(p):
    """Bản sao p với cash (tỷ VND) và pnl (triệu VND) quy đổi từ VND thô."""
    q = dict(p)
    q['cash'] = p['cash'] / 1e9
    q['pnl'] = p['pnl'] / 1e6
    q['d1'] = p['d1'].strftime('%Y-%m-%d')   # settlement chân đầu
    q['d2'] = p['d2'].strftime('%Y-%m-%d')   # settlement chân sau
    return q


PAIRS = [scale(p) for p in matched]
for p, orig in zip(PAIRS, matched):
    p['sset'] = orig['sset'].strftime('%Y-%m-%d')
    p['bset'] = orig['bset'].strftime('%Y-%m-%d')
    p['scap'] = orig['scap'].strftime('%Y-%m-%d')
    p['bcap'] = orig['bcap'].strftime('%Y-%m-%d')
    p['mat'] = orig['mat'].strftime('%Y-%m-%d')
    p['grp'] = GROUP_CODE[orig['loai']]


def dt(s):
    return datetime.datetime.strptime(s, '%Y-%m-%d')


def pnl_mn(p):
    """Lãi/lỗ hiển thị, triệu VND. Đã ở đơn vị này (xem scale())."""
    return p['pnl']


def wrate(ps):
    """Lãi suất ngụ ý %/năm, trọng số theo tiền x ngày — CHỈ trên cặp có ngày>0.

    Cặp ngày=0 (đa số ở pipeline này, xem docstring đầu file) không có mẫu số
    (cash*0/365=0) nên phải loại khỏi CẢ tử số lẫn mẫu số, không chỉ mẫu số —
    nếu không, pnl của các cặp 0-ngày dồn hết lên vài cặp có ngày>0 hiếm hoi,
    ra %/năm vô lý (đã thấy CVB: -15625%/năm khi còn bug).
    """
    ps2 = [p for p in ps if p['days'] > 0]
    den = sum(p['cash'] * p['days'] / 365 for p in ps2)
    return (sum(pnl_mn(p) / 1e3 for p in ps2) / den * 100) if den else 0.0


def blk(ps):
    return {
        'n': len(ps),
        'cash': round(sum(p['cash'] for p in ps), 3),
        'pnl': round(sum(pnl_mn(p) for p in ps), 2),
        'rate': wrate(ps),
        'med': statistics.median([p['days'] for p in ps]) if ps else 0,
        'minD': min((p['days'] for p in ps), default=0),
        'maxD': max((p['days'] for p in ps), default=0),
        'face': round(sum(p['face'] for p in ps), 2),
        'nLai': sum(1 for p in ps if pnl_mn(p) > 0),
        'nLo': sum(1 for p in ps if pnl_mn(p) < 0),
    }


def group_by(ps, key, extra=None, top=None):
    out = []
    g = defaultdict(list)
    for p in ps:
        g[key(p)].append(p)
    for k, v in g.items():
        b = blk(v)
        b['k'] = k
        if extra:
            b.update(extra(v))
        out.append(b)
    out.sort(key=lambda x: -x['cash'])
    return out[:top] if top else out


def outstanding(ps):
    ev = defaultdict(float)
    for p in ps:
        a, z = sorted([dt(p['sset']), dt(p['bset'])])
        ev[a] += p['cash']
        ev[z] -= p['cash']
    if not ev:
        return []
    cur, out, d, end = 0.0, [], min(ev), max(ev)
    while d <= end:
        cur += ev.get(d, 0.0)
        out.append([d.strftime('%Y-%m-%d'), round(cur, 1)])
        d += datetime.timedelta(days=1)
    return out


def tenor(ps):
    out = []
    for k, lo, hi in [('1 ngày', 1, 1), ('2–7 ngày', 2, 7), ('8–30 ngày', 8, 30),
                      ('31–90 ngày', 31, 90), ('> 90 ngày', 91, 10 ** 9)]:
        v = [p for p in ps if lo <= p['days'] <= hi]
        if v:
            b = blk(v)
            b['k'] = k
            out.append(b)
    return out


def month(ps):
    """Gom theo tháng của chân THANH TOÁN TRƯỚC — thời điểm dòng tiền bắt đầu."""
    g = defaultdict(list)
    for p in ps:
        g[dt(p['d1']).strftime('%m/%Y')].append(p)
    out = []
    for m in sorted(g, key=lambda x: (x[3:], x[:2])):
        b = blk(g[m])
        b['m'] = m
        out.append(b)
    return out


def cpty(ps):
    r = group_by(ps, lambda p: p['cpty'])
    for x in r:
        x['cpty'] = x.pop('k')
    return r


def bond(ps):
    r = group_by(ps, lambda p: p['paper'],
                 lambda v: {'mat': v[0]['mat'], 'cpn': v[0]['cpn']}, top=14)
    for x in r:
        x['b'] = x.pop('k')
    return r


def folder(ps):
    r = group_by(ps, lambda p: p['sf'] if p['dirn'] == 'A' else p['bf'])
    for x in r:
        x['f'] = x.pop('k')
    return r


groups = {}
for code in GCODES:
    ps = [p for p in PAIRS if p['grp'] == code]
    groups[code] = {
        'cpty': cpty(ps), 'tenor': tenor(ps), 'month': month(ps),
        'bond': bond(ps), 'fold': folder(ps), 'outst': outstanding(ps),
        'hist': [round(p['rate'], 3) for p in ps if p['days'] > 0],
    }

kpi_groups = {code: blk([p for p in PAIRS if p['grp'] == code]) for code in GCODES}

anom, noise = [], []
for p in PAIRS:
    if abs(p['dy']) >= 10 and p['days'] <= 30:
        q = dict(p)
        q['why'] = 'Δyield ≥ 10bp, kỳ hạn ≤ 30 ngày'
        anom.append(q)
    elif p['days'] >= 3 and abs(p['rate']) > 15:
        q = dict(p)
        q['why'] = '|%/năm| > 15 trên kỳ hạn ≥ 3 ngày'
        anom.append(q)
    elif p['days'] <= 2 and abs(p['rate']) > 8:
        q = dict(p)
        q['why'] = '|%/năm| > 8 trên kỳ hạn 1–2 ngày'
        noise.append(q)
anom.sort(key=lambda x: -abs(x['pnl']))
noise.sort(key=lambda x: -abs(x['pnl']))

caps = [x['CaptureDate'] for x in R]
d_from, d_to = min(caps).strftime('%Y-%m-%d'), max(caps).strftime('%Y-%m-%d')

D = {
    'kpi': {'nDeals': len(R), 'nPairs': len(PAIRS), 'nUn': len(left),
            'groups': kpi_groups, 'dFrom': d_from, 'dTo': d_to},
    'groups': groups,
    'anom': anom, 'noise': noise,
    'fmis': [p for p in PAIRS if p['sf'] != p['bf']],
    'cmis': [p for p in PAIRS if p['sclr'] != p['bclr']],
    # Cặp có lãi/lỗ đúng bằng 0: lãi repo ngụ ý nhỏ hơn bước làm tròn của GrossAmount
    # (10.000đ) nên biến mất. Liệt kê ra để soát, xếp nhóm theo cấu trúc (Cho vay bond).
    'zero': [p for p in PAIRS if p['pnl'] == 0],
    'pairs': PAIRS,
    'unmatched': [{'id': x['BondsDeals_Id'], 'bond': x['Bonds_ShortName'], 'cpty': x['Cpty_ShortName'],
                   'cap': x['CaptureDate'].strftime('%Y-%m-%d'), 'qty': x['Quantity']} for x in left],
}

if __name__ == '__main__':
    print('nPairs:', D['kpi']['nPairs'], 'nDeals:', D['kpi']['nDeals'], 'nUn:', D['kpi']['nUn'])
    for code, b in kpi_groups.items():
        print(f"  {GROUP_LABEL[code]:14s} ({code}): {b['n']:3d} cặp · {b['cash']:>10,.1f} tỷ · "
              f"trung vị {b['med']:>5.0f} ngày · {b['pnl']:>+11,.1f} tr · {b['rate']:7.2f}%/năm")
    print('anom:', len(D['anom']), 'noise:', len(D['noise']), 'fmis:', len(D['fmis']), 'cmis:', len(D['cmis']))
    blob = json.dumps(D, ensure_ascii=False, separators=(',', ':'))
    print(f'blob size: {len(blob):,} bytes')

#!/usr/bin/env python3
"""Dựng file HTML báo cáo GovBond từ template + D (build_D_govbond) + ECharts blob.

    python3 render_govbond_report.py [out.html]

Template `govbond_report_template.html` chứa CSS/HTML/JS render (4 nhóm phân loại,
2 trang "Buy trước"/"Sell trước" mỗi trang 2 cột con) với hai placeholder:
`%%ECHARTS%%` (thay bằng `echarts_5.4.3_blob.html`, bundle ECharts 5.4.3 minify —
copy nguyên khối từ `BaoCao_RRTT_Bond_key_v7.html`, offline, không CDN) và
`%%D_JSON%%` (thay bằng `const D=...;` build từ `build_D_govbond.py`).
"""
import json
import sys

from build_D_govbond import D

OUT = sys.argv[1] if len(sys.argv) > 1 else \
    f"/home/user/HTML/bond/Phan_tich_GD_Bond_{D['kpi']['dFrom'].replace('-', '')}_{D['kpi']['dTo'].replace('-', '')}.html"

with open('govbond_report_template.html', encoding='utf-8') as f:
    tpl = f.read()
with open('echarts_5.4.3_blob.html', encoding='utf-8') as f:
    echarts_blob = f.read()

blob = json.dumps(D, ensure_ascii=False, separators=(',', ':'))
out = tpl.replace('%%ECHARTS%%', echarts_blob).replace('%%D_JSON%%', 'const D=' + blob + ';')

with open(OUT, 'w', encoding='utf-8') as f:
    f.write(out)
print(f'→ {OUT}  ({len(out):,} bytes)')

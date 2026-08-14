# Trạng thái các dashboard

## Bond Portfolio Risk Management Report (File 02 HTML, v11+)

Có **key system riêng**, không dùng chuẩn chung.

- 92 keys qua `RRTT_setKeys` / `RRTT_keyList`
- Data lưu trong `src2.js`; render trong `src3.js`
- Wine red `#7B2D3B` cho active nav / sub-tab / table header
- Toggle MSB/SBV **chỉ** ở Indicator sub-tab (biến `SCOPE`)
- Key sheet: bảng phẳng ~2,116 rows trong File 02 (persistent file)
- VBA macro `XuatFileKey` export Key+Value dạng values-only ra `Key_YYYYMMDD.xlsx`
- 1,530 dòng công thức working đã validate; **29 dòng chưa linked**

## Peer Bond Dashboard

- Bản mới nhất: `Peer_Bond_Dashboard_v5.1.html`
- 27 ngân hàng; 4 tab: Overview / Cross-sectional / Time Series / Data Table
- Style palette chuẩn: `#b83a10 → #e8551a`, bg `#fdf3f0`, card `#fff7f4`
- Bank pills multi-select với All / Clear / Top10
- ECharts 5.4.3 inline

## VBMA Dashboard (August 2026)

- Wine-red `#8B2332`
- 6 tab: Overview / Money Market / FX / Primary TPCP / Secondary TPCP / Corporate Bonds
- Tab Data Validation **đã xóa** — discrepancy chỉ hiện qua footnote asterisk inline
- Nội dung: Top 10 outright trading table; Q2/2026 issuance plan (tr.8); Annex 2 (33 corporate
  bond records); regional 5Y yield comparison (7 nước, 25 tháng); monthly corporate bond
  maturity by sector; WoW/MoM interbank chart có dual-axis fix

## Macro knowledge-sharing session tools

- `du_bao_macro.html` — game numeric forecasting với deviation scoring. **Đã build.**
- "Đặt cược dự báo" — logged, **chưa build**
- "Phe Bull vs Phe Bear" — logged, **chưa build**

# Bond — Báo cáo rủi ro Desk Bond (Tool → File Key → HTML)

Báo cáo Desk Bond hằng ngày của QLRR Thị trường, gửi ĐVKD. Một file HTML duy nhất, mở offline.
Dữ liệu đi từ File 02 (Excel) qua một **file key** nhỏ, nạp vào HTML ở tab Quản trị dữ liệu.

## File trong thư mục

| File | Vai trò |
|---|---|
| `BaoCao_RRTT_Bond_v8.html` | **Báo cáo chính (~1.2MB, mẫu rỗng).** Chưa nạp key thì hiện "Chưa có dữ liệu". Có ECharts 5.4.3 nhúng inline |
| `Key_20260929.xlsx` | File key ngày 29/09/2026 do macro trên máy Jak sinh ra (schema `bond.key.v3`) |
| `Key_2026081x.xlsx` | Key mẫu cũ (13/14/17/08), chỉ để tham khảo |
| `tools/XuatFileKey.bas` | Macro VBA sinh file key từ File 02 trên máy bank |
| `tools/make_key.py` | Bản Python tương đương (dùng khi có Python) |
| `tools/admin_panel.js` | Mã tab Quản trị dữ liệu (đã ghép vào v8) |
| `ISSUES.md` | Nhật ký lỗi, nguyên nhân gốc, kết quả đối chiếu số, việc tồn |
| `TEN_CHI_SO_DE_XUAT.md` | Bảng đề xuất tên đầy đủ chỉ số (đợt 5) |
| `BaoCao_RRTT_Bond_v8_upload.html` | Bản thử nghiệm cũ (chỉ khối 2), không dùng nữa |

## Quy trình hằng ngày

```
File 02 (.xlsm)  ──macro XuatFileKey──►  Key_YYYYMMDD.xlsx
                                             │  Quản trị dữ liệu → nạp file key
                                             ▼
                        BaoCao_RRTT_Bond_v8.html  (điền RPT/TS → vẽ lại)
                                             │  Lưu đè vào file này  (hoặc tải bản mới)
                                             ▼
                        file báo cáo có dữ liệu ngày chốt, gửi mail
```

1. Mở File 02, chạy macro `XuatFileKey` (cài một lần: Alt+F11 → Import `tools/XuatFileKey.bas`,
   chạy `TaoCotKeyID` một lần).
2. Mở HTML → **Quản trị dữ liệu** (khoá mật khẩu, chỉ chặn bấm nhầm) → nạp file key.
3. Xem **Kiểm tra trước khi gửi** (chỉ hiện khi có lỗi) → bấm **Lưu đè vào file này**.
   Trình duyệt không hỗ trợ thì file mới được tải về.

Mở file lại không nạp key thì báo cáo trống. Dữ liệu nằm trong file đã lưu, không nằm trong `localStorage`.

## Nội dung báo cáo

| Trang | Nội dung |
|---|---|
| Highlight | KPI tổng quan |
| Chi tiết danh mục | Indicator (hạn mức, đèn) · Trạng thái (kỳ hạn, tổ chức phát hành, thời gian nắm giữ) · Độ nhạy PV01 · ItD Unrealized PnL · Lãi/lỗ · Sử dụng vốn |
| QTRR theo kịch bản | VaR / kịch bản yield |
| Thông tin thị trường | Yield curve, repo, thanh khoản, biến động |
| FI Bond & CD | Cơ cấu danh mục, rating, PnL theo QLHS, lịch sử sơ cấp/thứ cấp |
| Quản trị dữ liệu | Nạp key, trạng thái, kiểm tra, nhật ký |

### Ánh xạ file key → window.RPT

| Sheet | Nhánh RPT |
|---|---|
| `META` | `asOf`, `dates`, `colDates` |
| `DATA_S2` | `books.MSB[]`, `books.SBV[]`, `other[]`, `fibond[]` |
| `POS` | `posTrading[]`, `posBanking[]`, hai `*Total` |
| `CURVE` | `yieldCurve[]`, `repoCurve[]` |
| `GRID` | `issuerMix`, `holdTime`, `pnlBreakdown`, `capital`, `bsBook`/`bsLayers`, `fiOnBS`, `pnlScenario` |
| `SCEN` | `scenTB.recent[]` / `.varScen[]` (và `scenBB` nếu có) |
| `VIRA` | `vira.scenarios[]`, `vira.books[]` |
| `RATING` | `fiRating[]`, `fiRatingTotal` |
| `VOL` | `volatility` — tính SMA 252/504/756 + EWMA λ=0.98, annualize ×√252 |
| `TEXT` | `highlight.*`, `reportMarket`, `fvNote`, `fiNote`, `note10d`, `noteVar` |
| `TS` | 17 chuỗi trong `window.TS` |

`VOL` chở delta yield thô 800 ngày, HTML tự tính độ biến động — nhờ vậy λ và cửa sổ
252/504/756 thành tham số Tier 2, sửa được mà không phải đụng Excel.

## Những gì đã làm (29–30/09/2026)

Nguồn: biên bản cuộc họp + góp ý của Jak, làm theo từng đợt, mỗi đợt đối chiếu số với File 02.

**Giao diện theo góp ý cuộc họp**
- Nền bảng xám ấm, header bảng đậm hơn; header báo cáo thêm *Đơn vị báo cáo · Đối tượng nhận · Tần suất*.
- Tổ chức phát hành và thời gian nắm giữ chuyển vào tab Trạng thái; mục 2.x FI Bond & CD chuyển sang trang FI Bond & CD.
- Bỏ Nhận định thị trường và các comment thừa; đổi tên đầy đủ *ItD Unrealized PnL*.
- Hai bảng FI đặt song song để không phải kéo ngang; không còn bảng nào tràn ở 1100–1500px.
- Ngày thống nhất `dd/mm/yyyy`; ẩn Volatility; yield TPCP tối đa 2 số lẻ; bỏ cột PV01/1,000 tỷ.
- Cột T-1 và cuối tháng trùng nhau thì chỉ hiện một cột; cột ngày báo cáo ở Indicator nền xám đậm, in đậm.
- Đóng/mở nhóm chỉ ở 2 bảng Indicator (Trạng thái mở, Vốn yêu cầu, PnL) + nút *Mở hết / Đóng hết* ở thanh trái.
- Kiểm tra trước khi gửi chỉ hiện khi có lỗi.

**Sửa lỗi số liệu (quan trọng nhất)**
- Nguyên nhân: sheet `Linked` lệch 1 dòng so với toạ độ cứng → Trạng thái, Hold, PV01, PnL, Vốn sai sau khi nạp key.
  Sửa bằng cách định vị theo **tiêu đề khối** trong `make_key.py` và `XuatFileKey.bas`.
- Excel đảo ngày dd/mm ↔ mm/dd (ngày ≤ 12): macro ghi cột ngày dạng chữ; HTML cảnh báo khi trục ngày không đơn điệu.
- Loader chịu lỗi: không lấp số cũ khi thiếu trường, có ALIAS tên trường, báo lỗi cấu trúc.
- Sửa dòng lặp *Giảm trừ đối ứng*, biểu đồ Sử dụng vốn đếm đôi ItD, nhãn trục bị cắt.

**Lưu dữ liệu**
- Bỏ dữ liệu nhúng sẵn (không làm hỏng gì: khung rỗng vẫn chạy đủ).
- Nút *Lưu đè vào file này* (File System Access API). Bản export không phình sau mỗi lần lưu.

**Hiệu ứng UX (v10)**
- Từ skill html-dashboard-factory, chỉ lấy lớp trình bày: motion tokens, hover hàng bảng, hiện dần khi vào tab lần đầu,
  View Transitions khi chuyển tab và đóng/mở nhóm, tab dùng được bằng bàn phím, tắt hiệu ứng khi in / giảm chuyển động.
- Không làm: thanh hạn mức, popover giải nghĩa, dark mode, số chạy nhảy.

**Kết quả đối chiếu với File 02 (29/09)**
- 731 ô số trên RPT, 460 ô trên DOM, 2,428 điểm chuỗi thời gian: **0 lệch**. Đúng với cả key do Python sinh lẫn key do macro sinh.
- Ngoại lệ đã biết: `lsttPrimary` có 5 ngày cuối trống trong Excel.

## Việc chưa làm / để sau

- **Đợt 6**: rule cảnh báo biến động so với ngày hôm trước, và chọn ngày so sánh — hoãn theo yêu cầu.
- Chưa thử hộp thoại lưu thật trên Edge/Chrome (mới thử bằng stub); chưa thử đọc `.xlsx` trên máy bank.
- `docs/projects.md` và `CLAUDE.md` vẫn mô tả bản v7.

## Nguyên tắc cài trong HTML

- Thiếu số thì hiện `—`, không lấy số kỳ trước; thiếu key text thì về chữ gốc, không bao giờ `undefined`.
- Mọi chart qua registry `REG` + `ResizeObserver`.
- File key đọc ngay trong trình duyệt (`DecompressionStream`), không thư viện ngoài, không gửi đi đâu.
- Giới hạn: file gửi đi ≤ 5MB (mail bank chặn quá mức này); bản có dữ liệu hiện ~1.8MB.

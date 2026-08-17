# Bond — Tool → File Key → HTML (khối 2)

Bản thử nghiệm cơ chế: File 02 sinh ra một file key nhỏ, HTML nạp file key đó và render.
Phạm vi hiện tại **chỉ khối 2** (52 chỉ tiêu tuân thủ hạn mức) để kiểm tra tính năng.

## File trong thư mục

| File | Vai trò |
|---|---|
| `BaoCao_RRTT_Bond_v8_upload.html` | HTML nạp file key, render khối 2, xuất bản gửi đi. **22KB** |
| `Key_20260813.xlsx` | File key mẫu, sinh từ File 02 ngày 13/08. **16KB** |
| `tools/XuatFileKey.bas` | Macro VBA để tự sinh file key trên máy bank |
| `tools/make_key.py` | Bản Python tương đương, dùng khi có Python |

## Quy trình

```
File 02 (.xlsm)
  ├─ Linked (1)  A1:N80   → 52 chỉ tiêu khối 2
  ├─ Report      6 ô chữ  → nhận định, ghi chú
  └─ Run Tool    B2:B7    → trục ngày
        │  XuatFileKey
        ▼
  Key_YYYYMMDD.xlsx  (META · DATA_S2 · TEXT) — 16KB
        │  nạp qua nút trong HTML, parse tại chỗ
        ▼
  BaoCao_RRTT_Bond_v8_upload.html
        │  Xuất bản gửi đi
        ▼
  BaoCao_RRTT_Bond_DDMMYYYY.html — 70KB, gửi mail được
```

## Cài macro (một lần)

1. Mở File 02 → `Alt+F11` → chuột phải VBAProject → **Import File** → chọn
   `tools/XuatFileKey.bas`.
2. Chạy `TaoCotKeyID` **một lần**. Macro ghi cột `KeyID` vào `Linked (1)` cột **N**.
3. Từ đó về sau mỗi kỳ chỉ chạy `XuatFileKey`. File key nằm cùng thư mục File 02.

Chèn thêm dòng trong `Linked (1)` thì chạy lại `TaoCotKeyID` để cấp key cho dòng mới.

## Vì sao có cột KeyID

`KeyID` là **neo**. HTML tra cứu theo chuỗi key, không theo toạ độ ô. Chèn/xoá/đảo dòng
trong Excel không làm HTML đọc lệch — dòng nào không có key thì không ra báo cáo, thay vì
âm thầm lấy số của dòng bên cạnh.

## Nguyên tắc đã cài trong HTML

- **Thiếu số thì gãy to.** Ô trống hiện `—` đỏ, không bao giờ lấy số kỳ trước.
- **Banner ngày chốt** luôn hiện, kèm tên file nguồn và giờ sinh file.
- **Kiểm tra cấu trúc**: sai `schema`, trùng `KeyID`, thiếu sheet → banner đỏ liệt kê lỗi.
- **Offline hoàn toàn.** File key được parse ngay trong trình duyệt bằng `DecompressionStream`,
  không thư viện ngoài, không gửi đi đâu. Google Fonts chặn thì degrade về Segoe UI.
- **Không `localStorage`.** Bản xuất mang dữ liệu trong `window.RRTT_SNAPSHOT`.
- Key text vẫn theo cơ chế `data-k` + `KEY_DEF`: xoá sạch key thì chữ gốc quay lại.

## Đã kiểm

Chromium headless, mở bằng `file://`, ngắt mạng:

- Nạp `Key_20260813.xlsx` → 65 dòng, 7 section, 18 đèn hạn mức.
- Face Value TB nội bộ 17,513 = AFS 12,646 + HTM 4,867 — khớp `Linked (1)`.
- Hạn mức đầu tư ra nước ngoài 15.0/15m USD → đèn **đỏ 99.9%**.
- Ô rỗng trong Excel → `—`, không nội suy.
- Xuất bản gửi đi → 70KB, mở lại đủ 65 dòng, ẩn thanh công cụ, trạng thái "chỉ đọc".
- Google Fonts bị chặn (`ERR_CONNECTION_RESET`) → trang vẫn dựng bình thường.

## Giới hạn tuyệt đối

**File HTML gửi đi ≤ 5MB** (mail bank chặn quá 5MB). Hiện 70KB. Khi thêm ECharts inline
cho các khối 3.x/4/6.x sẽ tăng ~940KB — vẫn còn dư địa lớn.

## Chưa làm

- Khối 3.x · 4 · 6.x · 7.x (đọc từ sheet `Linked`, cần ECharts)
- 19 chuỗi time-series từ `Chart data`
- Block "Bình luận / Nhận định" chuẩn `mkCmt*`

## Tab "Quản trị dữ liệu"

Toàn bộ phần nạp/xuất chuyển sang tab riêng, tab "Báo cáo" chỉ còn nội dung gửi đi.
Số trên chip cạnh tên tab = số mục cần xem.

| Mục | Nội dung |
|---|---|
| A · Nạp dữ liệu | Vùng kéo thả, nút xuất bản, nút xoá dữ liệu |
| B · Trạng thái | Ngày chốt · số chỉ tiêu · số cảnh báo · file nguồn · giờ sinh file |
| C · Kiểm tra trước khi gửi | Tuổi dữ liệu · ô trống · đối chiếu tổng vs dòng con · chỉ tiêu chạm hạn mức |
| D · Nhật ký nạp | Mỗi lần nạp ghi một dòng, đi theo file khi xuất bản |
| E · Môi trường | Trình duyệt có đọc được .xlsx không — dùng khi test máy mới |

### Vòng đời dữ liệu

Mở file → hiện dữ liệu kỳ trước đã nhúng sẵn, kèm cảnh báo tuổi nếu quá 3 ngày.
Nạp file mới → thay toàn bộ. Nạp lỗi → giữ nguyên dữ liệu cũ, báo đỏ.
Xoá dữ liệu → báo cáo trống, ghi một dòng vào nhật ký.

### Kiểm tra nhất quán đang chạy

Đối chiếu tổng với các dòng con, ngay khi nạp:

- Trading nội bộ: Face Value = AFS + HTM · ItD = AFS + HTM
- Banking nội bộ: Face Value = DCM tự fund + fund từ Pool · ItD = tương tự

Với file 13/08, hai dòng Banking báo đỏ: Face Value lệch **6,101 tỷ**, ItD lệch **716 tỷ**.
Đúng vấn đề đã nêu ở `Linked (1)` — tổng lấy SUMIFS còn hai dòng con là số hardcode.

## v8 — ghép tab Quản trị vào chính file v7

`BaoCao_RRTT_Bond_v8.html` = v7 nguyên vẹn + một khối script chèn cuối `<body>`.
Không sửa ruột file gốc: tab và panel được dựng bằng DOM sau khi `RRTT.mount()` chạy xong.

- Rail trái thêm mục **Quản trị dữ liệu** (`p6`), khoá bằng mật khẩu.
- Nạp file key → ánh xạ vào `window.RPT` + `window.TS` → `RRTT.rerender()`.
- Mỗi trang p1–p5 gắn badge ngày chốt, quá 3 ngày thì badge vàng.
- Bản xuất giữ nguyên tab quản trị (vẫn khoá) để kỳ sau nạp tiếp.

Mật khẩu nằm trong mã nguồn nên **chỉ chặn bấm nhầm, không phải bảo mật**.

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

### Đã kiểm (Chromium, `file://`, ngắt mạng)

- Nạp key 13/08: `asOf` 24/07 → 13/08 · Face TB 17,513 · posTrading 10 kỳ hạn ·
  yield 10Y mid 4.423 · scen recent 6 / VaR 6 · rating 20 issuer · `TS.yieldTs` 195 điểm.
- Mật khẩu sai → không mở được tab.
- Xuất bản: 1.92MB (trần 5MB), mở lại giữ đúng 13/08, chart vẫn vẽ.

## v8 — bản giao kèm VBA (17/08)

- Thanh công cụ sửa **nổi cố định dưới màn hình**, theo anh qua cả 5 trang; có Ctrl+Z / Ctrl+Y.
- Bỏ badge ngày chốt khỏi p1–p5; ngày chốt + "đã cũ n ngày" chuyển vào tab Quản trị.
- Thêm `VIRA4` (5 dòng, từ `VIRA scenarios!A4:F8`) → biểu đồ dự báo VIRA vs thực tế.
- `tools/XuatFileKey.bas` sinh đủ 12 sheet, cùng cấu trúc bản Python.

### Macro — hai điểm thiết kế

**Dò theo nhãn, không theo toạ độ.** `DATA_S2` tìm dòng chứa `"TRADING BOOK NỘI BỘ"`,
`"BANKING BOOK NỘI BỘ"`… rồi đọc tới nhãn kế tiếp. Chèn/xoá dòng trong `Linked (1)` không
làm lệch. Không thấy nhãn thì macro dừng và báo tên nhãn thiếu, không xuất file sai.

**Named Range cho ô nhận định.** Ô trong `Report` dịch lên xuống theo độ dài bảng, nên đọc
theo địa chỉ cứng sẽ lấy nhầm ô. Chạy `TaoNameNhanDinh` một lần để đặt 17 name `HTML_*`;
Excel tự cập nhật vùng khi dòng dịch.

**`TS` ghi bằng mảng một lần** (`Range.Resize(n,4).Value = arr`) thay vì gán từng ô.

### Cài đặt

1. `Alt+F11` → Import `tools/XuatFileKey.bas`
2. Chạy `TaoNameNhanDinh` — **một lần duy nhất**
3. Mỗi kỳ: chạy `XuatFileKey` → `Key_YYYYMMDD.xlsx` nằm cùng thư mục File 02

## Sửa trình bày (17/08, sau rà soát của Jak)

Ba trang bị tràn ngang phải kéo touchpad — nguyên nhân chung: **thẻ nội dung là grid item
nhưng thiếu `min-width:0`**, nên nó nở theo bề rộng bảng thay vì bó lại. 24 thẻ đã được sửa.
Bảng rộng giờ cuộn **trong thẻ** (`min-width:max-content` + `overflow-x:auto`), kèm dòng nhắc
"kéo ngang trong bảng" chỉ hiện khi bảng thật sự rộng hơn khung.

| Chỗ | Trước | Sau |
|---|---|---|
| Chi tiết danh mục · Indicator | tràn 1.303px ở màn 1.280 | vừa khung |
| Chi tiết danh mục · Lãi/lỗ | tràn 1.902px | vừa khung |
| QTRR theo kịch bản | tràn 1.559px | vừa khung |

Kiểm ở 1.280 · 1.366 · 1.440px, 5 trang × 14 tab con: `scrollWidth == clientWidth` toàn bộ.

**Chart dẹt như đường thẳng** — trục giá trị ECharts mặc định kéo về mốc 0, nên dải yield
4.18–4.42% bị vẽ trên thang 0–5. Thêm `scale:true` cho `ch_vira4`, `ch_yield`, `ch_yieldTs`,
`ch_spread`, `ch_repoTs`; các chart cột vẫn giữ mốc 0. Bốn chart dạng đường được bó
`max-width:880px` và tăng chiều cao để bớt bè ngang.

**Tiêu đề cột dài** (`KB1: GIỮ DM ĐẾN KHI ĐÁO HẠN REPO`, `DỰ BÁO BÌNH QUÂN CỦA CÁC MARKET
MAKER`) giờ xuống dòng thay vì kéo bảng ra ngoài màn hình. Bảng VIRA nhờ vậy hiện đủ 4 kịch bản.

**Dòng "Giả định biến động yield"** đổi thành dải chân bảng — nền xám nhạt, chữ nhỏ in hoa,
gộp 3 cột đầu — thay vì một dòng dữ liệu trống hai ô.

**Màu ô chú giải lệch màu đường kẻ.** Series dạng đường chỉ đặt `lineStyle.color`, ECharts lấy
màu mặc định cho chú giải. `mk()` giờ tự gán `itemStyle` theo `lineStyle` cho mọi chart.

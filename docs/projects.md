# Trạng thái các file

Cập nhật sau khi đọc trực tiếp 6 file (14/08/2026).

## Tổng quan

| File | Loại | Tone | Thư viện | Trạng thái |
|---|---|---|---|---|
| `BaoCao_RRTT_Bond_key_v7` | Report 5 trang, key system | Wine `#7B2D3B` + paper | ECharts inline | Chuẩn tham chiếu |
| `VBMA_Weekly_20260807` | Weekly 6 trang | Wine `#8B2332` | ECharts inline | Chạy tốt |
| `Phan_tich_GD_Bond_20260608_20260810` | Phân tích deal 5 trang | Navy/gold + dark mode | ECharts inline | Chạy tốt |
| `QC.RR.022_lampd3` | Quy chế + đánh giá GAP | Terracotta `#b83a10` | SVG viết tay | Chạy tốt |
| `Peer_Bond_Dashboard_AutoReport_ByGemini_` | Dashboard 5 tab | Terracotta `#b84a32` | **CDN** | **Không dùng được offline** |
| `cfa_l1_console` | Study console (không phải dashboard) | Ink/gold/paper | KaTeX CDN, degrade được | Chạy được offline |

---

## BaoCao_RRTT_Bond_key_v7 — bản tham chiếu

File tốt nhất hiện có. Dùng làm mẫu cho mọi build mới.

- 5 trang: Highlight · Chi tiết danh mục · QTRR theo kịch bản · Thị trường TPCP · FI Bond & CD
- **Key system 2 tầng** `KEY_DEF` + `data-k`/`data-kDef` (dòng 119–184) — xem
  [`key-system.md`](key-system.md). 58 phần tử `data-k` trong DOM.
- **Vòng đời chart tốt nhất**: `REG` + `ResizeObserver` + lazy `tryInit` (dòng 669–687).
  Chart trong tab ẩn không init sớm nên không bị méo.
- Toggle MSB/SBV qua biến `SCOPE`, chỉ hiện ở sub-tab Indicator (`syncScopeVisible`, dòng 1309).
- API ngoài: `window.RRTT = { mount, resizeAll, goPage, setScope, rerender, keys }`.
- Dùng `localStorage` key `rrtt.nav.v2` — **chỉ nhớ tab đang mở**, bọc try/catch (dòng 1294–1296).
  Không lưu dữ liệu báo cáo. Nằm trong phạm vi cho phép.
- Font qua var `--f-head` / `--f-body` / `--f-num`, đều là `'DM Sans','Segoe UI',sans-serif`.

## VBMA_Weekly_20260807

- 6 trang: Tổng quan · Thị trường tiền tệ · Ngoại hối · TPCP sơ cấp · TPCP thứ cấp · TPDN
- Chart registry phẳng: `const charts=[]` + `mk()` (dòng 503–504). Đơn giản và đúng — mọi chart
  đi qua một cửa.
- Resize 3 nhịp sau khi chuyển trang: double `requestAnimationFrame` + `setTimeout(350)`.
- Token màu `--terra-*` nhưng accent là wine `#8B2332`, không phải terracotta cam.
- Tab Data Validation đã xóa — discrepancy chỉ hiện qua footnote asterisk inline.

## Phan_tich_GD_Bond_20260608_20260810

- 5 trang: Tổng quan · Nhóm A Đi vay · Nhóm B Cho vay · Bất thường · Chi tiết cặp deal
- **Dark mode đầy đủ**: `toggleTheme()` + `data-theme="dark"` trên `<html>`, toàn bộ màu qua
  CSS var kể cả `--chart-text` / `--chart-label`. Đây là file duy nhất làm được việc này.
- Resize không cần registry: quét `.ch` rồi `echarts.getInstanceByDom` (dòng 456–457).
- Tone navy/gold `#0A2463` / `#D4A843` — khác hẳn 2 tone còn lại.
- Data nằm trong một object `D` (dòng 244).

## QC.RR.022_lampd3

Tài liệu văn bản, không phải dashboard. Không dùng ECharts.

- Nội dung là dữ liệu: `DEF` (31 khoản Điều 4, nguyên văn) + `ARTS` (các điều), render bằng
  `khHTML` / `artPage` / `defPage`.
- **Chart là SVG viết tay** nhúng thẳng chuỗi — nhẹ, in được, không cần thư viện.
- Tìm kiếm bỏ dấu: `stripD()` dùng `normalize('NFD')` + xử lý riêng `đ/Đ`.
- **Bôi vàng / tẩy** trên `Range` thật (`wrapRange`, `highlightSelection`, `unwrapMark`).
- **File duy nhất có `@media print`** (dòng 114–115).
- Khớp đúng tone terracotta trong docs: `#b83a10` / `#fdf3f0` / `#fff7f4` / `#f5c9b8`.

## Peer_Bond_Dashboard_AutoReport_ByGemini_ — CẦN SỬA

Bản Gemini sinh, khác `Peer_Bond_Dashboard_v5.1.html`. **Chưa dùng được ở bank.**

Hai lỗi phải sửa trước khi giao:

1. **Load qua CDN** (dòng 4–5): `cdn.jsdelivr.net/npm/echarts@5.4.3` và `unpkg.com/xlsx`.
   Bank chặn outbound → mở ra trắng trang. Phải inline cả ECharts lẫn SheetJS.
2. **`resizeAllCharts()` sót chart** (dòng 537–538): hardcode 11 id, nhưng `renderGtgd()`
   dòng 831 tạo `let chart = echarts.init(...)` local không đăng ký → chart đó không bao giờ
   resize. Sửa bằng cách chuyển hết sang một registry (xem
   [`design-system.md`](design-system.md)).

Phần làm đúng, giữ lại: 27 bank trong `BANK_ORDER`, pills multi-select All/Clear/Top10,
export `outerHTML` (dòng 593) — cơ chế save duy nhất đang thực sự có trong 6 file.

## cfa_l1_console — sản phẩm khác loại

Không phải dashboard tài chính. Xem [`content-console.md`](content-console.md).

## Chưa build

- "Đặt cược dự báo" — logged, chưa build
- "Phe Bull vs Phe Bear" — logged, chưa build
- Block "Bình luận / Nhận định" (`mkCmt*`) — spec có, chưa file nào triển khai

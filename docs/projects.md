# Trạng thái các file

Cập nhật sau khi đọc trực tiếp 6 file (14/08/2026).

## Tổng quan

| File | Loại | Tone | Thư viện | Trạng thái |
|---|---|---|---|---|
| `BaoCao_RRTT_Bond_key_v7` | Report 5 trang, key system | Wine `#7B2D3B` + paper | ECharts inline | Chuẩn tham chiếu |
| `VBMA_Weekly_20260807` | Weekly 6 trang | Wine `#8B2332` | ECharts inline | Chạy tốt |
| `Phan_tich_GD_Bond_20260608_20260810` | Phân tích deal 5 trang | Navy/gold + dark mode | ECharts inline | Bản cũ, thay bằng bản 20/08 |
| `Phan_tich_GD_Bond_20260608_20260820` | Phân tích deal 6 trang, 2 nhóm A/B | Navy/gold + dark mode | ECharts inline | **Đã xoá** (MSB_RP_DM), thay bằng bản 04/09 |
| `Phan_tich_GD_Bond_20260608_20260904` | Phân tích deal 6 trang, 4 nhóm phân loại | Navy/gold + dark mode | ECharts inline | Bản hiện hành (RPBOD_B002) |
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

## Phan_tich_GD_Bond_20260608_20260820 — bản mở rộng

Dựng lại từ `bond/source/Dealps_0806_2008.xlsx` (346 chân deal, CaptureDate 08/06–20/08/2026),
thay bản `..._20260810`. Toàn bộ 136 cặp của bản cũ tái lập khớp từng đồng; thêm 37 cặp từ dữ
liệu mới.

**Sinh lại số:** `python3 bond/tools/ghep_cap_deal_bond.py bond/source/Dealps_0806_2008.xlsx
--html bond/Phan_tich_GD_Bond_20260608_20260820.html` — thay khối `const D` tại chỗ, không đụng
phần trình bày. Phương pháp ghép cặp và các bẫy dữ liệu: [`bond-deal-pairing.md`](bond-deal-pairing.md).

- 6 trang: Tổng quan · Nhóm A Đi vay · Nhóm B Cho vay · **Đối tác** · Bất thường · Chi tiết cặp deal
- **Ghép cặp 2 pass**: pass 1 khớp `(CaptureDate, Quantity)` rồi ưu tiên deal-id gần nhau — hai
  chân của một repo được book cùng lúc nên tiêu chí này thắng thứ tự thanh toán thuần túy khi
  nhiều chân trùng ngày; pass 2 FIFO phần dư, tách chân lớn. Kết quả 0 chân lẻ.
- **Ghép chéo đối tác** cho cặp KBNN mua ↔ PGBV-HO bán (deal 50120 ↔ 50127+50128).
- Chart registry `REG` + thunk `BUILDERS`: option chart bake màu theme nên phải giữ sau hàm để
  dựng lại được khi đổi dark mode. `chart(id, fn)` thay cho `mk(id, opt)` trực tiếp.
- **Bẫy CSS đã sửa**: `.pos`/`.neg` global đè lên `.kpi.pos` (dùng `.pos` làm tone nền) → chữ
  trắng thành xanh lá trên nền xanh đậm. Đã scope về `td.pos, span.pos`.

## Phan_tich_GD_Bond_20260608_20260904 — 4 nhóm phân loại, nguồn RPBOD_B002

Dựng lại từ `bond/source/RPBOD_B002_2026.08.27.xlsx` (601 dòng, lọc còn 374 sau khi bỏ
folder AFS-GOV/AFS-ALM/rail nội bộ), thay hoàn toàn bản `..._20260820` (MSB_RP_DM).
Phương pháp, bẫy dữ liệu, bug đã sửa: [`bond-deal-pairing.md`](bond-deal-pairing.md)
mục "Phụ lục — GovBond RPBOD_B002".

**Sinh lại số:** `cd bond/tools && python3 render_govbond_report.py` — chạy
`ghep_cap_deal_govbond.py` → `build_D_govbond.py` → ghép template + ECharts blob, ghi
đè `bond/Phan_tich_GD_Bond_20260608_20260904.html` (KHÔNG sửa in-place như bản cũ; file
này dựng lại toàn bộ từ template mỗi lần).

- 6 trang: Tổng quan · **Buy trước** (Cho vay tiền | Vay bond, 2 cột con) · **Sell trước**
  (Vay tiền | Cho vay bond, 2 cột con) · Đối tác (4 nhóm) · Tra cứu · Chi tiết cặp deal.
  Thay hẳn cấu trúc nhị phân Nhóm A/Nhóm B của bản cũ.
- `CaptureDate` chỉ để **ghép cặp**; `SettlementDate` để đọc chiều trước/sau và kỳ hạn.
  Bản đầu dùng CaptureDate cho cả hai nên hỏng — xem `bond-deal-pairing.md`.
- 198 cặp: 102 Vay tiền, 59 Cho vay bond, 37 Cho vay tiền, 0 Vay bond.
- Nhóm rỗng hiện một dòng "Không có cặp nào trong kỳ" thay vì chart/bảng trống.
- **Bẫy đơn vị đã sửa** (dễ tái phạm): `cash` là **tỷ**, `pnl` là **triệu** — trừ thẳng
  hai đại lượng này cho ra sai 1000 lần (đã xảy ra ở bảng deal đặc biệt). Chart nào đã
  đổi data sang nghìn tỷ thì `axisLabel.formatter` **không** được chia 1000 lần nữa.
- **Bẫy dữ liệu**: cột `GrossAmount` có 16/346 dòng là **string** dạng `'  1012036120.0000000K'`
  (đơn vị nghìn), không phải số. Không parse là lệch 1000 lần.
- Ngưỡng ngoại lệ **theo kỳ hạn**: yield lệch ≥ 10bp chỉ là định giá lại khi kỳ hạn ≤ 30 ngày;
  deal 351 ngày lệch 31bp là bình thường. Deal 1–2 ngày có %/năm phóng đại do làm tròn giá —
  tách thành nhóm `noise` riêng, không đếm là lỗi.
- Tách **lãi đã đáo hạn** (−4,4 tỷ) khỏi **lãi theo hợp đồng** (271,8 tỷ, phần lớn đáo hạn sau
  20/08). Trộn hai con số là đọc sai P&L trong kỳ.

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

## Sắp build — Internal Rating Report (HTML)

Jak sẽ start dự án chuyển bảng internal rating của FI Bond thành report HTML.

**Nguồn dữ liệu:** `bond/source/01.RptTool_FIBond_2026.08.19.xlsm` (đã lưu trong repo).

Ba thứ cần lấy từ file đó:

| Vùng | Nội dung |
|---|---|
| `Rating!A1:K22` | Bảng `Portfolio_rating` — Issuer · Total Amount · Rating · Review Date · Fitch · Moody · S&P · Tỷ lệ · Tỷ lệ lũy kế · Đối ứng · Tổng đã điều chỉnh |
| `Rating!M1:P8` | Tổng hợp theo bậc rating — Rating · Amount · Tỷ lệ · Tỷ lệ lũy kế |
| `Rating!R1:V22` | **Bảng quy chuẩn rating** — Internal rating ↔ Fitch ↔ S&P ↔ Moody's, kèm nhãn Investment grade / High yield / Default |

Thang nội bộ: `AAA · AA · A · BBB · BB · B · N/A`. Thứ tự xếp hạng đúng nằm ở query
`Portfolio rating` sau khi sửa (xem [`file01-issues.md`](file01-issues.md) #18) — dùng
`List.PositionOf` trên thang tường minh, không phải chuỗi `if/else` cũ.

Lưu ý khi build:
- Rating là **thang thứ bậc**, không phải nhãn rời. Mọi sắp xếp và tô màu phải theo thứ tự
  `AAA → D`, mã chưa có rating xuống cuối.
- Có mã ngoại tệ (MFKR, ccy KGS) đã quy đổi VND qua bảng `Tygia` — số trên bảng là tỷ VND.
- Tone phù hợp: wine (bond risk) theo [`design-system.md`](design-system.md).

## Chưa build

- "Đặt cược dự báo" — logged, chưa build
- "Phe Bull vs Phe Bear" — logged, chưa build
- Block "Bình luận / Nhận định" (`mkCmt*`) — spec có, chưa file nào triển khai

# Design System

Tài liệu này mô tả **thực tế đang chạy** trong các file production, không phải mẫu lý tưởng.
Mọi khẳng định đều trỏ được về file + dòng cụ thể.

## Nền tảng kỹ thuật

- **ECharts 5.4.3** nhúng inline trong `<script>` — không CDN (IT bank block outbound).
  Blob chiếm ~936KB, nên file dashboard thật thường ~1.1MB.
- Output luôn là **một file HTML standalone**.
- Ngoại lệ đang tồn tại: `Peer_Bond_Dashboard_AutoReport_ByGemini_` load ECharts + SheetJS
  qua CDN → mở ở bank là trắng trang. Xem `projects.md`.

## Font

Không phải "Segoe UI" đơn lẻ mà là **stack có chủ đích**:

```css
--f-head: 'DM Sans','Segoe UI',sans-serif;
--f-body: 'DM Sans','Segoe UI',sans-serif;
--f-num:  'DM Sans','Segoe UI',sans-serif;   /* + font-variant-numeric:tabular-nums */
```

- `<link>` Google Fonts bị bank chặn → **tự động degrade về Segoe UI**. Đây là thiết kế,
  không phải lỗi; giữ nguyên `<link>` để máy ngoài bank vẫn đẹp.
- Heading serif (`Crimson Pro`) chỉ dùng ở tài liệu dạng văn bản như `QC.RR.022`.
- Số liệu luôn `font-variant-numeric: tabular-nums` để cột số thẳng hàng.

## Palette — cấu trúc token, 3 tone

Đừng nhớ hex rời. Nhớ **vai trò token**, rồi tra giá trị theo tone của dự án.

| Token | Vai trò |
|---|---|
| `bg` | Nền trang |
| `surf` / `card` | Nền panel, card |
| `border` | Viền, đường kẻ bảng |
| `ink` | Chữ chính |
| `muted` / `faint` | Chữ phụ, label trục |
| `accent` | Màu thương hiệu của tone — nav active, header bảng, đường chart chính |

### Tone A — Terracotta (tài liệu, Peer Bond)

| Token | Hex |
|---|---|
| accent | `#b83a10` → `#e8551a` (gradient) |
| bg | `#fdf3f0` |
| surf | `#fff7f4` |
| border | `#f5c9b8` |
| muted | `#9a4020` |

### Tone B — Wine (VBMA, Bond Risk Report)

VBMA (`--terra-*`):

| Token | Hex |
|---|---|
| accent | `#8B2332` (dark `#6E1A26`) |
| bg | `#f4f1ef` |
| surf | `#faf8f6` |
| border | `#e5dcd8` |
| muted | `#6B5A5C` |

RRTT v7 — nền giấy, accent kép wine + gold:

| Token | Hex |
|---|---|
| accent | `#7B2D3B` (dark `#5E1F2B`), gold `#b68235` |
| bg | `#fdfcfa` / `#f3f2f2` |
| ink | `#201f1d` |
| muted | `#8d877c`, `#6b665e` |

### Tone C — Navy/Gold (Phân tích GD Bond) — có dark mode

Khai qua CSS var, đổi bằng `toggleTheme()` + `data-theme="dark"` trên `<html>`:

| Token | Light | Dark |
|---|---|---|
| accent | `#0A2463` navy | `#F5E6B8` gold-light |
| gold | `#D4A843` | `#D4A843` |
| bg | `#F0F2F5` | `#070F1F` |
| surface | `#fff` | `#0C1A30` |
| card | `#fff` | `#0F2040` |
| chart-text | `#1A1A2E` | `#E8EDF4` |
| chart-label | `#555` | `#7A8FA8` |

### Semantic (chung mọi tone)

| Ý nghĩa | Hex |
|---|---|
| Chữ chính / chart label sáng | `#1a202c` |
| Xanh — đạt, tích cực | `#276749` |
| Đỏ — vi phạm, tiêu cực | `#9B2C2C` / `#B3261E` |
| Hổ phách — cảnh báo | `#6B5A5C` (VBMA), `#b68235` (RRTT) |

Đèn ngưỡng RRTT: **XANH < 80% · HỔ PHÁCH 80–100% · ĐỎ ≥ 100%** (`thrWarn:0.8`, `thrBreach:1`).

Chart label mặc định đen `#1a202c`. File có dark mode thì label đi theo token
`--chart-text` / `--chart-label`, không hardcode.

## Layout & UX

- **Sidebar dọc** (không phải tab ngang) là bố cục thực tế của cả 4 dashboard: `.ni` /
  `.sb-item` gọi `goPage('pN', this)`.
- **KPI cards** hàng đầu mỗi trang.
- **Sub-tab** khi cần cấp hai (RRTT v7: `goSub(pageId, subId)`).
- **Filter pills** (`.fbtn` / `.pill`) cho multi-select — Peer Bond có All / Clear / Top10.

## Vòng đời chart — chuẩn là pattern `REG` của RRTT v7

Mọi chart phải đi qua **một cửa duy nhất**. Không bao giờ `echarts.init()` rải rác rồi gán
biến local — chart đó sẽ không bao giờ resize.

Pattern tốt nhất (RRTT v7, dòng 669–687): registry + `ResizeObserver` + lazy init.

```js
const REG = {};
function mk(id, opt) {
  const el = document.getElementById(id); if (!el) return;
  const e = REG[id] || (REG[id] = {});
  e.el = el; e.opt = opt;
  if (!e.ro && window.ResizeObserver) { e.ro = new ResizeObserver(() => tryInit(id)); e.ro.observe(el); }
  tryInit(id);
}
function tryInit(id) {
  const e = REG[id]; if (!e || !e.opt || !window.echarts) return;
  const el = e.el;
  if (!el.isConnected || !el.clientWidth || !el.clientHeight) return;   // tab ẩn → hoãn
  const sig = el.clientWidth + 'x' + el.clientHeight;
  if (!e.chart || e.chart.isDisposed()) {
    e.chart = echarts.init(el, null, { renderer: 'canvas' });
    e.chart.setOption(e.opt, true); e.sig = sig; e.applied = e.opt; return;
  }
  if (e.applied !== e.opt) { e.chart.setOption(e.opt, true); e.applied = e.opt; e.sig = null; }
  if (e.sig !== sig) { e.chart.resize(); e.sig = sig; }
}
function resizeAll()  { Object.keys(REG).forEach(tryInit); }
function resizeSoon() { requestAnimationFrame(resizeAll); setTimeout(resizeAll, 60); setTimeout(resizeAll, 320); }
window.addEventListener('resize', resizeAll);
```

Vì sao pattern này thắng:

- Chart trong tab ẩn có `clientWidth = 0` → init lúc đó sẽ ra chart méo. `tryInit` **hoãn**
  đến khi tab mở.
- `ResizeObserver` bắt cả trường hợp container đổi kích thước mà `window` không resize
  (mở/đóng sidebar, bảng giãn).
- `resizeSoon()` gọi 3 nhịp (rAF, 60ms, 320ms) để phủ animation chuyển trang.

### Hai biến thể chấp nhận được

```js
// VBMA — registry phẳng, mọi chart qua mk()
const charts = [];
function mk(id, opt) { const el = document.getElementById(id); if (!el) return;
  const c = echarts.init(el); c.setOption(opt); charts.push(c); }
window.addEventListener('resize', () => charts.forEach(c => c.resize()));

// Phân tích GD Bond — không cần registry, hỏi thẳng DOM
window.addEventListener('resize', () =>
  document.querySelectorAll('.ch').forEach(e => {
    const c = echarts.getInstanceByDom(e); if (c) c.resize();
  }));
```

### Anti-pattern (đang có thật trong Peer Bond bản Gemini)

```js
// SAI — danh sách id hardcode
function resizeAllCharts() {
  ['ovEChart1','ovEChart2',/* …11 id… */].forEach(c => { if (window[c]) window[c].resize(); });
}
// …nhưng chỗ khác lại:
let chart = echarts.init(document.getElementById(chartId));  // local, không đăng ký → chết
```

## Save mechanism

Không lưu **dữ liệu báo cáo** vào `localStorage`. Quy trình export:

1. Serialize state/data hiện tại vào các biến `RAW_*` trong DOM.
2. `"<!DOCTYPE html>\n" + document.documentElement.outerHTML`
3. `Blob` → `<a download>` → file HTML standalone mở lại là đủ dữ liệu.

**Phạm vi rule localStorage**: cấm lưu dữ liệu/nội dung báo cáo. Nhớ vị trí UI (tab đang mở)
thì chấp nhận được nếu bọc `try/catch` và hỏng vẫn chạy — RRTT v7 làm vậy với key
`rrtt.nav.v2` (dòng 1294–1296).

## In ấn

Mới `QC.RR.022` có `@media print` (dòng 114–115). 4 file còn lại chưa có — là gap chung nếu
report cần in:

```css
@media print{
  .sidebar,.navfoot,.searchwrap,.hltb{display:none}
  .page{display:block!important;page-break-after:always}
  .kh{break-inside:avoid}
  .hl-mark{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
```

# Design System

## Nền tảng kỹ thuật

- **ECharts 5.4.3** nhúng inline trong `<script>` — không CDN (IT bank block outbound).
- **Font**: Segoe UI.
- Output luôn là **một file HTML standalone**.

## Palette

### Mặc định — terracotta / cam gạch

| Vai trò | Mã |
|---|---|
| Primary gradient | `#b83a10` → `#e8551a` |
| Background | `#fdf3f0` |
| Card | `#fff7f4` |
| Border | `#f5c9b8` |
| Chart label | `#1a202c` (đen, trên **tất cả** chart) |

### Wine-red — chỉ dùng cho bond risk report

| Vai trò | Mã |
|---|---|
| VBMA Dashboard | `#8B2332` |
| Bond Portfolio Risk Report (nav active, sub-tab, table header) | `#7B2D3B` |

Không dùng wine-red cho dashboard khác; không dùng terracotta cho bond risk report.

## Layout & UX

- **Tab navigation** ở đầu trang; sub-tab khi cần cấp hai.
- **KPI cards** hàng đầu mỗi tab.
- **Lazy chart init theo tab**: chart chỉ khởi tạo khi tab của nó được mở lần đầu.
- **Resize-responsive**: bind `window.resize` → `resizeAllCharts()`.

### `resizeAllCharts()`

Phải cover **TẤT CẢ** chart instances, không chỉ các chart trong object `CH` được
lazy-load. Mọi instance tạo ngoài `CH` (chart khởi tạo eager, chart trong modal,
chart trong bảng) đều phải được đăng ký để resize.

```js
var CH = {};            // lazy-loaded, theo tab
var EXTRA_CHARTS = [];  // mọi instance khác — push ngay sau khi init

function resizeAllCharts() {
  Object.keys(CH).forEach(function (k) { CH[k] && CH[k].resize(); });
  EXTRA_CHARTS.forEach(function (c) { c && c.resize(); });
}
window.addEventListener('resize', resizeAllCharts);
```

## Save mechanism

Không dùng `localStorage`. Quy trình:

1. Serialize toàn bộ state/data hiện tại vào các biến `RAW_*` trong DOM
   (ghi đè nội dung `<script id="raw-data">`).
2. Lấy `document.documentElement.outerHTML`.
3. Tạo `Blob` → `<a download>` → export ra file HTML standalone mở lại là đủ dữ liệu.

Hệ quả: file export ra phải mở được offline, không phụ thuộc session cũ.

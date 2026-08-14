# CLAUDE.md

Repo này chứa **design system chuẩn** cho mọi dashboard/report HTML. Đọc `docs/` trước khi
build hoặc sửa bất kỳ file HTML nào.

- [`docs/design-system.md`](docs/design-system.md) — palette, font, layout, chart, save mechanism
- [`docs/key-system.md`](docs/key-system.md) — key-ification 3 tier, Excel `Key_Config` / `Key_Map`
- [`docs/comment-block.md`](docs/comment-block.md) — block "Bình luận / Nhận định"
- [`docs/projects.md`](docs/projects.md) — trạng thái từng dashboard

## Ràng buộc không thương lượng

1. **Offline single-file.** ECharts 5.4.3 nhúng inline. Không CDN, không external asset —
   IT bank block outbound. Không đề xuất `<script src="https://...">`.
2. **Không `localStorage`.** State lưu trong RAW variables, export qua
   `document.documentElement.outerHTML`.
3. **Thiếu key → fallback text hardcode.** Không bao giờ render `undefined`. Xóa sạch
   `Key_Config` thì file vẫn phải chạy.
4. **`resizeAllCharts()` cover TẤT CẢ chart instances**, không chỉ object `CH` lazy-loaded.
5. **Palette theo đúng dự án.** Terracotta `#b83a10 → #e8551a` là mặc định; wine-red
   `#8B2332` / `#7B2D3B` chỉ dùng cho bond risk report và VBMA.
6. **Chart label luôn đen `#1a202c`** trên mọi chart.
7. **Font Segoe UI.**

## Lưu ý

Full pipeline automation không khả thi do IT constraints — mọi deliverable phải là một file
HTML đứng độc lập, mở được offline.

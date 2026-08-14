# HTML Dashboard Design System

Chuẩn chung cho tất cả dashboard/report HTML build offline (single-file, embed toàn bộ,
không CDN do IT bank block outbound).

## Tài liệu

| File | Nội dung |
|---|---|
| [`docs/design-system.md`](docs/design-system.md) | Palette, font, layout, chart, save mechanism |
| [`docs/key-system.md`](docs/key-system.md) | Chuẩn key-ification 3 tier + Excel `Key_Config` / `Key_Map` |
| [`docs/comment-block.md`](docs/comment-block.md) | Block "Bình luận / Nhận định" tái sử dụng |
| [`docs/projects.md`](docs/projects.md) | Trạng thái từng dashboard (Peer Bond, VBMA, Bond Risk Report, ...) |

## Ràng buộc nền tảng (không thương lượng)

- **Offline single-file**: ECharts 5.4.3 nhúng inline, không CDN, không external asset.
- **Không localStorage**: state lưu trong RAW variables, export qua
  `document.documentElement.outerHTML`.
- **Full pipeline automation không khả thi** do IT constraints — mọi build phải ra được
  một file HTML đứng độc lập.

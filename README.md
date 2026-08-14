# HTML Design System

Chuẩn chung cho các file HTML build offline (single-file, embed toàn bộ, không phụ thuộc mạng
do IT bank block outbound).

Docs viết từ việc đọc trực tiếp 6 file production, mỗi khẳng định trỏ được về file + số dòng.

## Tài liệu

| File | Nội dung |
|---|---|
| [`docs/design-system.md`](docs/design-system.md) | Token màu 3 tone dashboard, font stack, layout, vòng đời chart, save, in ấn |
| [`docs/key-system.md`](docs/key-system.md) | Key-ification 3 tier, cơ chế fallback 2 tầng, Excel `Key_Config` / `Key_Map` |
| [`docs/content-console.md`](docs/content-console.md) | Pattern công cụ đọc/học: DSL nội dung, KaTeX degrade, screen system, tone ink/gold/paper |
| [`docs/comment-block.md`](docs/comment-block.md) | Block "Bình luận / Nhận định" — **spec, chưa triển khai** |
| [`docs/projects.md`](docs/projects.md) | Trạng thái từng file, kèm việc cần sửa |

## Hai loại sản phẩm

- **Dashboard tài chính** — dữ liệu từ Excel, ECharts inline, sidebar nhiều trang.
  Mẫu tốt nhất: `BaoCao_RRTT_Bond_key_v7.html`.
- **Content console** — nội dung viết tay có cấu trúc, không ECharts, có state cá nhân.
  Mẫu: `cfa_l1_console.html`, `QC.RR.022_lampd3.html`.

## Ràng buộc nền tảng

- **Offline single-file**: thứ gì thiếu nó thì trang chết phải nhúng inline. Asset chỉ để
  nâng cấp (font, KaTeX) được phép qua CDN nếu đã có fallback đọc được.
- **Không lưu dữ liệu báo cáo vào localStorage**: state đi trong RAW variables, export qua
  `document.documentElement.outerHTML`.
- **Thiếu key → về chữ gốc**, không bao giờ `undefined`.
- **Mọi chart qua một registry duy nhất**, nếu không sẽ có chart không bao giờ resize.

Chi tiết và lý do trong [`CLAUDE.md`](CLAUDE.md).

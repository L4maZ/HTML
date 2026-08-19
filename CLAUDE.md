# CLAUDE.md

Repo này chứa **design system** cho các file HTML của Jak. Docs được viết từ việc đọc trực
tiếp file production, không phải từ mẫu lý tưởng. Đọc `docs/` trước khi build hoặc sửa bất kỳ
file HTML nào.

- [`docs/design-system.md`](docs/design-system.md) — token màu 3 tone, font stack, vòng đời chart, save
- [`docs/key-system.md`](docs/key-system.md) — key-ification 3 tier, cơ chế fallback 2 tầng, Excel
- [`docs/content-console.md`](docs/content-console.md) — pattern cho công cụ đọc/học (không phải dashboard)
- [`docs/comment-block.md`](docs/comment-block.md) — block "Bình luận / Nhận định" (**chưa triển khai**)
- [`docs/file01-rpttool.md`](docs/file01-rpttool.md) — File 01 `RptTool_Bond`: cấu trúc thật, đối chiếu memory
- [`docs/projects.md`](docs/projects.md) — trạng thái từng file

## Ràng buộc không thương lượng

1. **Offline single-file.** ECharts 5.4.3 nhúng inline. Không CDN, không external asset cho
   thứ mà thiếu nó thì trang chết — IT bank block outbound. Không đề xuất
   `<script src="https://...">`.
   *Ngoại lệ hợp lệ*: asset chỉ để **nâng cấp** trải nghiệm và đã có fallback đọc được khi
   chặn — Google Fonts (degrade về Segoe UI), KaTeX (degrade về text thô `.fraw`).
2. **Không lưu dữ liệu báo cáo vào `localStorage`.** State đi trong RAW variables, export qua
   `document.documentElement.outerHTML`. Nhớ vị trí UI (tab đang mở) thì được, miễn bọc
   `try/catch` và hỏng vẫn chạy. Công cụ cá nhân (study console) không thuộc phạm vi rule này.
3. **Thiếu key → fallback text hardcode.** Không bao giờ render `undefined`. Xóa sạch
   `Key_Config` thì file vẫn phải chạy. Kiểm bằng `RRTT_setKeys({})`.
4. **Mọi chart phải đi qua một cửa duy nhất.** Không `echarts.init()` rải rác gán biến local —
   chart đó sẽ không bao giờ resize. Dùng registry `REG` + `ResizeObserver` của RRTT v7.
5. **Palette theo đúng tone của dự án** — có 4 tone, không phải một:
   terracotta (tài liệu, Peer Bond) · wine (VBMA, bond risk) · navy-gold (phân tích deal,
   có dark mode) · ink-gold-paper (content console). Nhớ vai trò token, tra hex trong docs.
6. **Chart label mặc định đen `#1a202c`.** File có dark mode thì label theo token
   `--chart-text` / `--chart-label`, không hardcode.
7. **Font là stack, không phải một font.** Dashboard:
   `'DM Sans','Segoe UI',sans-serif`. Content console dùng font hệ thống. Số liệu luôn
   `font-variant-numeric: tabular-nums`.

## Lưu ý

Full pipeline automation không khả thi do IT constraints — mọi deliverable phải là một file
HTML đứng độc lập, mở được offline.

File tham chiếu tốt nhất khi cần mẫu: `BaoCao_RRTT_Bond_key_v7.html` (key system + vòng đời
chart). Đừng lấy mẫu từ bản Peer Bond do Gemini sinh — nó dùng CDN và sót chart khi resize.

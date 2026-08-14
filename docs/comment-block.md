# Block "Bình luận / Nhận định"

Pattern tái sử dụng across mọi file HTML. Cho phép analyst thêm nhận định vào dashboard
mà không đụng code, và ẩn sạch block khi không có gì để nói.

## Keys

| Key | Ý nghĩa |
|---|---|
| `mkCmtShow` | `TRUE` / `FALSE`. `FALSE` → **ẩn cả block, không chừa khoảng trắng** (`display:none`, không phải `visibility`) |
| `mkCmtTitle` | Tiêu đề block |
| `mkCmtDate` | Ngày. `AUTO_UNLESS_OVERRIDE` → tự lấy ngày build trừ khi có giá trị override |
| `mkCmtAuthor` | Người viết |
| `mkCmtTone` | `bull` / `bear` / `neutral` → **chỉ đổi màu viền trái**: xanh / đỏ / terracotta. Không đổi nền, không đổi chữ |
| content lines | Theo từng tab, mỗi tab một bộ key riêng |
| `mkCmtWatch` | Nền tối hơn, **hiện trên mọi tab** |
| `mkCmtEmpty` | Text thay thế khi tất cả content line đều trống |

## Rule

- **Dòng trống = ẩn hoàn toàn.** Không bao giờ render bullet rỗng.
- Tất cả line trống → hiện `mkCmtEmpty` thay vì block rỗng.
- `mkCmtShow = FALSE` → block biến mất, layout phía dưới dồn lên, không để lại gap.

## Tone → viền trái

```
bull    → xanh
bear    → đỏ
neutral → terracotta (mặc định palette)
```

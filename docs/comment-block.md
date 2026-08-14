# Block "Bình luận / Nhận định"

> **Trạng thái: CHƯA TRIỂN KHAI.** Không file HTML nào hiện có key `mkCmt*`. Đây là spec
> mục tiêu, không phải mô tả code đang chạy.

## Thứ gần nhất đang chạy: các key `hi*` của RRTT v7

`BaoCao_RRTT_Bond_key_v7.html` dòng 139 khai 6 key nhận định, mặc định **rỗng**, Jak nhập tay
hằng ngày qua Excel:

```js
hiCompliance: '', hiTb: '', hiBb: '', hiFi: '', hiRiskPrice: '', hiRiskNim: ''
```

Kèm 6 key tiêu đề khối tương ứng (`hiComplianceTitle`, `hiTbTitle`, `hiBbTitle`, `hiFiTitle`,
`hiRiskPriceTitle`, `hiRiskNimTitle`).

Đây đã là một nửa của pattern: **nội dung tách khỏi code, rỗng thì không hiện**. Phần còn
thiếu là metadata (ngày / tác giả / tone) và cơ chế ẩn cả block.

## Spec đầy đủ

| Key | Ý nghĩa |
|---|---|
| `mkCmtShow` | `TRUE` / `FALSE`. `FALSE` → **ẩn cả block, không chừa khoảng trắng** (`display:none`, không phải `visibility`) |
| `mkCmtTitle` | Tiêu đề block |
| `mkCmtDate` | Ngày. `AUTO_UNLESS_OVERRIDE` → tự lấy ngày build trừ khi có giá trị override |
| `mkCmtAuthor` | Người viết |
| `mkCmtTone` | `bull` / `bear` / `neutral` → **chỉ đổi màu viền trái**: xanh `#276749` / đỏ `#9B2C2C` / accent của tone. Không đổi nền, không đổi chữ |
| content lines | Theo từng tab, mỗi tab một bộ key riêng |
| `mkCmtWatch` | Nền tối hơn, **hiện trên mọi tab** |
| `mkCmtEmpty` | Text thay thế khi tất cả content line đều trống |

## Rule

- **Dòng trống = ẩn hoàn toàn.** Không bao giờ render bullet rỗng.
- Tất cả line trống → hiện `mkCmtEmpty` thay vì block rỗng.
- `mkCmtShow = FALSE` → block biến mất, layout phía dưới dồn lên, không để lại gap.

## Khi triển khai

Đi qua đúng cơ chế key ở [`key-system.md`](key-system.md): khai trong `KEY_DEF`, đọc bằng
`kv()`, để `applyDataKeys()` lo phần fallback. Không tự viết nhánh đọc key riêng.

Đường nâng cấp gọn nhất: đổi tên 6 key `hi*` của RRTT v7 thành bộ `mkCmt*` chuẩn, rồi bọc
chúng trong một component dùng lại được cho VBMA và Peer Bond.

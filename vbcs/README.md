# vbcs/ — Văn bản chính sách (thẩm định rủi ro)

Nhánh sản phẩm mới, tách khỏi `bond/`: các file HTML dạng **content console** cho việc đọc và
thẩm định văn bản chính sách quản lý rủi ro (quy chế, quy định nội bộ), không phải dashboard
số liệu. Cùng loại sản phẩm với `cfa_l1_console.html` — xem
[`../docs/content-console.md`](../docs/content-console.md).

## QC.RR.022

`QC.RR.022/QC.RR.022_lampd3.html` — Quy chế Quản lý Rủi ro thị trường, PIC: Lâm (lampd3).
Nội dung quy chế (Điều 4 — 31 khoản định nghĩa, nguyên văn) + đánh giá GAP theo từng khoản
(khối `.gapbox` đánh dấu GAP có/không, tuyến phòng thủ liên quan). Tone terracotta `#b83a10`.
Chi tiết kỹ thuật đã có trong [`../docs/projects.md`](../docs/projects.md#qcrr022_lampd3).

Bản đánh giá GAP gốc: `QC.RR.022/source/Danh_gia_GAP_QC.RR.022.xlsx` (sheet `QC.RR.022_MR` +
`Gap`). Kết quả đối chiếu Excel ↔ HTML: [`QC.RR.022/THAM_DINH_GAP.md`](QC.RR.022/THAM_DINH_GAP.md)
— 6 mục GAP thật (Điều 5, 7) chưa được build vào HTML; ngược lại 2 GAP mới HTML phát hiện ở
Điều 9 chưa được ghi ngược vào Excel. Hai nguồn GAP hiện **chưa đồng bộ**.

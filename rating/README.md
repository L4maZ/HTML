# Rating — Dashboard FI Rating (Key → HTML)

Báo cáo xếp hạng và danh mục FI Bond/CD. Quy trình: `01.FIBond_Rating.xlsb` → macro `XuatRatingKey` → `Rating_Key_YYYYMMDD.xlsx` → nạp vào HTML ở tab Quản trị dữ liệu.

| File | Vai trò |
|---|---|
| `Rating_Report.html` | Dashboard (offline, JSZip nhúng sẵn). Mang sẵn dữ liệu kỳ 02/10/2026 |
| `Rating_Key_20261002.xlsx` | File key mẫu (schema `rating.key.v1`) |

## Thay đổi 05/10/2026
- Tab **Quản trị dữ liệu** khóa mật khẩu (chỉ chặn bấm nhầm, mật khẩu nằm trong mã nguồn); bản lưu/xuất luôn khóa lại.
- Nút **Lưu đè vào file này** (File System Access API; trình duyệt không hỗ trợ thì tải bản mới). Nút "Xuất HTML mới" vẫn giữ.
- Số theo kiểu Anh (`39,228`), tỷ lệ 1 số lẻ.
- Kiểm tra thêm khi nạp key: cơ cấu rating vs cộng bond-level theo từng rating, kỳ hạn vs bond-level, Amount sau đối ứng âm, Amount âm, trùng mã trái phiếu, issuer nhiều rating, lũy kế trước, KPI vs bảng, ngày review sau reporting date. Tổng lệch ≤ 0.5 tỷ bỏ qua (làm tròn KGS).
- Trình bày theo hệ Bond: bảng nền xám ấm, header wine, số căn phải, vạch ngưỡng BBB trên biểu đồ, cột tỷ trọng ở bảng kỳ hạn, ghi chú "Sau đối ứng" ở KPI, bảng chi tiết không còn phải kéo ngang ở 1366px.
- Hiệu ứng: hiện dần khi vào tab lần đầu, chuyển tab mượt (View Transitions), số KPI chạy một lần khi mở; tắt khi in / giảm chuyển động.

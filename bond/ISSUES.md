# Việc còn mở — Bond tool → file key → HTML

Cập nhật 17/08/2026.

## Đang chờ Jak quyết (chặn việc mở rộng)

| # | Nội dung | Chi tiết |
|---|---|---|
| D1 | `Linked (1)!G9` lấy mẫu số nào | `D9`/`F9` chia tổng face, `G9` chia face AFS → cột Last month của Average Price/1bond lệch. `Report!Q10` ghi PnL chỉ tính trên AFS, nên có thể là chủ ý. Cần chốt: sửa `G9` thành `G8/G5`, hay đổi `D9`/`F9` về AFS cho nhất quán |
| D2 | Banking Book nội bộ: tổng hay tổng dòng con | `D32` = 10,858.31 (SUMIFS) nhưng `D33+D34` = 16,959.31, lệch **6,101 tỷ**. ItD tương tự lệch **716 tỷ**. Trading Book thì `D5 = D6+D7`, ngược lại. Hiện HTML báo đỏ mỗi lần nạp |
| D3 | Số `392` hardcode trong `D74`/`K74` | Hạn mức đầu tư ra nước ngoài đang 14.99/15m USD = **99.94%**. v7 (24/07) ghi 11.72m. Cần xác nhận 392 đúng trước khi lên báo cáo |
| D4 | Ánh xạ `highlight.*` | Mình đoán: `Report!H4` → compliance · `B84` → tb · `L84` → bb · `Q66` → fibond · `B166` → riskPrice · `M95` → riskNim. Jak mở trang Highlight xem có đoạn nào sai chỗ |

## Việc phải làm

| # | Nội dung | Ghi chú |
|---|---|---|
| T1 | **Macro VBA sinh file key 11 sheet** | Bản `.bas` hiện tại mới làm khối 2 theo cấu trúc cũ. Đây là mảnh duy nhất chặn Jak chạy độc lập. Viết theo hướng **dò nhãn** thay vì toạ độ cứng, để chèn/xoá dòng trong Excel không làm vỡ. Sheet `TS` ~6,000 dòng phải ghi bằng mảng |
| T2 | `scenBB` — v7 chưa có nhánh này | Sheet `SCEN` đã chở 132 dòng kịch bản Banking nhưng v7 chỉ vẽ Trading. Cần thêm khối hiển thị trong HTML |
| T3 | `vira4` chưa map | Biểu đồ dự báo VIRA 5 tháng, vẫn đứng ở 24/07 |
| T4 | `volMsb` chưa map | Vẫn đứng ở 24/07 |
| T5 | Block "Bình luận / Nhận định" chuẩn `mkCmt*` | Chưa file nào triển khai. Đường ngắn nhất: đổi 6 key `hi*` của v7 sang chuẩn dùng lại được |

## Chờ test ở bank (thứ Hai)

| # | Nội dung |
|---|---|
| B1 | `DecompressionStream` trên Edge ở bank. Xem mục **Môi trường** trong tab Quản trị: "Đọc file .xlsx: Hỗ trợ" là xanh thì chạy được. Đỏ → đổi sang nhúng SheetJS, file tăng ~900KB, vẫn dưới trần 5MB |
| B2 | Kéo thả file trên máy bank. Nếu lại `NotReadableError` thì dùng nút chọn file |

## Đã đóng

| # | Nội dung | Cách xử lý |
|---|---|---|
| C1 | Vùng kéo thả vỡ layout | `.drop` là `<label>` nên mặc định inline — thêm `display:block` |
| C2 | Nút nạp file không mở được hộp chọn trong khung preview | Đổi `<button>` + JS `.click()` sang `<label for>` |
| C3 | `NotReadableError` khi kéo thả file | Thêm fallback `FileReader` khi `arrayBuffer()` hỏng, kèm thông báo nêu đúng nguyên nhân thường gặp |
| C4 | Xoá `input.value` ngay sau khi bắt đầu đọc | Lỗi tiềm ẩn của riêng đường "bấm chọn file": xoá value trong lúc đọc bất đồng bộ có thể làm Chrome huỷ tham chiếu File. Chuyển sang xoá sau khi đọc xong |
| C5 | Bộ đọc `Chart data` ra 0 dòng | Header nằm cả ở hàng 1, không phải hai hàng |
| C6 | Bản xuất bị xoá mất tab Quản trị | Jak chọn giữ lại (có khoá) — bỏ đoạn xoá |

## Ghi chú kỹ thuật

- Mật khẩu `MR@123` nằm trong mã nguồn — **chỉ chặn bấm nhầm, không phải bảo mật**.
- `dataMaxDate` (đối chiếu ngày dữ liệu thật với `asOf`) đã bàn và **quyết định không làm**:
  Jak luôn đổi `Run Tool!B2` ngay khi mở file.
- Ranh giới dễ vỡ là **File 02 → macro**, không phải **file key → HTML**. HTML khớp theo
  tên cột và KeyID nên thêm dòng/cột vào file key là vô hại; macro thì đang đọc theo toạ độ.

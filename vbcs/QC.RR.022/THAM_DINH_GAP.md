# Thẩm định GAP — QC.RR.022

Đối chiếu `source/Danh_gia_GAP_QC.RR.022.xlsx` (sheet `QC.RR.022_MR` + `Gap`, theo yêu cầu chỉ
dùng 2 sheet này) với công cụ đọc `QC.RR.022_lampd3.html`. Không đọc `Ref`, `Tương quan RR`,
`Sheet3`.

## 1. Đối chiếu nội bộ 2 sheet nguồn — khớp

Sheet `QC.RR.022_MR` (698 dòng, cột G = cờ GAP theo từng khoản) chỉ có **7 khoản GAP=Y** thực
chất (2 dòng Y khác là dòng tiêu đề trùng lặp cờ với dòng nội dung con ngay dưới — Điều 5
khoản 8 và khoản 9, khoản 11). Cả 7 khoản này khớp 1-1 với 7 dòng trong sheet `Gap` (bảng tổng
hợp action). Không có sai lệch giữa 2 sheet nguồn.

| # | Điều — Khoản | Nội dung thiếu (GAP) | Đơn vị phối hợp |
|---|---|---|---|
| 1 | Đ5 K6 — Hạn mức, cảnh báo sớm, kiểm soát tổn thất | Chưa có văn bản liệt kê đầy đủ: chỉ báo rủi ro, cơ chế stop-loss, kiểm soát tổn thất | — |
| 2 | Đ5 K7 — Nguồn giá, dữ liệu, mô hình định giá, IPV | Option: thiếu nguồn giá đáng tin cậy (vola curve). FIBond: thiếu dữ liệu thị trường (curve), mô hình định giá cho SP có embedded. IRSCCS: thiếu nguồn giá kỳ dài | MHCC |
| 3 | Đ5 K8 — Hệ thống báo cáo kịp thời, cảnh báo | Thiếu tính sẵn có của báo cáo đánh giá realtime (Kondor, intraday) | FI-Product |
| 4 | Đ5 K9 — Stress testing, kịch bản, ICAAP | Thiếu kịch bản định tính (event-risk): sự kiện → tác động risk factor → impact danh mục; tần suất quý/adhoc | — |
| 5 | Đ5 K11 — Kiểm soát sản phẩm mới | Chưa rõ quy trình kiểm soát SP mới có bắt buộc đi qua QLRR; cần template quy trình thẩm định SP mới (thí điểm: Bond KGS) | FI |
| 6 | Đ7 K2f — Văn bản thuộc thẩm quyền TGĐ | Quy trình đo lường/giám sát/báo cáo/xử lý vi phạm hạn mức mới chỉ có ở QC.RR.022 cấp HĐQT, chưa có ở văn bản cấp TGĐ (QD.RR.012) | — |

(Bảng gộp khoản 8/9/11 — dòng tiêu đề Y trùng dòng nội dung, không tính là mục riêng.)

## 2. Đối chiếu với `QC.RR.022_lampd3.html` — LỆCH

Đây là phát hiện chính của thẩm định.

**File HTML hiện chỉ build được Điều 1, 2, 3, 4 (định nghĩa), 9, 10** — xem `id:'d1'…'d3'`
dòng 272–310, nhảy thẳng sang `id:'d9'` dòng 313. **Điều 5, 6, 7, 8 chưa được đưa vào file.**

→ Cả 6 mục GAP thật sự (mục 1–6 ở bảng trên, đều nằm ở Điều 5 và Điều 7) **chưa xuất hiện
trong công cụ đọc**. Người đọc mở file HTML sẽ không thấy các GAP này vì phần nội dung tương
ứng chưa tồn tại.

**Ngược lại**, `QC.RR.022_lampd3.html` có 2 khối `gapbox` (dòng 331 và 346) đánh dấu GAP tại
**Điều 9 Khoản 2b** và **Khoản 3** (phân tách Sổ ngân hàng/Sổ kinh doanh) — **hai GAP này
không có trong sheet `QC.RR.022_MR`/`Gap`** của Excel. Đây là phát hiện mới (khái niệm "hoạt
động ngân hàng truyền thống" không được định nghĩa; cơ chế chuyển sổ qua bên thứ ba không được
quy định) do người đọc bổ sung khi rà từng khoản, nhưng chưa được ghi ngược lại vào bảng GAP
tổng (Excel).

## 3. Kết luận thẩm định

1. **Hai nguồn GAP đang không đồng bộ** — Excel là bảng tổng hợp chính thức (7 mục, Điều 5/7),
   HTML là bản rà chi tiết đang xây (2 mục, Điều 9). Không mục nào trùng nhau. Nếu dùng một
   trong hai như danh sách GAP đầy đủ sẽ thiếu ít nhất 6–7 mục của phía còn lại.
2. **Việc cần làm tiếp** (chưa thực hiện — chờ xác nhận trước khi sửa nội dung):
   - Build tiếp Điều 5, 6, 7, 8 trong HTML và gắn 6 `gapbox` tương ứng theo bảng mục 1.
   - Ghi ngược 2 GAP phát hiện ở Điều 9 (Khoản 2b, Khoản 3) vào sheet `QC.RR.022_MR`/`Gap` của
     Excel để bảng tổng hợp là nguồn duy nhất, đầy đủ.
3. Không sửa nội dung `QC.RR.022_lampd3.html` hay file Excel trong lần thẩm định này — chỉ đọc
   và đối chiếu, để anh xác nhận hướng xử lý trước khi tôi chỉnh nội dung.

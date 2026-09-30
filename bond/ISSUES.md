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

## Đang xử lý

| # | Nội dung |
|---|---|
| E1 | `NotReadableError` chỉ với file key tải từ chat, trên máy nhà của Jak. Đã chứng minh **không phải lỗi code**: một `.xlsx` khác nạp vào cùng HTML thì đọc và giải nén thành công (dừng ở bước kiểm cấu trúc, báo thiếu sheet). File key có kích thước đúng 201,732 byte, cả 3 đường đọc đều `NotReadableError` → Windows cho đọc metadata nhưng chặn nội dung. Nguyên nhân khả dĩ theo thứ tự: file đang mở trong Excel · Downloads đồng bộ OneDrive nên file mới còn là placeholder · phần mềm bảo mật khoá file tải từ Internet. Cách xử lý: đóng Excel · "Always keep on this device" · copy sang `C:\Temp` rồi nạp từ đó. **Sẽ tự hết khi vận hành thật** vì file key do macro sinh tại chỗ, không qua tải mạng |

## Đã kiểm định trước khi viết VBA (17/08)

Chạy bộ trích xuất trên 3 bản File 02: **12/08 · 13/08 · 14/08**.

- Nhãn khối trong `Linked (1)` nằm **đúng cùng một dòng** ở cả ba kỳ
  (Trading nội bộ r4 · Banking nội bộ r31 · TB SBV r54 · BB SBV r61 · Khác r68 · FI Bond r70).
- `Linked` và `Linked (1)` cùng kích thước `A1:GH311` / `A1:N89` ở cả ba.
- File key sinh ra **giống hệt cấu trúc**: 65 · 22 · 18 · 54 · 264 · 52 · 21 · 800 · 17 · 5,990 dòng.
- `asOf` chạy đúng 12 → 13 → 14/08; điểm cuối chuỗi `yield 10Y` cũng tiến theo ngày.
- 6 đoạn nhận định đổi mỗi ngày (`txt.market`, `txt.assessment`, `txt.noteFV`, `txt.itdTB`,
  `txt.note10d`, `txt.noteBB10d`); 10 đoạn còn lại là ghi chú phương pháp luận nên đứng yên —
  đúng như mong đợi.

### Phát hiện

| # | Nội dung |
|---|---|
| E2 | `Report!B254` (`txt.noteRating`) **rỗng ở bản 12/08**, có nội dung ở 13 và 14. Macro phải chịu được ô trống, không được lỗi runtime |
| E3 | `txt.itdBB` (nhận định lỗ MtM Banking) giống hệt cả ba ngày trong khi `txt.itdTB` đổi mỗi ngày. Không phải lỗi công cụ — cần Jak xác nhận có phải quên cập nhật không |
| E4 | Chưa test được **file thưa**. Ba bản này đều đủ dữ liệu. Bản 13/02 mà briefing nhắc (2 deal, cột AK rỗng, từng gây VBA runtime error) vẫn nên chạy thử trước khi chốt macro | **(Jak 30/09: bỏ qua, không xử lý)** |

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

## Trạng thái 18/08 — còn mở

| # | Nội dung | Ai xử lý |
|---|---|---|
| O1 | **Macro chưa chạy trên file thưa.** Bản 13/02 mà briefing nhắc (2 deal, cột AK rỗng, từng gây VBA runtime error) vẫn chưa test. Đây là rủi ro còn lại lớn nhất — mọi bản đã chạy đều là file đủ dữ liệu | Jak chạy thử, gửi lại nếu gãy |
| O2 | **Chuỗi "Sơ cấp / thứ cấp bond" không nạp được từ file key.** Khối trong `Chart data` có cột `STT` đứng ngay sau tiêu đề khối nên bộ đọc lấy STT làm trục, chuỗi ra tên `stt`/`month` không khớp bảng ánh xạ. Biểu đồ vẫn vẽ — **bằng số nhúng sẵn từ 24/07**. Đã thêm cảnh báo đỏ ở tab Quản trị để không im lặng nữa | Jak chuyển cột `STT` ra sau cột `Date` là hết; hoặc báo tôi sửa bộ đọc |
| O3 | `txt.itdBB` (nhận định lỗ MtM Banking) giống hệt nhau ở cả 12 · 13 · 14/08 trong khi `txt.itdTB` đổi mỗi ngày. Không phải lỗi công cụ | Jak xác nhận có phải quên cập nhật không |
| O4 | Dòng dự báo của **tháng báo cáo** trong sheet `VIRA scenarios` còn trống → giả định biến động yield ra −442 bps. HTML đang chặn bằng banner đỏ | Jak điền dự báo |
| O5 | Sheet `Linked` vẫn đọc theo **toạ độ dòng/cột cứng** (POS · CURVE · GRID · SCEN · VIRA). Chèn/xoá dòng ở đó là lệch im lặng. `Linked (1)` thì đã dò theo nhãn nên an toàn | để ngỏ; chuyển sang dò nhãn được nếu muốn |
| O6 | Block "Bình luận / Nhận định" chuẩn `mkCmt*` — chưa dựng | chưa cần |

### Quyết định của Jak (18/08)

Dòng Total của bảng kịch bản VaR **quy hết về `SUM` các ô kỳ hạn** cho thống nhất, không dùng
số từ mô hình VaR mục 21 sheet `Trading`. Hệ quả trên mặt báo cáo, bốn ô sẽ đổi:

| | Trước (mô hình) | Sau (SUM) |
|---|---|---|
| Trading · VaR95% 1 ngày | −23,15 | −0,81 |
| Trading · VaR99% 1 ngày | −48,52 | −29,93 |
| Banking · VaR95% 1 ngày | −25,64 | −6,92 |
| Banking · VaR99% 1 ngày | −57,56 | **+19,12** |

Ô cuối ra số dương trong cột "lỗ tăng thêm" — cần biết trước khi gửi đi.

## Cập nhật 30/09/2026

### Đã xử lý
- **Key lệch 1 dòng (nguyên nhân số sai ở Trạng thái, Thời gian nắm giữ, PV01, Lãi/lỗ, Sử dụng vốn).** Từ File 02 ngày 29/09, sheet `Linked` bị chèn thêm 1 dòng. `make_key.py` và `XuatFileKey.bas` đọc theo toạ độ cố định nên nuốt dòng tiêu đề vào dữ liệu và cắt mất dòng Tổng cuối khối (POS Trading, 3.3, 3.4, 3.5, 3.7, 3.8, 7.1). Riêng khối Xếp hạng bị cắt 3 issuer cuối (VCCB, PVFC, OJBV rating D) và dòng Total vì danh sách dài thêm. Cả hai công cụ nay **dò theo tiêu đề bảng** (3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8, "Chỉ tiêu" của 7.1, "Issuer"). Đã kiểm hồi quy trên File 02 ngày 17/08: kết quả không đổi. `XuatFileKey.bas` **chưa chạy thử trong Excel** (không có VBA ở đây), cần chạy thử một lần.
- Sheet POS thêm cột `YtdPnl` (cột AA của Trading) để bảng "PnL theo kỳ hạn" có YtD PnL.
- HTML nạp key: bỏ dòng tiêu đề bị đọc nhầm, tự cộng Tổng khi key thiếu (POS, Thời gian nắm giữ, Sử dụng vốn) và báo đỏ ở "Kiểm tra trước khi gửi". Tổng PnL và Tổng kịch bản không tự cộng được nên hiện `—`, không dùng số cũ.
- **Biểu đồ vẽ bằng số cũ nhúng sẵn.** Tên trường trong key (theo tiêu đề cột Excel) khác tên trường biểu đồ đọc, phần thiếu bị điền bằng số 24/07. Đã thêm bảng đổi tên `ALIAS`; trường nào vẫn thiếu thì để trống và báo lỗi, không điền số cũ. Thêm ánh xạ `month` → `primary` (sơ cấp/thứ cấp).
- Biểu đồ Sử dụng vốn: cột 4 (TT41) = `Vốn yêu cầu − Lỗ tiềm ẩn`. Trong Chart data, "Vốn yêu cầu" đã gồm ItD nên trước đây ItD bị tính hai lần.
- Biểu đồ QLHS ở tab Lãi/lỗ chỉ còn GovBond; phần FI Bond & CD chuyển sang trang FI Bond & CD.

### Cần Jak quyết (phía Excel, HTML không tự sửa được)
| # | Nội dung |
|---|---|
| E1 | Sheet `Chart data`, cột DE "Realized&NIM lũy kế năm" = **199.9 (hằng số, không đổi theo ngày)**, trong khi bảng 3.8 dòng (1) = **59.2** (= −74.8 + 134.0). Tổng trên biểu đồ (2,733) vì thế khác Tổng trên bảng (2,592.5) đúng 140.7 |
| E2 | Bảng 3.5: `Total PnL` YtD −1,275.1 không bằng (1)+(2) = −1,214.4; MtD −109.6 không bằng −251.1. Cần xác nhận định nghĩa dòng Total |
| E3 | Chuỗi `lsttSoCap` (LS sơ cấp) dừng ở 19/08, chuỗi QLHS dừng ở 01/08 |

### Hoãn (đã thống nhất, chưa làm)
- Chọn ngày so sánh (filter) trên báo cáo. Lý do: chưa có dữ liệu lịch sử trong file, kéo hết lên sẽ nặng.

### Bổ sung 30/09/2026 (chiều)
- **Excel đảo ngày/tháng khi macro ghi key** (ví dụ 03/08 thành 08/03, chỉ với ngày ≤ 12): macro ghi ngày dạng chuỗi vào ô định dạng chung nên Excel đọc theo locale. Trục ngày biểu đồ vì thế sai. `XuatFileKey.bas` nay đặt cột ngày (META, VIRA4, RATING, TS) là định dạng Text trước khi ghi. HTML tự báo đỏ khi ngày trong chuỗi không tăng dần. `make_key.py` không bị (ghi chuỗi thẳng).
- "Kiểm tra trước khi gửi" chỉ hiện khi có lỗi/cảnh báo; ẩn Volatility trên báo cáo; tiêu đề cột ngày thống nhất dd/mm/yyyy.

### Rà số 30/09/2026 (đợt 3–4)
Đối chiếu HTML (nạp `Key_20260929.xlsx`) với File 02 ngày 29/09, đọc thẳng từ sheet `Linked` / `Linked (1)` / `Chart data`:
- **731** giá trị cấp dữ liệu (Position, Nắm giữ, Cơ cấu phát hành, PnL, Sử dụng vốn, BS, FI trên BS, Kịch bản, đường cong, Xếp hạng, Khối 2) và **460** ô hiển thị thật trên bảng (Trạng thái, ItD, PV01, Nắm giữ, Vốn, PnL, Kịch bản, FI): **0 lệch**.
- **6,515** điểm chuỗi biểu đồ (21 chuỗi trong `Chart data`): **0 lệch**, riêng chuỗi `lsttSoCap` có 5 ngày cuối (26/08–23/09) trống trong Excel nên không vẽ.
- **Lỗi tìm thấy và đã sửa:** bảng "FI Bond trên bảng cân đối" hiện dòng "4. Giảm trừ đối ứng" hai lần (dòng bắt đầu bằng "4." bị lọc nhầm vào phần thân bảng).
- Trục nhãn chuỗi theo tháng: `month` (sơ cấp/thứ cấp) trước đây hiện dd/mm, nay hiện mm/yy.

| # | Cần Jak xác nhận (phía Excel) |
|---|---|
| E4 | Sheet `Chart data`, khối "Sơ cấp/thứ cấp bond" (cột AX–AZ): hai dòng **01/03/2026 và 29/03/2026 trùng hoàn toàn số liệu** (19,560 và 272,650), các tháng sau lệch sang ngày 29. Nghi dòng bị sao chép; HTML đã báo cảnh báo khi nạp | **(Jak 30/09: bỏ qua, không xử lý)** |

### Kế hoạch còn lại
| Đợt | Nội dung | Trạng thái |
|---|---|---|
| 5 | Tên chỉ số đầy đủ | Đã lập bảng đề xuất `TEN_CHI_SO_DE_XUAT.md`, **chờ Jak duyệt** |
| 6 | Rules cảnh báo biến động so với ngày hôm trước ở "Kiểm tra trước khi gửi" | **Hoãn (Jak 30/09: log lại, chưa làm).** Cần: báo cáo mới nhất và ngưỡng cho từng chỉ số. Đã có sẵn: khung kiểm tra và cảnh báo hạn mức; tính phân bố biến động ngày từ lịch sử `Chart data` để đề xuất ngưỡng |


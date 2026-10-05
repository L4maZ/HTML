# XuatRatingKey.bas

Sinh `Rating_Key_YYYYMMDD.xlsx` (schema `rating.key.v1`) từ workbook `01.FIBond_Rating`
để upload vào `Rating_Dashboard_v0.10.html`.

## Dùng

1. Mở workbook (`.xlsm` / `.xlsb`) → `Alt+F11` → *File → Import File…* → chọn `XuatRatingKey.bas`.
2. Refresh + Calculate, xem `Report!C10`.
3. `Alt+F8` → `XuatRatingKey` → Run. File Key nằm cùng thư mục workbook
   (đổi hằng `OUT_FOLDER` đầu module nếu muốn thư mục khác).
4. Mở file Key, nhập nhận định ở sheet `TEXT`, cột **Override Value** (cột `Value` = Override nếu có, không thì Auto).
5. Upload file Key ở tab *Quản trị dữ liệu* của Dashboard.

File `.bas` chỉ chứa ASCII (tiếng Việt có dấu trong dữ liệu ghi ra được dựng bằng `ChrW`) để import không bị lỗi font.

## Macro đọc gì

| Sheet Key | Nguồn trong workbook | Nhận diện bằng |
|---|---|---|
| META | `RPT` (name), `Report!C10`, các tổng | name `RPT`; dòng `Check` ở `Report!B:C` |
| EVENTS (RATING_CHANGE) | `Report!G:L`, khối 1 | tiêu đề `1.` ở cột G; dữ liệu từ tiêu đề + 2 |
| EVENTS (REVIEW_CHANGE) | `Report!G:K`, khối 2 | tiêu đề `2.`; dữ liệu từ tiêu đề + 2 |
| EVENTS (NEW_OR_REMOVED_RATING) | `Report!G:Q`, khối 3 | tiêu đề `3.`; dữ liệu từ tiêu đề + 3 (2 dòng header) |
| RATING_STRUCTURE, KPI | `Portfolio!A:H` | ô `Internal Rating` → `Grand Total` |
| TENOR | `Portfolio!A:C` | ô `Tenor` → `Total` |
| PORTFOLIO | `Portfolio_by_bond` + cột AH:AK | header `Bond_Code`, cửa sổ `Bond_Code-2 … +12` |

Dòng Report có `Issuer name` (cột H) rỗng bị bỏ qua (XLOOKUP/FILTER trả 0 hoặc "").

## Dừng xuất (HTML sẽ từ chối)

Thiếu `RPT` / khối Report / bảng Portfolio · rating trong Portfolio_by_bond không có trong bảng cơ cấu ·
tổng Portfolio hoặc Tenor lệch > 0,01 · `AmountAfterCpty > Amount` · thiếu Issuer/BondCode/Rating.

## Chỉ cảnh báo (hỏi có xuất tiếp không)

`Report!C10 <> OK` · tổng theo rating lệch bond-level · dòng Total sau đối ứng lệch tổng chi tiết ·
lũy kế rating không khớp Outstanding. Đây chính là các cảnh báo vàng HTML hiện ở tab Quản trị dữ liệu.

## Kiểm chứng đã làm / chưa làm

- Đã: port logic ánh xạ sang Python, đối chiếu với dữ liệu nhúng sẵn trong HTML → trùng 100%.
- Đã: dựng file Key đúng layout, upload vào HTML thật bằng Chromium → 0 lỗi.
- **Chưa:** chạy macro trong Excel thật (môi trường dựng macro không có Excel). Chạy lần đầu trên bản copy
  của workbook và mở file Key kiểm tra.

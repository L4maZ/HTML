# Bond — Tool → File Key → HTML (khối 2)

Bản thử nghiệm cơ chế: File 02 sinh ra một file key nhỏ, HTML nạp file key đó và render.
Phạm vi hiện tại **chỉ khối 2** (52 chỉ tiêu tuân thủ hạn mức) để kiểm tra tính năng.

## File trong thư mục

| File | Vai trò |
|---|---|
| `BaoCao_RRTT_Bond_v8_upload.html` | HTML nạp file key, render khối 2, xuất bản gửi đi. **22KB** |
| `Key_20260813.xlsx` | File key mẫu, sinh từ File 02 ngày 13/08. **16KB** |
| `tools/XuatFileKey.bas` | Macro VBA để tự sinh file key trên máy bank |
| `tools/make_key.py` | Bản Python tương đương, dùng khi có Python |

## Quy trình

```
File 02 (.xlsm)
  ├─ Linked (1)  A1:N80   → 52 chỉ tiêu khối 2
  ├─ Report      6 ô chữ  → nhận định, ghi chú
  └─ Run Tool    B2:B7    → trục ngày
        │  XuatFileKey
        ▼
  Key_YYYYMMDD.xlsx  (META · DATA_S2 · TEXT) — 16KB
        │  nạp qua nút trong HTML, parse tại chỗ
        ▼
  BaoCao_RRTT_Bond_v8_upload.html
        │  Xuất bản gửi đi
        ▼
  BaoCao_RRTT_Bond_DDMMYYYY.html — 70KB, gửi mail được
```

## Cài macro (một lần)

1. Mở File 02 → `Alt+F11` → chuột phải VBAProject → **Import File** → chọn
   `tools/XuatFileKey.bas`.
2. Chạy `TaoCotKeyID` **một lần**. Macro ghi cột `KeyID` vào `Linked (1)` cột **N**.
3. Từ đó về sau mỗi kỳ chỉ chạy `XuatFileKey`. File key nằm cùng thư mục File 02.

Chèn thêm dòng trong `Linked (1)` thì chạy lại `TaoCotKeyID` để cấp key cho dòng mới.

## Vì sao có cột KeyID

`KeyID` là **neo**. HTML tra cứu theo chuỗi key, không theo toạ độ ô. Chèn/xoá/đảo dòng
trong Excel không làm HTML đọc lệch — dòng nào không có key thì không ra báo cáo, thay vì
âm thầm lấy số của dòng bên cạnh.

## Nguyên tắc đã cài trong HTML

- **Thiếu số thì gãy to.** Ô trống hiện `—` đỏ, không bao giờ lấy số kỳ trước.
- **Banner ngày chốt** luôn hiện, kèm tên file nguồn và giờ sinh file.
- **Kiểm tra cấu trúc**: sai `schema`, trùng `KeyID`, thiếu sheet → banner đỏ liệt kê lỗi.
- **Offline hoàn toàn.** File key được parse ngay trong trình duyệt bằng `DecompressionStream`,
  không thư viện ngoài, không gửi đi đâu. Google Fonts chặn thì degrade về Segoe UI.
- **Không `localStorage`.** Bản xuất mang dữ liệu trong `window.RRTT_SNAPSHOT`.
- Key text vẫn theo cơ chế `data-k` + `KEY_DEF`: xoá sạch key thì chữ gốc quay lại.

## Đã kiểm

Chromium headless, mở bằng `file://`, ngắt mạng:

- Nạp `Key_20260813.xlsx` → 65 dòng, 7 section, 18 đèn hạn mức.
- Face Value TB nội bộ 17,513 = AFS 12,646 + HTM 4,867 — khớp `Linked (1)`.
- Hạn mức đầu tư ra nước ngoài 15.0/15m USD → đèn **đỏ 99.9%**.
- Ô rỗng trong Excel → `—`, không nội suy.
- Xuất bản gửi đi → 70KB, mở lại đủ 65 dòng, ẩn thanh công cụ, trạng thái "chỉ đọc".
- Google Fonts bị chặn (`ERR_CONNECTION_RESET`) → trang vẫn dựng bình thường.

## Giới hạn tuyệt đối

**File HTML gửi đi ≤ 5MB** (mail bank chặn quá 5MB). Hiện 70KB. Khi thêm ECharts inline
cho các khối 3.x/4/6.x sẽ tăng ~940KB — vẫn còn dư địa lớn.

## Chưa làm

- Khối 3.x · 4 · 6.x · 7.x (đọc từ sheet `Linked`, cần ECharts)
- 19 chuỗi time-series từ `Chart data`
- Block "Bình luận / Nhận định" chuẩn `mkCmt*`

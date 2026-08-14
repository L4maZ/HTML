# Key-ification Standard

Áp dụng cho **mọi HTML build mới**. Mục tiêu: người dùng nghiệp vụ sửa được text và
tham số phương pháp luận qua Excel, không cần đụng vào HTML.

## 3 tier

### Tier 1 — luôn key (~30–40 keys)

- Tiêu đề trang, subtitle
- Footnote (`*`)
- KPI labels
- Button labels
- Disclaimer

### Tier 2 — chọn lọc

Chỉ key những tham số phương pháp luận **có thể bị challenge**:

- Ngưỡng z-score
- MA windows: 252 / 504 / 756
- Ngưỡng cảnh báo
- Ngưỡng limit

### Tier 3 — không bao giờ key

- Time-series data → đi qua Excel upload vào RAW object
- Màu của từng component

## Rule bắt buộc

- Thiếu key → **fallback về text hardcode hiện tại**. Không bao giờ render `undefined`.
- Xóa sạch `Key_Config` thì file vẫn phải chạy bình thường.

```js
function K(key, fallback) {
  var v = KEYS[key];
  return (v === undefined || v === null || v === '') ? fallback : v;
}
// dùng: el.textContent = K('pgTitle', 'Bond Portfolio Risk Report');
```

## Excel `Key_Config` — 9 cột

`Key | Value | Nhom | Trang | Vi tri tren giao dien | Loai key | Type | Vi du gia tri hop le | Ghi chu`

Format:

- Cột **Value** = chữ xanh — cột duy nhất được sửa.
- Ô `AUTO_UNLESS_OVERRIDE` = fill vàng.
- Header: nền `#B83A10`, chữ trắng.
- Freeze pane tại **B5**.
- Dropdown validation ở cột **Type**.
- Dòng tổng hợp `COUNTA` / `COUNTIF` ở cuối sheet.

## Excel `Key_Map` — 6 cột

`Nhom | Key | Phan tu tren HTML | Vi tri cu the | Khoi du lieu Sheet1 lien quan | Doi key nay thi doi gi`

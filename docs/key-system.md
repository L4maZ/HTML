# Key-ification Standard

Mục tiêu: người dùng nghiệp vụ sửa được text và tham số phương pháp luận qua Excel, không cần
đụng vào HTML.

Bản tham chiếu đang chạy: `BaoCao_RRTT_Bond_key_v7.html` (dòng 119–184).

## 3 tier

### Tier 1 — luôn key (~30–40 keys)

Tiêu đề trang, subtitle, footnote (`*`), KPI labels, button labels, disclaimer.

### Tier 2 — chọn lọc

Chỉ key những tham số phương pháp luận **có thể bị challenge**: ngưỡng z-score, MA windows
252 / 504 / 756, ngưỡng cảnh báo, ngưỡng limit.

Ví dụ thật trong RRTT v7: `thrWarn: 0.8`, `thrBreach: 1` (đèn XANH < 80% · HỔ PHÁCH 80–100% ·
ĐỎ ≥ 100%).

### Tier 3 — không bao giờ key

Time-series data (đi qua Excel upload vào RAW object) và màu của từng component.

## Cơ chế: hai tầng fallback

Đây là điểm cốt lõi khiến rule "không bao giờ `undefined`" thành sự thật, không phải lời hứa.

### Tầng 1 — `KEY_DEF`: object hằng chứa giá trị mặc định

```js
const KEY_DEF = {
  rptTitle: 'Báo cáo rủi ro thị trường', rptSubtitle: 'Báo cáo Desk Bond',
  navP1: 'Highlight', navP2: 'Chi tiết danh mục',
  /* nội dung nhận định — Jak nhập hằng ngày, mặc định rỗng */
  hiCompliance: '', hiTb: '', hiBb: '', hiFi: '', hiRiskPrice: '', hiRiskNim: '',
  /* tham số ngưỡng — Tier 2 */
  thrWarn: 0.8, thrBreach: 1
};
window.RRTT_KEYS = window.RRTT_KEYS || {};
```

### Tầng 2 — `data-k` trên DOM: chữ gốc trong HTML là fallback cuối

Lần render đầu ghi nhớ `textContent` gốc vào `dataset.kDef`. Nhờ vậy **xóa key trong Excel thì
chữ gốc quay lại**, không bao giờ ra rỗng.

```js
function applyDataKeys(root) {
  (root || document).querySelectorAll('[data-k]').forEach(el => {
    const k = el.dataset.k;
    if (el.dataset.kDef == null) el.dataset.kDef = el.textContent;   // chốt chữ gốc 1 lần
    const v = kv(k, el.dataset.kDef);
    if (v != null && String(v).trim() !== '' && el.textContent !== v) el.textContent = v;
  });
}
```

### Accessor

```js
function kv(k, fb) {
  const o = window.RRTT_KEYS || {}, v = o[k];
  if (v === undefined || v === null || String(v).trim() === '') {
    if (fb !== undefined && fb !== null && String(fb).trim() !== '') return fb;
    const d = KEY_DEF[k];
    return (d === undefined || d === null) ? '' : d;    // cùng lắm là chuỗi rỗng, không undefined
  }
  return v;
}
function kvNum(k, fb) {
  const v = window.RRTT_KEYS && window.RRTT_KEYS[k];
  const n = parseFloat(v);
  return isNaN(n) ? (fb != null ? fb : KEY_DEF[k]) : n;
}
```

Chuỗi ưu tiên: `RRTT_KEYS[k]` (Excel) → `fb` truyền vào → `dataset.kDef` (chữ gốc HTML) →
`KEY_DEF[k]` → `''`.

### API ngoài

```js
window.RRTT_keyList = function () {      // xuất toàn bộ key + chữ gốc, để dựng sheet Excel
  const o = {};
  document.querySelectorAll('[data-k]').forEach(el => {
    o[el.dataset.k] = el.dataset.kDef != null ? el.dataset.kDef : el.textContent;
  });
  return o;
};
window.RRTT_setKeys = function (obj) {   // nạp key từ Excel rồi vẽ lại
  window.RRTT_KEYS = obj || {};
  if (window.RRTT && window.RRTT.rerender) window.RRTT.rerender();
};
```

**Về con số "92 keys"**: file v7 có **58 phần tử `data-k`** trong DOM, cộng các key chỉ tồn tại
trong `KEY_DEF` (nav, ngưỡng, nhận định) mà không gắn `data-k`. Khi đếm phải nói rõ đang đếm
tầng nào.

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

## Kiểm thử bắt buộc

Trước khi giao file: chạy `RRTT_setKeys({})` trong console. Báo cáo phải hiện **nguyên vẹn**
chữ gốc, không một chỗ nào rỗng hay `undefined`.

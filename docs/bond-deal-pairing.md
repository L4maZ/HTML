# Ghép cặp deal Bond — phương pháp và bẫy

Cách tái tạo quan hệ cặp repo từ export deal `MSB_RP_DM`, và ý nghĩa nghiệp vụ của
từng bước. Viết từ lần dựng `Phan_tich_GD_Bond_20260608_20260820.html`.

Script: [`bond/tools/ghep_cap_deal_bond.py`](../bond/tools/ghep_cap_deal_bond.py)
Dữ liệu mẫu: `bond/source/Dealps_0806_2008.xlsx` (346 chân, CaptureDate 08/06–20/08/2026)

```
python3 bond/tools/ghep_cap_deal_bond.py bond/source/Dealps_0806_2008.xlsx \
        --html bond/Phan_tich_GD_Bond_20260608_20260820.html
```

Không có `--html` thì chỉ in thống kê. Có `--html` thì thay khối `const D={...};`
tại chỗ — phần trình bày không đụng tới.

---

## Vấn đề gốc

File export là danh sách **phẳng**, mỗi dòng một **chân** giao dịch, chỉ có `DealType`
là `B` hay `S`. **Không có cột nào** nối chân bán với chân mua lại của cùng một hợp đồng
(không `RepoID`, không `GroupID`).

Không biết chân nào đi với chân nào thì không tính được lãi lỗ — từng chân riêng lẻ chỉ
là một dòng tiền vào hoặc ra. Chỉ khi ghép cặp mới có "bán 100 tỷ, mua lại 100,05 tỷ sau
3 ngày ⇒ chi phí 50 triệu cho 3 ngày".

Nên bước đầu bắt buộc là **tái tạo quan hệ cặp mà hệ thống không lưu**.

> **Ghép tay là cách ĐÚNG, không phải giải pháp tạm.** Xác nhận của Jak (08/2026): đội kinh
> doanh **tách một deal repo thành hai deal outright** khi nhập máy, nên trường nối cặp của hệ
> thống bị cherry-picking — không khớp điều kiện repo/reverse repo và mapping loạn. Vì vậy
> **không** lấy trường hệ thống làm chuẩn; ghép lại từ điều kiện nghiệp vụ như dưới đây.
>
> Hệ quả kép: vì mỗi chân là một lệnh outright riêng, **mỗi chân được định giá theo yield thị
> trường của chính ngày giao dịch đó**, không theo một lãi suất repo thỏa thuận. Xem mục
> "Giá vốn repo" cuối tài liệu — đây là điều chi phối toàn bộ cách đọc lãi/lỗ.

## Điều kiện ghép cặp

| Điều kiện | Vì sao |
|---|---|
| Cùng `Bonds_ShortName` | Repo là bán rồi mua lại **đúng mã đó** |
| Cùng `Cpty_ShortName` | Hợp đồng repo là với **một** đối tác |
| Ngược chiều (một S, một B) | Bản chất của repo |
| Cùng `Quantity` | Điều kiện khoá — chính là hiện tượng **offset khối lượng** quan sát được |

Bốn điều kiện lọc rất chặt, nhưng vẫn mơ hồ khi trong cùng mã + cùng đối tác có nhiều
cặp trùng ngày. Nên chạy hai lượt.

### Lượt 1 — khôi phục cặp theo dấu vết book lệnh

Ưu tiên ghép hai chân **cùng `CaptureDate`** và **cùng `Quantity`**, chọn cặp có
`BondsDeals_Id` gần nhau nhất. Chỉ ghép khi cả hai chân còn nguyên khối lượng — lượt này
không tách chân.

Lý do: hai vế của một hợp đồng repo được **nhập vào hệ thống cùng lúc**, nên dính liền
nhau cả về ngày nhập lẫn số hiệu deal.

**Tiêu chí này thắng thứ tự thanh toán thuần túy.** Ví dụ thật: bốn chân TD1646468 với
CK_AGRIBAN — hai chân S cùng thanh toán 46237, hai chân B thanh toán 46238 và 46245.
Ghép theo ngày thanh toán sẽ nối 50269 với 50272; ghép theo dấu vết book lệnh cho
50269↔50270 và 50271↔50272, đúng với cách hai lệnh được nhập.

### Lượt 2 — FIFO phần dư

Phần còn lại ghép theo thứ tự ngày thanh toán, tách chân lớn nếu khối lượng lệch (bán
1 lệnh 3.000 tỷ, mua lại bằng 2 lệnh 1.500 tỷ).

### Lượt 3 — ghép chéo đối tác

Nới điều kiện "cùng đối tác" cho phần dư cuối: cùng mã, cùng tổng khối lượng, sát ngày.
Gắn tag `est` để tách riêng khi trình bày.

Trường hợp thực tế duy nhất trong kỳ: deal 50120 mua của **KBNN** qua kênh liên ngân hàng
`IB-P-VSD`, bán lại cho **PGBV-HO** qua `OT-P-VSD` bằng hai lệnh 50127 + 50128.

**Cặp loại này không phải repo theo nghĩa pháp lý** — MSB mua đứt rồi bán đứt, không có
quyền đòi lại tiền. Chân bán hỏng là ôm nguyên trái phiếu trên sổ. Phải tách khỏi nhóm
repo khi đọc rủi ro.

## Chia Nhóm A / Nhóm B: theo chân nào thanh toán TRƯỚC

Cặp nào cũng có một chân S và một chân B. Cái phân biệt là **thứ tự thời gian**:

**Nhóm A — chân S thanh toán trước** (bán trước, mua lại sau)
MSB giao trái phiếu nhận tiền, ít ngày sau trả tiền lấy trái phiếu về.
MSB cầm tiền của đối tác ⇒ **MSB đi vay**. Trái phiếu là tài sản bảo đảm. Phần trả thêm
khi mua lại là **lãi vay**.

**Nhóm B — chân B thanh toán trước** (mua trước, bán lại sau)
MSB bỏ tiền nhận trái phiếu, sau đó nhận tiền trả trái phiếu.
Đối tác cầm tiền của MSB ⇒ **MSB cho vay**. Phần thu thêm là **lãi cho vay**.

Nói gọn: **ai giữ tiền của ai trong khoảng giữa hai ngày thanh toán**. Không xét S hay B,
mà xét chân nào đến trước.

Dùng `SettlementDate` chứ **không** dùng `TradeDate` — tiền chuyển thật vào ngày thanh
toán. Trong kỳ mẫu chỉ 25/173 cặp có hai chân cùng `TradeDate`, dùng `TradeDate` sẽ tính
sai kỳ hạn.

## Quy ước dấu

Một công thức duy nhất cho cả hai nhóm, **không đảo dấu**:

```
cost = tiền chân B − tiền chân S
lãi suất ngụ ý = cost ÷ tiền chân đầu ÷ số ngày × 365
```

- Nhóm A: `cost` **dương** = mua lại đắt hơn bán ra = MSB **trả lãi** (chi phí)
- Nhóm B: `cost` **âm** = bán lại nhiều hơn mua vào = MSB **thu lãi**

Phía trình bày mới đổi dấu cho dễ đọc (dương = có lợi cho MSB). Giữ tính toán một chiều
để không sai dấu ở đâu.

Lãi suất tổng hợp theo nhóm dùng **trọng số tiền × ngày**, không phải trung bình cộng —
nếu không thì một deal 1 ngày 30 tỷ nặng ngang một deal 351 ngày 1.000 tỷ.

## Kết quả kiểm chứng (kỳ 08/06–20/08/2026)

| | |
|---|---|
| 346 chân → **173 cặp** | 170 cặp ghép nguyên chân, 2 cặp phải tách, 1 cặp ghép chéo |
| **0 chân lẻ** | Toàn bộ khối lượng offset hết |
| **171/173** cặp có `CaptureDate` giống nhau cả hai chân | Ghép bừa thì con số này phải rời rạc |
| 136 cặp của bản `..._20260810` tái lập **khớp từng đồng** | Hai cách làm độc lập ra cùng kết quả |

---

## Bốn cái bẫy

### 1. `GrossAmount` có dòng là string

16/346 dòng ở dạng `'  1012036120.0000000K'` — **đơn vị nghìn**, không phải số. Không
parse riêng là lệch 1000 lần trên đúng những deal to nhất. Cần báo đội hệ thống.

```python
def parse_gross(v):
    if isinstance(v, str):
        return float(v.strip().rstrip('K')) * 1000.0
    return float(v)
```

### 2. Ngày là serial Excel

Gốc **1899-12-30**, không phải 1900-01-01.

### 3. Ngưỡng ngoại lệ phải theo kỳ hạn

Yield hai chân lệch ≥ 10bp trên **repo ngắn (≤ 30 ngày)** nghĩa là chân mua lại bị mark
theo thị trường — lãi/lỗ khi đó là **rủi ro giá**, không phải chi phí vốn, phải loại khỏi
số giá vốn. Trộn vào thì chi phí vay nhóm A đọc thành 3,80%/năm thay vì 2,66%.

Nhưng **cùng ngưỡng đó trên deal 351 ngày là vô nghĩa** — yield lệch nhau sau gần một năm
là bình thường, loại đi sẽ bóp méo lợi suất nhóm B. Ngưỡng phải kèm điều kiện kỳ hạn.

### 4. `%/năm` trên deal qua đêm không dùng được

Chênh lệch vài chục triệu nhân 365 ra con số hoang đường. Trong kỳ mẫu, 17 cặp 1–2 ngày
có `|%/năm| > 8`.

**Nguyên nhân KHÔNG phải làm tròn giá** (cách đọc ban đầu, đã sửa) — mà là **chênh lệch yield
giữa hai chân lấn át lãi repo**. Bóc cặp 49954/49955 (1 ngày, 148 tỷ):

```
Δyield chỉ +0,56bp  ->  giá giảm 39 đ/đơn vị  ->  -58,5 tr
accrual 1 ngày                                 ->  +10,5 tr
                                       gross chênh  -48,0 tr
lãi repo qua đêm đúng ra @3%/năm                ~  12 tr
```

Hiệu ứng yield lớn **gấp ~5 lần** lãi repo thật. Nên trên kỳ hạn ngắn, **dấu của lãi/lỗ gần như
do yield quyết định, không do lãi vay**. Bằng chứng: tỷ lệ cặp nhóm A "có lãi" giảm đều theo
kỳ hạn — 47% ở 1 ngày, 34% ở 2–7 ngày, 8% ở 8–30 ngày, **0% ở 31–90 ngày**. Kỳ hạn càng dài,
lãi repo càng lấn át nhiễu và mọi cặp đi vay đều lỗ đúng bản chất.

Tương quan giữa lãi/lỗ (trên mỗi tỷ đồng) và Δyield: **0,85–0,99 ở mọi nhóm kỳ hạn**.

Với deal qua đêm phải đọc bằng **số tiền tuyệt đối**, không đọc %/năm.

## Giá vốn repo = chính lợi suất trái phiếu đem cầm cố

Lọc 44 cặp có hai chân **cùng một mức yield** (`|Δy| < 0,05bp`) — tức chênh lệch tiền là
accretion thuần, không lẫn biến động giá. Với nhóm này, lãi suất ngụ ý **bám sát yield của
chính trái phiếu đó**: lệch trung vị **−0,08 điểm %**, 41/44 cặp nằm trong ±0,5 điểm %.

Nghĩa là repo ở đây không định giá theo lãi suất thị trường tiền tệ, mà theo **carry của trái
phiếu đem ra cầm cố**. Hệ quả khi đọc số:

- Chi phí vay của MSB **do việc chọn mã TP quyết định**, không hẳn do đối tác. Repo bằng mã
  yield 3% thì vay rẻ, bằng mã yield 4,5% thì vay đắt — cùng một đối tác.
- Khi so chi phí giữa các đối tác, phải đặt cạnh **yield trung vị của mã đem repo**, nếu không
  là quy nhầm chênh lệch cho đối tác.
- Với 129 cặp còn lại (hai chân khác yield), **lãi suất repo thỏa thuận không khôi phục được
  từ dữ liệu này** — nó bị chôn dưới biến động giá. Muốn biết giá vốn thật phải lấy từ hợp
  đồng repo hoặc sổ của desk.

## Lãi đã thực hiện ≠ lãi theo hợp đồng

Nhóm B kỳ hạn trung vị 154 ngày, phần lớn **đáo hạn sau ngày chốt sổ**. Cộng gộp là báo
sai P&L trong kỳ:

| | Kỳ 08/06–20/08/2026 |
|---|---|
| Đã tất toán (111/173 cặp) | **−4,4 tỷ** |
| Còn hiệu lực (62 cặp), theo hợp đồng | +271,8 tỷ |

Số thứ hai là **cam kết**, ghi nhận dần tới khi chân sau thanh toán (xa nhất là 2027).
`blk()` trong script trả cả `pnlDone` lẫn `pnlOpen` để không lẫn hai thứ.

Mốc chia nằm ở hằng số `AS_OF` — **đổi khi kỳ báo cáo đổi**.

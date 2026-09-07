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
| Cùng `Cpty_ShortName` | Hợp đồng repo là với **một** đối tác (trừ deal nghiệp vụ chỉ đích danh) |
| Cùng `CouponRate` | Nghiệp vụ yêu cầu. Thực tế coupon suy ra từ mã TP (42/42 mã chỉ một coupon) nên không đổi kết quả — giữ để dữ liệu sau này có mã nhiều coupon thì tự tách |
| Ngược chiều (một S, một B) | Bản chất của repo |
| Cùng `Quantity` | Hiện tượng **offset khối lượng** quan sát được |

**Khối lượng lệch không phải là chân lẻ.** Một hợp đồng có thể được nhập thành
**1 deal B đối ứng 2 deal S** (hoặc ngược lại). `merge_multi_leg()` gộp các mảnh dùng chung
một chân và cùng kỳ hạn về **một dòng**, đúng như cách nghiệp vụ nhìn. Trong kỳ mẫu có
2 trường hợp: B 50100 ↔ S 50099 + S 50103 (CK_AB), và B 50120 ↔ S 50127 + S 50128 (KBNN).

Các điều kiện trên lọc rất chặt, nhưng vẫn mơ hồ khi trong cùng mã + cùng đối tác có nhiều
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

### Lượt 3 — ghép chéo đối tác (chỉ cho deal đã được xác nhận)

Nới điều kiện "cùng đối tác" cho phần dư cuối: cùng mã, cùng tổng khối lượng, sát ngày.
Gắn tag `est` để tách riêng khi trình bày.

Trường hợp thực tế duy nhất trong kỳ: deal 50120 mua của **KBNN** qua kênh liên ngân hàng
`IB-P-VSD`, bán lại cho **PGBV-HO** qua `OT-P-VSD` bằng hai lệnh 50127 + 50128.

**Đây VẪN LÀ REPO.** Xác nhận của Jak (08/2026): đối tác làm việc với chính phủ, MSB đứng
**trung gian**. Hai chân khác đối tác là do vai trò trung gian, không phải mua đứt bán đứt.
**100% deal trong file là repo** — không có ngoại lệ.

> **Lượt 3 là ngoại lệ có kiểm soát, không phải quy tắc.** Quy tắc ghép là **cùng mã TP +
> cùng đối tác**, áp cho mọi deal; chỉ nới ra với deal mà nghiệp vụ đã chỉ đích danh. Nếu kỳ
> sau script tự ghép chéo một cặp chưa ai xác nhận thì **phải hỏi lại nghiệp vụ trước khi
> dùng số**, vì ghép chéo nhầm sẽ tạo ra một cặp không có thật mà vẫn khớp khối lượng.

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

Lãi/lỗ là **chênh lệch Gross Amount hai chân theo thứ tự thanh toán**, không trừ gì thêm:

| Nhóm | Chân 1 | Lãi/lỗ |
|---|---|---|
| A · đi vay | tiền **vào** | Gross chân 1 − Gross chân 2 |
| B · cho vay | tiền **ra** | Gross chân 2 − Gross chân 1 |

**Dương = có lợi cho MSB ở cả hai nhóm.** Hai biểu thức trên rút gọn về cùng một thứ:
`−cost`, với `cost = gross B − gross S`. Đó là `pnl_mn()` trong script.

Lãi suất ngụ ý = lãi/lỗ ÷ tiền chân đầu ÷ số ngày × 365, cùng dấu với lãi/lỗ. Tổng hợp theo
nhóm dùng **trọng số tiền × ngày**, không phải trung bình cộng.

> **Không loại trừ cặp nào.** Mọi con số tổng hợp tính đủ 172 cặp. Các danh sách "Δyield lớn",
> "%/năm lớn trên deal ngắn", "lệch folder" chỉ để **tra cứu**, không phải để trừ ra.

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

### 3. Mặt bằng Δyield khác nhau theo kỳ hạn

|Δyield| trung vị trong dữ liệu: **0,2bp** ở deal ≤ 90 ngày, **3,2bp** ở deal > 90 ngày. Nên
cùng một ngưỡng 10bp mang ý nghĩa hoàn toàn khác nhau ở hai nhóm — ở deal ngắn là cách mặt
bằng 50 lần, ở deal dài là chuyện thường.

Vì vậy `is_dy_big()` kèm điều kiện `days <= 30`. Đây chỉ là **tiêu chí liệt kê để tra cứu**,
không loại khỏi bất kỳ con số tổng hợp nào.

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


## Phụ lục — FIBond (CD / trái phiếu doanh nghiệp)

Script: [`bond/tools/ghep_cap_deal_fibond.py`](../bond/tools/ghep_cap_deal_fibond.py)
Dữ liệu: `bond/source/List_deal_ps_FIBond_2026.08.24.xlsx` (805 dòng)

Schema khác hẳn file TPCP — ba bẫy: `QUANTITY` vô dụng (682/805 dòng = 1, khối
lượng thật ở `FACE_AMOUNT`), `ACCRUED` không phải tiền (chỉ nhận 0–6, là SỐ
NGÀY), `PRICE` không tái tạo được `GROSS_AMOUNT` (luôn lấy thẳng cột này).

**Điều kiện ghép = repo/reverse repo:** cùng mã giấy tờ + cùng đối tác + ngược
chiều + cùng `FACE_AMOUNT`. **Không loại trừ theo `CAPTURE_DATE`** — xác nhận
của Jak (08/2026): BO có thể nhập hai chân cùng một hợp đồng cách xa ngày
nhau, không có nghĩa là không phải repo. `CAPTURE_DATE` lệch xa chỉ tách
riêng một sheet để soát BO, không loại khỏi kết quả.

Bug đã sửa: lượt ghép phần dư từng sắp theo `|lệch mã deal|` — tiêu chí chỉ
đúng cho lượt ưu tiên (cùng ngày nhập). Áp nhầm sang phần dư đẩy một chân về
cuối hàng đợi, ghép với đối tượng cách xa ngày giả tạo (ví dụ
VPBCD080427/TCBV-HO: đáng lẽ cách 14 ngày, bị đẩy thành 84 ngày). Đã sửa:
lượt phần dư ghép theo **thứ tự ngày thanh toán** (FIFO), không theo mã deal.

Kết quả kỳ 24/08/2026: 330 cặp Repo/Reverse Repo (54 cùng ngày nhập, 276 khác
ngày nhập), 133 chân không ghép được.

### Phân loại 4 nhóm (chốt 09/2026, thay nhị phân đi vay/cho vay cũ)

Theo (chân nào thanh toán TRƯỚC) × (Gross bên nào LỚN HƠN). `pnl = Gross(S) −
Gross(B)` luôn luôn, bất kể chân nào trước — hai nhánh `dirn` rút gọn về cùng
biểu thức này (xem `emit()` trong script).

| Chân trước | So sánh Gross | Loai |
|---|---|---|
| Buy trước, Sell sau | Gross(S) > Gross(B) | Cho vay tien |
| Buy trước, Sell sau | Gross(S) < Gross(B) | Vay bond |
| Sell trước, Buy sau | Gross(S) > Gross(B) | Di vay tien |
| Sell trước, Buy sau | Gross(S) < Gross(B) | Cho vay bond |

Đây không phải "lãi/lỗ của đi vay/cho vay tiền" — **"Cho vay bond"** (Sell
trước, trả nhiều hơn khi mua lại) là MSB cho mượn bond, trả rebate interest
trên tiền cọc nhận được: đó là chiều BÌNH THƯỜNG của repo bán trước, không
phải một khoản lỗ bất thường. **"Vay bond"** (Buy trước, bán lại rẻ hơn) là
MSB mượn bond, trả phí ngầm qua phần thu hồi khi bán lại thấp hơn.

Phân bổ thực tế kỳ 24/08/2026 (330 cặp): 270 Cho vay bond, 59 Cho vay tien,
1 Di vay tien, 0 Vay bond.


## Phụ lục — GovBond RPBOD_B002 (thay MSB_RP_DM cũ, 09/2026)

Script ghép cặp: [`bond/tools/ghep_cap_deal_govbond.py`](../bond/tools/ghep_cap_deal_govbond.py)
Tổng hợp D + render HTML: [`bond/tools/build_D_govbond.py`](../bond/tools/build_D_govbond.py),
[`bond/tools/render_govbond_report.py`](../bond/tools/render_govbond_report.py)
Dữ liệu: `bond/source/RPBOD_B002_2026.08.27.xlsx` (601 dòng, CaptureDate 08/06–04/09/2026)
Output: `bond/Phan_tich_GD_Bond_20260608_20260904.html`

Export mới đầy đủ hơn `MSB_RP_DM` cũ nhưng có bẫy: `Cpty_ShortName` không phải lúc nào
cũng là đối tác ngoài thật. `Folders_ShortName` là sổ (book) MSB dùng book chân giao
dịch, và 2 trong 5 sổ AFS-* xuất hiện — `AFS-GOV`, `AFS-ALM` — là sổ nội bộ không tách
bạch được đối tác ngoài (xác nhận của Jak, 09/2026). **Quyết định:** chỉ giữ dòng có
`Folders_ShortName` ∈ {AFS-DCM, AFS-HUONG, AFS-HUY} trước khi ghép cặp (374/601 dòng).

**Thứ tự trước/sau và kỳ hạn dùng `CaptureDate`**, không dùng `SettlementDate` — khác
với `ghep_cap_deal_bond.py` (MSB_RP_DM) và `ghep_cap_deal_fibond.py` (VALUE_DATE).
Xác nhận riêng của Jak cho pipeline này: "logic là lấy capture date, ngày ghi nhận
giao dịch". Hệ quả đã kiểm chứng: **196/198 cặp có CaptureDate hai chân trùng ngày**
(kỳ hạn hiển thị = 0 ngày), trong khi SettlementDate lệch 1–351 ngày — đã báo cho Jak
và được xác nhận giữ nguyên theo CaptureDate, không dùng SettlementDate cho kỳ hạn.

CROSS_OK — một ngoại lệ chéo đối tác đã xác nhận (MSB đứng trung gian, cùng case với
file cũ): deal 50120 (B, Cpty=KBNN) ghép với 50127+50128 (S, Cpty=PGBV-HO), mã
TD2636023, cùng CaptureDate 24/06/2026, cùng khối lượng 6,000,000.

Phân loại 4 nhóm giống hệt FIBond (xem phụ lục trên). Kết quả kỳ 08/06–04/09/2026:
**198 cặp** (104 Cho vay bond, 93 Đi vay tiền, 1 Cho vay tiền, 0 Vay bond), 0 chân dư.

Bug đã sửa trong `build_D_govbond.py::wrate()`: công thức %/năm bình quân gia quyền
gốc (chuyển thể từ `ghep_cap_deal_bond.py::blk()`) cộng dồn pnl trên TẤT CẢ cặp ở tử
số nhưng chỉ tính trọng số cặp có ngày>0 ở mẫu số — với nhóm có rất ít cặp ngày>0 (do
hệ quả CaptureDate ở trên), một cặp lẻ gánh hết pnl của các cặp còn lại, ra %/năm vô
lý (đã thấy -15625%/năm ở nhóm Cho vay bond). Sửa: cả tử số và mẫu số cùng lọc trên
tập cặp có ngày>0.

HTML: 6 trang, cấu trúc khác 2 nhóm A/B cũ — trang "Buy trước" (Cho vay tiền | Vay
bond) và "Sell trước" (Đi vay tiền | Cho vay bond), mỗi trang 2 cột con đầy đủ
(KPI+kỳ hạn+mã TP+folder) đặt cạnh nhau, theo xác nhận của Jak.
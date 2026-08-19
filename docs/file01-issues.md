# File 01 — Nhật ký issue

Ghi các vấn đề phát hiện khi rà soát `01.RptTool_Bond`, kèm bằng chứng số và cách sửa.
Cấu trúc file xem [`file01-rpttool.md`](file01-rpttool.md).

> Bản `.xlsm` trong `bond/source/` là **bản 18/08 trước khi sửa**. Các fix dưới đây Jak áp
> trực tiếp trên máy bank, không commit ngược lại. Số "trước" trong tài liệu này đọc được từ
> file trong repo; số "sau" là kết quả Jak xác nhận.

---

## #1 — PnL trôi mỗi khi điều chuyển deal TB ↔ BB

**Trạng thái**: đã sửa 18/08/2026 · còn phần restate lịch sử

### Bối cảnh nghiệp vụ

Đơn vị kinh doanh từng chuyển bond lỗ từ Trading Book sang Banking Book (coi như giữ đến đáo
hạn). Nay bị yêu cầu bóc tách ngược về TB, làm thủ công, key dần qua mua bán nội bộ với ALM.

Bảng `RP_SEC!CX:DV` (`TB_HTM`) là nơi giữ thủ công: deal **đang nằm ở BB nhưng phải tính như
TB**. Key được deal nào về TB thật thì xoá dòng đó khỏi `TB_HTM`.

### Triệu chứng

Ngày 18/08, `Double check` dòng 46 báo *"Lỗ vượt 71,2 tỷ"* — Daily PnL −92,58 tỷ so VaR₉₅ 21,38.
Dòng 4 (ngưỡng 15 tỷ), 19, 23 cùng đỏ.

### Truy nguyên

`TB_HTM` giảm từ 29 xuống 26 dòng. Ba dòng biến mất:

| Deal ID | Bond | Face |
|---|---|---|
| 30558 | BVBS20126 | 100 tỷ |
| 48295 | TD2434025 | 200 tỷ |
| 48337 | TD2439031 | 400 tỷ |

`SUM(TB_HTM[Amt])` giảm đúng **−700,0 tỷ**. Khớp ba cặp deal `INTERDEAL` trong `Deal ps`
ngày 18/08, đối ứng `AFS-ALM` ↔ `AFS-GOV`:

| Deal | `AD` Book MSB | `AM` Chênh MtM |
|---|---|---|
| 50459 B BVBS20126 | Trading | −14,569 |
| 50458 S BVBS20126 | Banking | +14,569 |
| 50457 B TD2434025 | Trading | −21,234 |
| 50456 S TD2434025 | Banking | +21,234 |
| 50461 B TD2439031 | Trading | −59,608 |
| 50460 S TD2439031 | Banking | +59,608 |

**Nguyên nhân gốc.** `M2` (Daily PnL) = `$M$27 − His53[hôm qua]`, mà

```
M27 = K27 + L27 = SUM(FNRP!Z)/10^9 + SUM(FNRP!AA)/10^9
```

`FNRP!A:AR` chỉ là **MSB - Trading book**. Khoản lỗ của 700 tỷ này hôm qua nằm trong `TB_HTM`
→ chảy vào `M7`, **không** chảy vào `M27`. Hôm nay deal key về TB thật → lỗ hiện trong FNRP →
`M2` đọc như lỗ mới phát sinh trong ngày.

Cùng một khoản lỗ, ghi nhận hai lần ở hai thời điểm, trên hai thước đo.

Hai leg triệt tiêu nhau ở cấp toàn hàng nhưng **rơi vào hai book khác nhau** (`AD` = Trading vs
Banking), mà `M27` chỉ nhìn vế Trading. Ba cặp này **không phải repo** — cột `Z` trống. Các cặp
repo thật (đối tác `CK_KAFI`, `CK_THIENVI`) cộng lại ra **+0,12312**, đúng bằng `Double check!G6`.

### Kiểm chứng

Chuỗi xác nhận khép kín: `M27` file 17/08 = **561,0902**, đúng bằng `Data53_PnL` hôm qua
(1717,1872 − 1156,0970). `M27` 18/08 = 468,5117. Hiệu = **−92,5785** = `M2`.

| Cách đo Daily PnL 18/08 | tỷ VND |
|---|---|
| `M2` thô | −92,5785 |
| `S12` ước lượng PV01 × Δyield | +2,8023 |
| Qua `Deal ps[AM]` của deal `INTERDEAL` | +2,8325 |
| **Cơ sở bất biến `M27 + SUM(TB_HTM[ItD])`** | **+3,9132** |

Chênh ~1,1 tỷ giữa hai nhóm: ba deal `INTERDEAL` chốt ở **yield sổ sách** (2,92 / 2,75 / 2,93)
chứ không phải MtM (4,914 / 4,355 / 4,519), nên hai thước đo không thể trùng. **Chưa kết luận
nhóm nào chuẩn kế toán** — cần biết đợt điều chuyển chốt theo giá sổ hay giá thị trường.

### Cách KHÔNG dùng được

Prorate PnL cả book theo `% amt FNRP` (`=XLOOKUP([@[Bond code]],FNRP!D:D,FNRP!Z:Z)*[@[% amt FNRP]]`)
— thử và loại, ba lý do:

1. **Sai vùng.** `FNRP!D:D`/`Z:Z`/`AA:AA` là Trading book; deal `TB_HTM` nằm ở Banking
   (`AV`/`BR`/`BS`). Cả 26/26 mã tồn tại ở **cả hai** book nên `XLOOKUP` trả số chứ không
   `#N/A` — sai âm thầm.
2. **Tỷ trọng khối lượng không áp được cho PnL.** PnL phụ thuộc giá vốn riêng từng deal. Nhân
   đúng cho Market Value (cột `DI`), sai cho PnL.
3. **Không bất biến.** Delta ra +24,84 (vùng TB) / +603,56 (vùng BB), trong khi mục tiêu ~+2,8…3,9.

`Outstanding_pos_BB_MSB` cũng **không có** cột realised — đúng, vì đây là trạng thái còn mở,
realised = 0 theo định nghĩa. Nguồn per-deal đúng là `ItD adj`, và cột `DD` của `TB_HTM` đã
dùng sẵn.

### Fix đã áp

Theo đúng quy ước có sẵn trong file (`M7` cộng vào TB, `P7` trừ khỏi BB):

```excel
Phan tich_TB!L27  =SUM(FNRP!AA:AA)/10^9+SUM(TB_HTM[ItD])/10^9    ' cong vao TB
Phan tich_BB!P2   =SUM(FNRP!BS:BS)/10^9-SUM(TB_HTM[ItD])/10^9    ' tru khoi BB
```

`K27` và `AU4`/`DS1` giữ nguyên. `SUM(TB_HTM[ItD])` 18/08 = **−497,3855**.

Deal rời `TB_HTM` sang FNRP thì rời vế sau, vào vế trước — tổng không nhảy. **Bất biến với
điều chuyển.** `M2`/`M3`/`M4`/`M5` không phải sửa; `His.TB!L150/L151/L256` và `His.BB!L106`
tự đúng, **không đụng VBA**.

Kết quả Jak xác nhận: Daily PnL = **+3,91**, dòng 46 và dòng 4 về OK.

### Bẫy đã vấp — sửa nửa vời làm vỡ dòng 9

Lần đầu chỉ sửa `L27` mà quên vế BB → 497,4 tỷ bị đếm hai lần trong bộ MSB:

| | MSB TB | MSB BB | Tổng MSB |
|---|---|---|---|
| Trước | 468,5117 | −849,2576 | −380,7459 |
| Chỉ sửa `L27` | −28,8738 | −849,2576 | −878,1314 |
| Sửa cả `P2` | −28,8738 | −351,8721 | −380,7459 |

Tổng SBV (`DS1+DS2` + `AU4+AU5`) = **−380,7459**, không đổi. Nên `Double check` dòng 9
(*YtD PnL của bộ SBV = MSB?*) đỏ ra đúng −497,3855.

**Vế SBV không sai, không được đụng.** Trong phân loại SBV cả hai leg đều là Banking
(`Deal ps!AE`), nên đợt điều chuyển không làm gì bộ SBV.

### Còn treo

- **Restate baseline trên DB53** — chỉ còn **2 mốc**: 31/07 (`M3`) và 04/08 (`M5`).
  **31/12/2025 không cần restate**: `TB_HTM` chưa tồn tại tại ngày đó nên
  `SUM(TB_HTM[ItD])` = 0, cơ sở cũ đã là cơ sở mới. `M4` = −1.161,7875 hiện đã đúng nền.
  Điều này cũng gỡ rủi ro lớn nhất — ngày 31/12/2025 trên 53 chỉ có **328 dòng** so với **349**
  của hai mốc kia, tức là có lệch phiên bản; không đụng vào là an toàn nhất.
- **Dòng 9 xanh không chứng minh YtD đã đúng.** Nó so *tổng*; sai số ở TB và BB ngược dấu nên
  triệt tiêu. Từng vế vẫn lệch tới khi restate xong.
- **Giá điều chuyển nội bộ** chốt theo sổ hay theo thị trường — quyết định khoảng hở 1,1 tỷ.

Bóc theo tenor / loại TP: đã xử lý, xem [#5](#5).

---

## #2 — `Double check!I9` cộng hằng số `+1`

**Trạng thái**: đã biết, Jak sửa dần — `+1` là manual adjust có chủ đích

```excel
I9 = ROUND(G9-H9,0)+1      ' moi dong khac: ROUND(G-H,0)
```

Dòng 9 kiểm *Unrealized ItD MtM PnL nguồn RP_SEC_POS = Basel?*:
`G9` = −3.386,3981 · `H9` = −3.385,7795 · chênh thật **−0,6186 tỷ** → `ROUND` = −1 → `+1` = 0 → "OK".

Dòng 8 dùng chung vế `G` và chênh bằng 0, nên lệch nằm ở phía `BS!G6+L6+Q6`.

**Đề xuất khi gỡ**: đưa ngưỡng ra cột riêng, `E = IF(ABS(G-H)<=K,"OK","CHECK")`, đừng cộng
hằng số vào hiệu. Hiện `ROUND(...,0)` trên đơn vị tỷ đang ngầm chấp nhận ±500 triệu mà không
ghi ở đâu, trong khi dòng 30 ghi rõ ngưỡng ±100 đồng ở cột `J`.

---

## #3 — Thiếu dòng kiểm upload `His.TB`

**Trạng thái**: chưa sửa

`Double check` dòng 37–40 kiểm `His.BB`, `His.TB.SBV`, `His.BB.SBV`, `His.Other` đẩy lên 53 có
lỗi không. **Không có dòng cho `His.TB`** — book lớn nhất là book duy nhất không được kiểm.

File 02 chỉ đọc DB53. Upload hỏng thì File 02 vẫn chạy, vẫn ra báo cáo, chỉ là số cũ.

---

## #4 — `Keo_cong_thuc` sai tên hằng số

**Trạng thái**: chưa sửa

```vb
Application.Calculation = xlCalculationAutomat   ' thieu 'ic'
```

Không có `Option Explicit` nên VBA coi là biến rỗng → gán `0` → **error 1004** khi nhánh
`dieukien > 4` chạy. Đúng phải là `xlCalculationAutomatic`.


---

<a id="5"></a>
## #5 — Bóc theo tenor và loại TP không cộng bằng tổng sau khi sửa #1

**Trạng thái**: đã sửa và đã đẩy lại 53 cho ngày 18/08

`L27` và `P2` nhận phần `TB_HTM` nhưng các bảng bóc chi tiết thì không, nên tổng chi tiết
lệch tổng chung đúng **497,3855 tỷ**. Ba bảng vỡ, không phải một:

| Bảng bóc | Ô | BB có tương ứng? |
|---|---|---|
| TB theo tenor | `L13:L22` | `P12:P21` |
| TB theo loại TP | `L28:L30` | không có |

Chỉ cột **Unrealised** phải chỉnh. `K`/`O` (Realised) giữ nguyên — `TB_HTM` là trạng thái còn
mở, realised = 0 theo định nghĩa.

### Phân bổ `TB_HTM[ItD]` theo tenor (tỷ VND)

| Bucket | Điều chỉnh | `L` mới (TB, cộng) | `P` mới (BB, trừ) |
|---|---|---|---|
| 1Y *(3M+6M+1Y)* | −5,1453 | `L13` −13,8731 | `P12` −19,9590 |
| 2Y | 0 | `L14` 0 | `P13` −22,6806 |
| 3Y | 0 | `L15` 7,5784 | `P14` −12,5550 |
| 4Y | −6,8511 | `L16` −34,0430 | `P15` −21,6536 |
| 5Y | −0,2094 | `L17` −5,5884 | `P16` −3,7943 |
| 7Y | −99,4464 | `L18` −495,0197 | `P17` −199,7578 |
| 10Y | −53,4993 | `L19` −572,5067 | `P18` −163,4337 |
| 15Y | −332,2340 | `L20` −630,0648 | `P19` −1.248,5682 |
| 20Y | 0 | `L21` 0 | `P20` −10,0655 |
| 30Y | 0 | `L22` 0 | `P21` −182,7995 |
| **Tổng** | **−497,3855** | | |

Mọi tenor trong `TB_HTM` (3M · 6M · 4Y · 5Y · 7Y · 10Y · 15Y) đều rơi đúng vào bucket có sẵn.
Dòng dư (`L12`, `P11`) không phải đụng.

Theo loại TP: TPCP **−221,5648** · TPCPBL **−275,8207** · TPCQĐP **0**.

### Công thức

TB tenor `L13` (bucket gộp), rồi `L14` fill xuống `L22`:

```excel
L13  =SUM(SUMIFS(FNRP!$AA:$AA,FNRP!$AI:$AI,{"3M","6M","1Y"}))/10^9
     +SUM(SUMIFS(TB_HTM[ItD],TB_HTM[Tenor],{"3M","6M","1Y"}))/10^9
L14  =SUMIFS(FNRP!$AA:$AA,FNRP!$AI:$AI,'Phan tich_TB'!I14)/10^9
     +SUMIFS(TB_HTM[ItD],TB_HTM[Tenor],'Phan tich_TB'!I14)/10^9
```

TB loại TP `L28` fill xuống `L30`:

```excel
=SUMIFS(FNRP!$AA:$AA,FNRP!$AJ:$AJ,'Phan tich_TB'!I28)/10^9
+SUMIFS(TB_HTM[ItD],TB_HTM[Type],'Phan tich_TB'!I28)/10^9
```

BB tenor `P12` (bucket gộp), rồi `P13` fill xuống `P21` — **dấu trừ**:

```excel
P12  =SUM(SUMIFS(FNRP!BS:BS,FNRP!$CA:$CA,{"3M","6M","1Y"}))/10^9
     -SUM(SUMIFS(TB_HTM[ItD],TB_HTM[Tenor],{"3M","6M","1Y"}))/10^9
P13  =SUMIFS(FNRP!BS:BS,FNRP!$CA:$CA,'Phan tich_BB'!M13)/10^9
     -SUMIFS(TB_HTM[ItD],TB_HTM[Tenor],'Phan tich_BB'!M13)/10^9
```

Upload tự theo: `His.TB!L173:L182` (tenor), `L154/L157/L160` (loại TP), `His.BB!L150:L159`.
Không đụng VBA.

### Phải sửa kèm — bộ kiểm tra sẵn có báo sai

`Phan tich_BB!P23` = `=SUM(FNRP!BS:BS)/10^9-P22` đang bằng 0. Sau khi sửa `P12:P21` thì `P22`
có phần `TB_HTM` còn `SUM(FNRP!BS)` thì không → `P23` nhảy **+497,3855** dù không có gì sai.

Đổi thành so với chính ô tổng, tự bám theo mọi điều chỉnh về sau:

```excel
P23  =P2-P22
O23  =P1-O22
Q23  =(P1+P2)-Q22
```

`Q23` là ô kiểm **cộng gộp** (`=SUM(FNRP!BR:BS)/10^9-Q22`), nằm ngoài khung `M:P` nên rất dễ
sót — Jak phát hiện. Nó vỡ y hệt `P23`. `N23` (coupon) giữ nguyên, `TB_HTM` không đụng coupon.

Hai chỗ trong file **đã theo đúng quy ước này từ trước**, không phải sửa: `T9:T18` (bóc theo
thời gian nắm giữ, đã có `-SUMIFS(TB_HTM[Amt],...)`) và `O26` "ItD Unrealized adj".

Vế TB không có ô đối chiếu tương đương (`L23` vs `L27` không được so ở đâu cả).
**Optimization suggestion**: thêm một ô `=L27-L23` và `=K27-K23`, cùng dạng với `O23`/`P23`.

### Kiểm chứng

| Phải bằng nhau | Giá trị |
|---|---|
| `L23` = `L27` | −1.743,5172 |
| `L28+L29+L30` = `L27` | −1.743,5172 |
| `P22` = `P2` | −1.885,2672 |
| `Q22` | −351,8721 |
| `O23`, `P23`, `Q23` | 0 |

### Điểm mỏng

`I12` / `M11` là dòng dư, tiêu chí `SUMIFS` là **ô trống** — Excel hiểu là `0`, may là cột
`FNRP!AI` / `CA` dùng đúng số `0` cho nhóm dư nên khớp. Gõ nhầm gì vào `I12` là bucket đó lệch
âm thầm, không có cảnh báo.

---

## #6 — Bóc coupon TB lệch tổng 44,25 tỷ (có sẵn từ trước)

**Trạng thái**: chưa sửa · không liên quan #1, phát hiện khi rà #5

| | Giá trị |
|---|---|
| `J23` = `SUM(J12:J22)` | 8.488,8849 |
| `J27` = `SUMIFS(FNRP!P:P,FNRP!J:J,">="&0)/10^9` | 8.533,1349 |
| Chênh | **−44,2500** |

Nguyên nhân: `J27` lọc `FNRP!J:J >= 0` (bỏ trạng thái âm), các dòng tenor `J12:J22` **không
lọc**. Phần `QUANTITY < 0` có coupon đúng bằng **−44,2500**.

Vế BB không dính — `N23` không lọc nên hai bên cùng cơ sở.

Cần Jak quyết: coupon của trạng thái âm nên tính hay không. Bỏ lọc ở `J27` hay thêm lọc vào
`J12:J22` — hai hướng cho hai con số khác nhau, không tự chọn được.


---

## #7 — File 03 `Phan_tich_PnL`: hai ô cross-check vỡ sau khi đổi nền File 01

**Trạng thái**: chưa sửa

File 03 giải thích PnL **theo từng deal**, dựng bottom-up từ `Deal YtD/MtD/DtD`, phân deal
thành 6 nhóm (`OutT0,OutT1` · `OutT0,SellT1` · `BuyT1,OutT1` …). Nó **không biết gì về
`TB_HTM`**, nên toàn bộ số của nó ở **cơ sở cũ**.

`Runtool!J4` và `J9` nối thẳng sang File 01 qua external link `[1]`:

```excel
J4 = ROUND('3.8.PnL Breakdown'!D17 - '[1]Phan tich_TB'!$M$4, 0)
J9 = ROUND('3.8.PnL Breakdown'!F17 - '[1]Phan tich_TB'!$M$2, 0)
```

| | File 03 | File 01 cũ | File 01 mới |
|---|---|---|---|
| YtD `D17` | −664,450 | −664,402 | −1.161,788 |
| MtD `E17` | −237,216 | −236,797 | −734,183 |
| DtD `F17` | −92,578 | −92,578 | +3,913 |

→ `J4` nhảy **+497**, `J9` nhảy **−96**. Đúng bằng `SUM(TB_HTM[ItD])` và `ΔHTM`.

### Cách sửa

**Giữ nguyên phần bóc theo nhóm deal** — nó đang đúng, đo biến động thị trường theo từng deal.
Điều chuyển TB↔BB không phải sự kiện thị trường mà là phân loại lại; ép vào 6 nhóm kia sẽ phá
mất giá trị giải thích của bảng.

Thêm một dòng vào `3.8.PnL Breakdown` — *"3. Điều chỉnh deal điều chuyển TB↔BB"* — rồi đổi
`Total` dòng 17 thành `=D14+D11+D16b`.

| Cột | Công thức | Giá trị |
|---|---|---|
| `D` YtD | `HTM(18/08) − HTM(31/12/2025)` = `−497,3855 − 0` | **−497,3855** |
| `E` MtD | `HTM(18/08) − HTM(31/07)` | *cần số 31/07* |
| `F` DtD | `HTM(18/08) − HTM(17/08)` = `−497,3855 + 593,8772` | **+96,4917** |

Kiểm: `F17` mới = −92,578 + 96,492 = **+3,914** vs `M2` = +3,913 → `J9` = 0.
`D17` mới = −664,450 − 497,386 = **−1.161,836** vs `M4` = −1.161,788 → `J4` = 0.
*(Lệch 0,05 là khoảng hở có sẵn từ trước, không phải do việc này.)*

Nên trỏ qua external link `[1]` sang một ô trên `Phan tich_TB` chứa `SUM(TB_HTM[ItD])/10^9`,
đừng gõ tay — cùng cơ chế mà `J4`/`J9` đang dùng.

---

## #8 — File 02: hằng số cộng/trừ tay trong `Linked (1)`

**Trạng thái**: `−90` là cố ý (Jak xác nhận, sẽ bỏ hôm sau) · còn lại chưa rõ

File 02 đọc `Bond_Trading_Historical` qua SQL ở `Code!D18` rồi `SUMIFS` theo
`"Lo ngay"/"Lo thang"/"Lo nam"`. Nền mới **tự chảy xuống**, không phải sửa công thức.

Nhưng sheet `Linked (1)` — nuôi thẳng mặt báo cáo và file key — có 5 chỗ cộng trừ hằng số:

| Ô | Hằng số | Ghi chú |
|---|---|---|
| `Linked (1)!D36` · `D37` · `D38` | `−90` | **cố ý**, Jak bỏ hôm sau |
| `Linked (1)!F16` | `+5` | chưa rõ nguồn gốc |
| `Linked (1)!G63` | `−100` | chưa rõ nguồn gốc |
| `Linked!T67` | `+0` | vô hại, cùng họ |

Cùng họ với `+1` ở [#2](#2). Mỗi lần đổi nền tính toán thì phải rà lại các hằng số này — nếu
một trong số chúng từng vá cho chính vấn đề mà `TB_HTM` giờ xử lý đúng thì đang đếm hai lần.

---

## #9 — `01.RptTool_FIBond` chưa đọc được

**Trạng thái**: chặn bởi định dạng

File ở dạng `.xlsb` (nhị phân). Tooling hiện có không parse được công thức. Muốn soát thì cần
bản `Save As` sang `.xlsm`.

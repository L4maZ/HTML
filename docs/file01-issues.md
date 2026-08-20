# File 01 — Nhật ký issue

Ghi các vấn đề phát hiện khi rà soát `01.RptTool_Bond`, kèm bằng chứng số và cách sửa.
Cấu trúc file xem [`file01-rpttool.md`](file01-rpttool.md).

> `bond/source/` giữ **bản mới nhất** của 4 file (18/08, sau khi đã áp mọi fix trong tài liệu
> này). Các bản cũ đã xoá khỏi thư mục — vẫn truy được trong lịch sử git nếu cần đối chiếu.
> Số "trước" trong tài liệu đọc từ các bản cũ đó, số "sau" là trạng thái hiện tại.

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

- **Restate baseline trên DB53** — ✅ **31/07 đã xong** (19/08). Xem [#10](#10) về việc
  vì sao không restate 04/08 và các mốc còn lại.
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

**Trạng thái**: ✅ đã sửa (19/08) — `J4` và `J9` về 0. `J7` (MtD) còn lệch, xem [#11](#11).

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

**Trạng thái**: ✅ đã gỡ chặn — Jak đã xuất sang `.xlsm` (20/08). **Chưa rà soát.**

Trước đây file ở dạng `.xlsb` (nhị phân), tooling không parse được công thức. Bản `.xlsm` đã
có trong `bond/source/`, soát được bất cứ lúc nào.


---

<a id="10"></a>
## #10 — Restate baseline 53: làm 31/07, bỏ phần còn lại

**Trạng thái**: 31/07 đã xong · các mốc khác chủ động bỏ

### Quy tắc

Một mốc cần restate khi **và chỉ khi** cả hai điều đúng:

1. Bản ghi ngày đó được đẩy từ file **chưa có fix** `L27`/`P2`
2. `SUM(TB_HTM[ItD])` tại ngày đó **≠ 0**

Từ 18/08 trở đi mọi file đều có fix → mọi bản ghi mới đều nền mới.

### Quyết định từng mốc

| Mốc | Chỉ tiêu | Làm | Lý do |
|---|---|---|---|
| 17/08 | `M2` Lỗ ngày | ✅ đã xong | |
| **31/07** | `M3` Lỗ tháng | ✅ **đã xong 19/08** | Một lần, chữa `M3` cho cả tháng 8 |
| 31/12/2025 | `M4` Lỗ năm | ❌ bỏ | `TB_HTM` chưa tồn tại → `HTM` = 0, cơ sở cũ đã là cơ sở mới |
| 04/08 … 14/08 | `M5` 14D | ❌ bỏ | **Mốc của `M5` chạy theo từng ngày** |

Điểm dễ sai (tôi đã tư vấn hụt lần đầu): restate 04/08 **chỉ** chữa `M5` cho riêng ngày 18/08.
Hôm sau `M5` lấy mốc 05/08, hôm sau nữa 06/08… Muốn `M5` đúng mọi ngày phải restate ~8 ngày.
Trong khi `M3` có mốc `Lastmonth` = 31/07 **cố định suốt tháng 8** — một lần là xong.

`M5` chấp nhận sai tới 01/09 rồi tự lành, có ghi chú. Tỷ lệ công/lợi ích quá xấu.

### Từ 01/09 không phải restate gì nữa

| Chỉ tiêu | Mốc tháng 9 | Nền |
|---|---|---|
| `M2` | ngày làm việc trước | mới |
| `M3` | 31/08 | mới |
| `M4` | 31/12/2025 | `HTM` = 0 |
| `M5` | `RPT−14` ≥ 18/08 | mới |

### Vì sao dùng `UPDATE` chứ không chạy macro

Với **18/08** thì chạy macro (delete-then-insert) là đúng — cả file hôm đó nhất quán.

Với **ngày lịch sử thì không**. Macro đẩy lại toàn bộ ngày đó, gồm ba dòng `Lo ngay`/`Lo thang`/
`Lo nam` = `M2`/`M3`/`M4` **của chính file 31/07** — mà `M2`/`M3` trong đó so với mốc 30/07 và
30/06, hai ngày chưa restate, nên lệch đúng −687,85. File 02 đọc chuỗi `"Lo ngay"/"Lo nam"`
theo ngày để vẽ biểu đồ xu hướng (`Chart data!P2`…), nên đó là ghi số rác vào biểu đồ, rồi lại
phải restate 30/07 và 30/06 — dây chuyền không có điểm dừng.

`M3` chỉ đọc **2 dòng** từ 53 (bộ lọc PQ `Data53_PnL`), và `ItD_PnL_Realised` không đổi vì
`K27` giữ nguyên. Nên phạm vi thật sự là **1 dòng mỗi bảng**.

### Đã chạy

`SUM(TB_HTM[ItD])` tại 31/07 = **−687,8519698095** (28 dòng).

Đối chiếu trước khi sửa — file phải khớp đúng bản ghi trên 53:

| | File 31/07 | DB53 |
|---|---|---|
| `SUM(FNRP!AA:AA)/10^9` | −1014,0872236090 | −1014,087223609 |
| `SUM(FNRP!BS:BS)/10^9` | −2600,3273748588 | −2600,3273748588 |

```sql
UPDATE Bond_Trading_Historical
SET Actual = Actual + (-687.8519698095)
WHERE Rptdate='2026-07-31' AND Indicators='PnL' AND Input_data='Spot'
  AND Class='All' AND Tenor_Value='' AND Rpt_Grp='ItD_PnL_Unrealised';

UPDATE Bond_Banking_Historical
SET Actual = Actual - (-687.8519698095)
WHERE Rptdate='2026-07-31' AND Indicators='PnL' AND Input_data='Spot'
  AND Class='All' AND Tenor_Value='' AND Rpt_Grp='ItD_PnL_Unrealised';
```

Kết quả trên 53:

| Book | Rpt_Grp | Trước | Sau |
|---|---|---|---|
| TB | ItD_PnL_Realised | 1719,3963730977 | *không đổi* |
| TB | ItD_PnL_Unrealised | −1014,087223609 | **−1701,9391934185** |
| BB | ItD_PnL_Realised | 1553,2253933152 | *không đổi* |
| BB | ItD_PnL_Unrealised | −2600,3273748588 | **−1912,4754050493** |

Trong file 18/08, refresh riêng `Data53_PnL` và `His53_BB_MSB` (đừng chạy `Refresh_data`):

| Ô | Trước | Sau |
|---|---|---|
| `Phan tich_TB!M3` | −734,183 | **−46,33** |
| `Phan tich_BB!P4` | 197,844 | **+7,38** |

Kiểm chéo hai đường độc lập, khớp:

```
M27(18/08) − [1719,3964 + (−1701,9392)] = −28,8738 − 17,4572 = −46,331
M3 nền cũ + [HTM(18/08) − HTM(31/07)]   = −236,797 + 190,467 = −46,331
```

### Giá trị `SUM(TB_HTM[ItD])` đã đo được

| Ngày | Số dòng | `SUM(TB_HTM[ItD])` tỷ VND |
|---|---|---|
| 31/12/2025 | *(chưa có bảng)* | 0 |
| 31/07/2026 | 28 | **−687,8519698095** |
| 17/08/2026 | 29 | **−593,8772111732** |
| 18/08/2026 | 26 | **−497,3855320562** |


---

<a id="11"></a>
## #11 — File 03: PnL realized MtD thừa 1.407 tỷ, đang bị vá tay

**Trạng thái**: ✅ **ĐÃ SỬA XONG** (19/08) — `J5`/`J6`/`J7` đều về 0, plug 1407 đã gỡ

Ô `3.8.PnL Breakdown!E19` (trước khi chèn dòng là `E18`) chứa hằng số **1407** gõ tay, bị trừ
khỏi `E16` qua đuôi `-E19` của công thức. Jak xác nhận đây là adjust tay vì số đang sai, và
muốn bỏ hẳn nó.

### Cơ chế — đã chứng minh

Mọi dòng bán bị ép `X=0`:

```excel
X  = IF(F="S", 0, ...)
Y  = IF(X=0, 0, ...)                    ' Goc = 0
AD = IF(F="S", -AB*H, -AB*Y)            ' = qty x gia ban, TOAN BO GROSS
AB = IF(F="S",-G, IF(X=2,V, IF(X=4,G, IF(X=5,V-W, IF(X=6,G-W, 0)))))
```

Giá vốn của lô đã bán phải đến từ dòng đối ứng. Nhưng khi lô đó được **mua lại trong kỳ và còn
giữ tới cuối kỳ**, dòng mua mang `X=3` — mà `X=3` **rơi vào nhánh `0`** của `AB` → `AD = 0`.
Giá vốn không bao giờ vào realized.

Ví dụ `TD2232114` (`Deal MtD` r77–r78):

| | qty | giá | X | Gốc | AD |
|---|---|---|---|---|---|
| Bán | 1.000.000 | 111.212 | 0 | **0** | **+111,212 tỷ** |
| Mua lại | 1.000.000 | 111.117 | 3 | 111.117 | **0** |

Lãi thật ≈ 1.000.000 × (111.212 − 111.117) = **+0,095 tỷ**. File ghi **+111,212 tỷ**.

### Phạm vi — 10 mã

`TD2434022` · `TD2530010` · `TD2434024` · `TD2131012` · `TD2535023` · `TD2434026` ·
`TD2232114` · `TD1934188` · `TD2333118` · `TD2535028`

Đều là bán trong kỳ + mua lại trong kỳ còn giữ tới cuối kỳ (dạng cặp repo).

Cùng gốc với hai ô check đang đỏ mà chưa ai để ý: `Runtool!J5` = −15.050.000 và
**`J6` = 17.702 tỷ** (kiểm position MtD). `Deal DtD` có `SUM(AG)` = 0 nên không dính.

### Hai cách tái lập đã thử và LOẠI

`PPL!F4` ghi phương pháp Risk: *"Tính theo giá vốn bình quân gia quyền của các deal tồn đầu năm
(theo giá mtm đầu năm + giá mua các deal mới trong năm đến thời điểm bán)"*.

**Cách 1 — bình quân gia quyền tĩnh cả kỳ** (tồn đầu kỳ theo giá MtM đầu kỳ + toàn bộ mua trong kỳ):

| | File | Cách 1 | Check hiện tại |
|---|---|---|---|
| YtD | 252,648 | 200,983 | `J4` = 0 ✓ |
| MtD | 1.490,343 | **84,035** | `J7` = 1407 ✗ |
| DtD | 6,944 | 1,436 | `J9` = 0 ✓ |

MtD cần 83,762 → cách 1 lệch **0,273**, nằm trong khoảng hở 0,419 vốn có giữa File 03 và File 01.
Rất sát. **Nhưng nó làm hỏng YtD (−51,7) và DtD (−5,5)** — hai chỗ đang khớp File 01 qua một
đường tính hoàn toàn độc lập. Nên cách 1 không phải phương pháp file đang dùng.

**Cách 2 — bình quân gia quyền động theo ngày giao dịch** (đúng chữ *"đến thời điểm bán"*, deal
mua sau khi bán không vào giá vốn của lần bán đó): số bung ra vô nghĩa ở cả ba cửa sổ
(YtD ra 191 triệu tỷ). Mô hình trình tự deal của tôi sai — có bán khống, deal ngoài cửa sổ,
và lượng tồn về 0 làm bình quân nổ.

### NGUYÊN NHÂN GỐC — `Table.ExpandTableColumn` sai tên cột

Bộ máy `X`/`Gốc`/`AB`/`AH`/`AI` **không sai** — nó chạy đúng ở YtD và DtD. MtD bị **bỏ đói dữ
liệu đầu vào**.

| | Số deal trong `ItD_*` (tồn T0) | Trong `Deal *` | Giao nhau |
|---|---|---|---|
| DtD | 89 | 99 | **88** ✓ |
| **MtD** | 79 | 108 | **0** ✗ |

`Deal MtD` chỉ chứa deal phát sinh trong tháng (ID 50273+), không một deal tồn đầu kỳ nào
(ID 41818, 42756…). Vì thế `V` (OutT0) luôn = 0 ở **0/26 mã**, và không sell nào có giá vốn.

```m
' Bond_Deals MtD
#"Merged Queries"   = Table.NestedJoin(..., {"DEAL_ID"}, #"ItD MtD", {"DEAL_BUY_ID"}, ...)   ' join DUNG
#"Expanded ItD MtD" = Table.ExpandTableColumn(..., "ItD MtD", {"BondsDeals_Id"}, ...)        ' expand SAI
```

`ItD MtD` chỉ có `DEAL_BUY_ID` và `Amt` — **không có** `BondsDeals_Id`.
`Table.ExpandTableColumn` gặp cột không tồn tại thì **không báo lỗi**, nó tạo cột toàn `null`.
Nên `([ItD MtD.BondsDeals_Id] <> null)` luôn FALSE, điều kiện lọc rút về còn
`[CAPTURE_DATE] > Lastmonth`, và toàn bộ deal tồn đầu kỳ bị loại.

`DtD`/`YtD` không dính vì nguồn của chúng đặt tên cột là `BondsDeals_Id`, khớp lệnh expand.
Join của MtD đã được sửa sang `DEAL_BUY_ID` nhưng expand thì quên.

### Bản sửa

Đổi 3 chỗ `BondsDeals_Id` → `DEAL_BUY_ID` trong `Bond_Deals MtD` (expand, điều kiện lọc,
remove column). Không đụng gì trên sheet.

### Kiểm

| Kiểm | Trước | Sau |
|---|---|---|
| `Deal MtD` có dòng `V ≠ 0` | 0/26 mã | phải có |
| `Runtool!J5` | −15.050.000 | 0 |
| `Runtool!J6` position MtD | **17.702 tỷ** | 0 |
| `Runtool!J7` | 1407 | 0 |
| `3.8!E14` | 1.490,343 | **100,8** |
| `3.8!E11` | −320,559 | **−337,6** |
| `3.8` Total MtD | 1.360,250 | **−46,3** khớp `M3` = −46,331 |

`J6` = 17.702 tỷ là ô đã báo động về việc này từ đầu, nằm ngay trong `Runtool`, chưa ai xử lý.

Lưu ý khi đọc lại: tôi từng dự đoán `E14` ≈ 83,8. Thực tế ra 100,8 vì việc đưa lại các dòng
tồn đầu kỳ làm dịch **cả hai** vế realized và unrealized, không chỉ realized. Con số 83,8 đến
từ bản tái lập tĩnh đã bị loại ở trên, đừng dùng nó làm mốc.

### Bài học

Tôi mất bốn lượt đi tái lập phương pháp tính giá vốn (bình quân tĩnh, bình quân động theo
TRADING_DATE, theo CAPTURE_DATE — xem phần dưới) trong khi vấn đề là dữ liệu không tới. Ba
sheet YtD/MtD/DtD **cùng một cấu trúc** và hai trong ba đang đúng — lẽ ra phải so chúng với
nhau trước tiên, chỗ khác nhau mới là chỗ cần nhìn.

Bẫy chung cần nhớ: **`Table.ExpandTableColumn` với tên cột không tồn tại trả về `null` chứ
không báo lỗi.** Mọi điều kiện `<> null` dựa trên nó sẽ âm thầm thành FALSE.

### Liên quan

`Runtool` r11 Log của chính file: *"Nếu có deal bị xóa thì số DtD bị lệch, nếu xóa vắt tháng
hoặc vắt năm thì số YtD và MtD bị lệch => Pending"*. Cùng vùng vấn đề nhưng **khác cơ chế** —
mục này là phân loại `X=3`, không phải deal bị xoá.


---

## #12 — Hai ô check YtD của File 03 chưa bao giờ bằng 0

**Trạng thái**: chưa sửa · có sẵn từ trước, không do đợt sửa 19/08

| Ô | Nội dung | Giá trị |
|---|---|---|
| `Runtool!J2` | YtD — Amt realized tổng mua = tổng bán | **19.117.854** |
| `Runtool!J3` | YtD — Check position cuối kỳ datamart vs FNRP | **−3.823.570.800** |

Cùng họ với `J5`/`J6` đã chữa ở [#11](#11), chỉ khác cửa sổ và nhẹ hơn nhiều (3,8 tỷ so với
17.702 tỷ). Đáng nghi là cũng lệch khớp dữ liệu bên `Bond_Deals YtD`.

Đáng chú ý: `J4` (YtD PnL vs File 02) **đang bằng 0**, tức là tổng PnL YtD vẫn khớp dù hai ô
kiểm dữ liệu đầu vào thì không. Nên đây không phải lỗi lớn như MtD, nhưng cũng chưa được giải
thích. Cần soi `Bond_Deals YtD` theo cùng cách đã làm với MtD:

```
so sánh tập DEAL_ID giữa ItD_YtD ('Amt Outs'!F:G) và Deal YtD
```

MtD trước khi sửa có giao nhau = 0. Nếu YtD ra một con số nhỏ hơn kỳ vọng thì cùng bản chất.


---

## #13 — VaR bộ MSB Trading upload số của hôm trước

**Trạng thái**: ✅ đã sửa (20/08)

`His.TB!L20:L25` — cả 6 giá trị VaR **gõ tay**, giống hệt nhau ở cả ba bản file 17/08 và 18/08,
và khớp đúng bản ghi **17/08** trên DB53. Tức là file 18/08 sắp đẩy VaR của 17/08 lên làm VaR
của 18/08.

Nguồn số đúng **đã có sẵn trong file** (bảng `VaR`, query `VaR Bond MHCC`), đúng ngày 18/08:

| Chỉ tiêu | Gõ tay (= 17/08) | Đúng (18/08) | Lệch |
|---|---|---|---|
| `VaR_95` | 21,3831 | 24,2188 | −2,84 |
| `VaR_99` | 44,5061 | 51,5956 | −7,09 |
| `VaR10D_95` | 119,5598 | 136,1272 | −16,57 |
| `VaR10D_99` | 160,1864 | 181,8744 | −21,69 |
| `VaR20D_95` | 172,9808 | 196,8365 | −23,86 |
| `VaR20D_99` | 214,8738 | 246,2924 | −31,42 |

**Thấp hơn thực tế 13–15%.** VaR là số hạn mức, và `Double check` dòng 47 backtest PnL dựa
trên chuỗi này.

Bằng chứng đây là lỗi chứ không phải chủ ý: **ba book còn lại** (`His.BB`, `His.TB.SBV`,
`His.BB.SBV`) đều đọc bằng `SUMIFS` từ cùng bảng `VaR` và khớp 100%. Chỉ `His.TB` gõ tay.

Bản sửa — theo đúng pattern đang chạy ở 3 sheet kia:

```excel
L20  =SUMIFS(VaR!K:K,VaR!J:J,"VaR99%",VaR!H:H,"VaR20D",VaR!G:G,"TD",VaR!F:F,"MSB")/10^9
L21  =SUMIFS(VaR!K:K,VaR!J:J,"VaR95%",VaR!H:H,"VaR20D",VaR!G:G,"TD",VaR!F:F,"MSB")/10^9
L22  =SUMIFS(VaR!K:K,VaR!J:J,"VaR99%",VaR!H:H,"VaR10D",VaR!G:G,"TD",VaR!F:F,"MSB")/10^9
L23  =SUMIFS(VaR!K:K,VaR!J:J,"VaR95%",VaR!H:H,"VaR10D",VaR!G:G,"TD",VaR!F:F,"MSB")/10^9
L24  =SUMIFS(VaR!K:K,VaR!J:J,"VaR99%",VaR!H:H,"VaR1D", VaR!G:G,"TD",VaR!F:F,"MSB")/10^9
L25  =SUMIFS(VaR!K:K,VaR!J:J,"VaR95%",VaR!H:H,"VaR1D", VaR!G:G,"TD",VaR!F:F,"MSB")/10^9
```

### Còn treo — kiểm chuỗi lịch sử

Chưa chạy. Cần biết việc copy-forward đã kéo dài bao lâu:

```sql
SELECT Rptdate, Indicators, Actual
FROM Bond_Trading_Historical
WHERE Indicators IN ('VaR_95','VaR_99') AND Class='All'
ORDER BY Rptdate DESC;
```

Các ngày liên tiếp lặp đúng cùng giá trị = chuỗi VaR bộ MSB Trading trên 53 bị trôi.

---

## #14 — CVaR: bốn book dùng bộ kịch bản khác nhau

**Trạng thái**: ✅ đã thống nhất (20/08)

Trước khi sửa, cùng một chỉ tiêu nhưng mỗi book một bộ kịch bản, và `His.TB` còn tự mâu thuẫn
giữa các horizon:

| Sheet | `CVaR_95` | `CVaR_99` | `CVaR10D_95` | `CVaR10D_99` |
|---|---|---|---|---|
| `His.BB.SBV` | gõ cứng | gõ cứng | {1..12} | {1,2} |
| `His.TB.SBV` | {1..12} | {1,2} | {1..12} | {1,2} |
| `His.BB` | {1..12} | {1,2} | {1..12} | {1,2} |
| `His.TB` | {1..12} | {1,2} | **{1..13}** | **{1,2,3}** |

Hệ quả: CVaR của bốn book không so sánh được với nhau và không cộng lại được.

**Quy ước đã chốt** — áp cho **cả 4 book, cả 3 horizon**:

| | Bộ kịch bản |
|---|---|
| 95% | `{1,2,3,4,5,6,7,8,9,10,11,12,13}` |
| 99% | `{1,2,3}` |

```excel
=-AVERAGE(AVERAGEIFS(VaR!K:K,VaR!J:J,{BỘ},VaR!H:H,"HORIZON",VaR!G:G,"BOOK",VaR!F:F,"BỘ"))/10^9
```

Lưu ý khi đọc số cũ: `CVaR_95`, `CVaR_99`, `CVaR20D_99` của cả 4 book đổi giá trị do đổi quy
ước, không phải do số liệu (ví dụ `His.TB` CVaR_95: 40,3823 → 39,1883). Báo cáo sẽ thấy một
bậc nhảy tại ngày áp dụng.

Còn một khác biệt chưa xử lý: `His.TB` có dòng **`CVaR10D_97.5`** (`{1..6}`) mà ba book kia
không có. Thêm hay bỏ là quyết định về bộ chỉ tiêu, không phải sửa lỗi.

### Ghi chú về 14 kịch bản

Bảng `VaR` có 14 kịch bản, sắp xếp lỗ nặng nhất trước (`KB=1`). Lấy trung bình 13/14 kịch bản
cho mức "95%" không phải đuôi 5% theo nghĩa thông thường. Đây là quy ước của phòng, ghi lại để
người đọc sau không hiểu nhầm là lỗi.


---

## #15 — FIBond: mã MFKR bị loại khỏi mọi query, phải bù bằng hardcode

**Trạng thái**: `Tygia` · `RPBOD_B003` · `Upload53_RPBOD03` ✅ đã sửa ·
`Portfolio rating` ⏸ **code sẵn sàng, chưa áp** · market value còn treo

### Triệu chứng

`Portfolio rating` (FIBond, sheet `Rating`) có 2 step cuối chèn tay một dòng:

```m
ITB_Amt_Value = Excel.CurrentWorkbook(){[Name="ITB_Amt"]}[Content]{0}[Column1],
#"Added MFKR" = Table.InsertRows(#"Renamed Columns", Table.RowCount(#"Renamed Columns"),
    {[Issuer = "MFKR", #"Total Amount" = ITB_Amt_Value, Rating = "B", ...]})
```

Cả số tiền lẫn rating `"B"` đều gõ tay.

### Nguyên nhân gốc

MFKR = 4 deal, folder **`AFS-ITB`**, đồng tiền **`KGS`**, trái phiếu chính phủ **Kyrgyz**
(`GBA05310727`, `GBA05310713`). Cả 4 mang `TypeOfInstr_ShortName = 'GOVBOND'`.

| Bước lọc trong `RPBOD_B003` | Còn lại | MFKR |
|---|---|---|
| Ban đầu | 840 | ✓ |
| `Issuer <> "MSB-BANK"` | 348 | ✓ |
| **bỏ `ECON_UNL` / `GOVBOND`** | **80** | **✗** |

Filter sinh ra để loại TPCP Việt Nam (thuộc File 01) nhưng quét luôn trái phiếu chính phủ
**nước ngoài**. Nên MFKR biến mất khỏi mọi query, và ai đó bù lại bằng hardcode.

Tiêu chí phân biệt: `Folder = "AFS-ITB"` và `Currency <> "VND"` cho **kết quả giống hệt** —
đúng 4 dòng MFKR, không dính issuer nào khác. Chọn `Folder` vì nó nói về *book nào thuộc tool
nào*, đúng bản chất filter.

### Tỷ giá — khớp tuyệt đối

`Tygia` (`Ref!AA1:AD2`): `KGS` = **299,3102** (SOM KYRGYZSTAN).

| | KGS | × 299,3102 |
|---|---|---|
| `FaceAmount` | 1.911.785.400 | **572,2169 tỷ** |
| `GrossAmount` | 1.310.911.248,78 | 392,3691 tỷ |
| `Accr` | 8.935.572,59 | 2,6745 tỷ |

**572,2169 trùng khít** `ITB_Amt` (`Summary!DB1` = `SUM(FNRP_ITB[Nominal_VND])/10^9`) và dòng
gõ tay `Upload53!r173`. Xác nhận đây đúng là nguồn của số hardcode.

### Đã sửa

**`Tygia`** — bỏ step cuối `Table.SelectRows(..., each [Indicators] = "KGS")` để trả về mọi
đồng tiền. An toàn: 3 chỗ dùng `Tygia` (`FNRP!EH2/EH3` XLOOKUP theo `Name`, `BS_KGS` và
`offBS_KGS` SelectRows theo `Indicators`) đều chọn theo khoá, không cái nào giả định 1 dòng.

**`RPBOD_B003`** và **`Upload53_RPBOD03`** — 3 thay đổi:
1. `Removed Columns`: **giữ lại** `"Currencies_ShortName"`
2. `Filtered Rows`: thêm `or [Folder] = "AFS-ITB"`
3. Thêm 4 step quy đổi ngoại tệ theo từng đồng tiền

```m
BangTyGia = Excel.CurrentWorkbook(){[Name="Tygia"]}[Content],
TyGiaCua = (ccy as text) as number =>
    let hit = Table.SelectRows(BangTyGia, (t) => t[Indicators] = ccy)
    in  if Table.IsEmpty(hit) then error "Tygia: thieu ty gia cho " & ccy else hit{0}[Average],
ColsInt = {"FaceValue", "Face_Amount", "Gross_Amount"},
ColsNum = {"Discount", "Premium", "Discount_Remain", "Premium_Remain", "ValueDate_Accred", "Accr"},
#"Quy doi ngoai te" = Table.FromRecords(
    Table.TransformRows(#"Filtered Rows1", (row) =>
        if row[Currencies_ShortName] = "VND" then row
        else let fx = TyGiaCua(row[Currencies_ShortName]) in
            Record.TransformFields(row,
                List.Transform(ColsInt, (c) => {c, (v) => if v = null then null else Number.Round(v * fx, 0)})
                & List.Transform(ColsNum, (c) => {c, (v) => if v = null then null else v * fx})
            )
    ),
    Table.ColumnNames(#"Filtered Rows1")
)
```

Không còn chuỗi `"MFKR"` hay `"KGS"` nào trong code. Thêm/bớt deal, issuer mới, đồng tiền mới
đều tự chạy; thiếu tỷ giá thì **báo lỗi rõ** thay vì âm thầm bỏ qua quy đổi.

`Quantity` không quy đổi (là số lượng). `FaceValue` có quy đổi để `Quantity × FaceValue =
Face_Amount` vẫn đúng.

### Còn treo

**1. `Portfolio rating` chưa áp.** Nó **không đọc từ `RPBOD_B003`** — đọc thẳng file raw và có
bản sao riêng của cùng cái filter. Cần 3 thay đổi: khai báo `BangTyGia`/`TyGiaCua` sau
`Promoted Headers`, thêm `or [Folder] = "AFS-ITB"` vào filter, và nhân tỷ giá trong cột `Amount`:

```m
fx = if [Currencies_ShortName] = "VND" then 1 else TyGiaCua([Currencies_ShortName]),
final = resultNumber / 1000000000 * fx
```

Kỳ vọng sau khi áp: bảng `Rating` 20 → **21 dòng**, tổng về lại **31.687,83** (đúng bằng số
thời còn hardcode, nhưng từ dữ liệu thật), rating MFKR lấy từ `FIBond_Rating` trên 53.

**2. Market value MFKR chưa dựng lại được.** Dòng gõ tay `Upload53!r193` = **405,688 tỷ**.
Từ file raw chỉ ra được `GrossAmount` × tỷ giá = 392,3691 hoặc `Gross+Accr` = 395,0436.
Không khớp. Market value đến từ một nguồn định giá khác, chưa xác định.

**3. `Upload53` là sheet gõ tay.** VBA `b_UploadDB` chỉ **đọc** nó, không điền. Nên dòng
`r173` (Face 572,217) và `r193` (Market 405,688) phải xoá bằng tay khi đã có nguồn thật.

### Lỗi trình tự của tôi

Tôi hướng dẫn gỡ hardcode ở `Portfolio rating` **trước khi** tìm ra nguyên nhân gốc. Kết quả:
MFKR bị lọc mất mà không còn gì bù vào, bảng `Rating` mất hẳn mã này. Đáng lẽ phải xác định
nguồn thay thế trước, gỡ hardcode sau.

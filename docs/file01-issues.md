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

- **Restate baseline trên DB53** cho `M3`/`M4`/`M5` (và `P3`/`P4`/`P5`). Mốc 17/08 đã xong.
  Còn 04/08 (14D), 31/07 (tháng), 31/12/2025 (năm) — cần `SUM(TB_HTM[ItD])` tại từng ngày,
  lấy từ file lưu theo ngày.
- **Dòng 9 xanh không chứng minh YtD đã đúng.** Nó so *tổng*; sai số ở TB và BB ngược dấu nên
  triệt tiêu. Từng vế vẫn lệch tới khi restate xong.
- **Bóc theo tenor không cộng bằng tổng.** `His.TB!L172/L183` lấy `K12`/`L12` theo tenor, chưa
  có phần `TB_HTM`. Muốn khớp phải phân bổ `TB_HTM[ItD]` theo `TB_HTM[Tenor]` (cột `DN`).
- **Giá điều chuyển nội bộ** chốt theo sổ hay theo thị trường — quyết định khoảng hở 1,1 tỷ.

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

# File 01 — `01.RptTool_Bond` — đối chiếu với file thật

Đọc trực tiếp bản `01.RptTool_Bond_2026.08.18.xlsm` (`bond/source/`), ngày chốt `RPT = 18/08/2026`.
Trang này **thay thế** phần cấu trúc trong `File01_Context_Export.md` — bản export đó lấy từ memory
của các phiên trước, nhiều chỗ đã lệch so với file hiện tại.

Nguồn số liệu ở đây là XML thô trong file (`xl/tables/`, `xl/workbook.xml`, `customXml/item1.xml`
giải nén ra `Formulas/Section1.m`, `vbaProject.bin`) — không phải suy đoán.

---

## 1. Cấu trúc thật vs memory

| Hạng mục | Memory nói | File 18/08 | |
|---|---|---|---|
| Sheet | 35 | **32** (2 hidden: `Cai tien`, `Rule & Gap`; +1 hidden `Thay doi_BB` → 3 hidden) | lệch |
| Table | 51 | **56** | lệch |
| Power Query | 47 | **50** | lệch |
| External link | — | **27** | memory không có |
| Param trong `Input` | 41 | **41** (`A1:C42`) | khớp |

Đừng dùng lại 3 con số 35/51/47 nữa.

### 32 sheet, theo đúng thứ tự tab

`Log` · `Data flows` · `Cai tien`⁽ᵃ⁾ · `Rule & Gap`⁽ᵃ⁾ · `Validate` · `Ref` · `Dictionary` ·
`Runtool` · `Double check` · `Phan tich_TB` · `Phan tich_BB` · `Thay doi_BB`⁽ᵃ⁾ · `Giam sat` ·
`BS` · `VaR senerio` · `Yield 1Y` · `Input>>` · `Market` · `Deal ps` · `FNRP` · `RP_SEC` ·
`His53` · `VaR` · `Funding` · `Output>>` · `His.DataTT` · `His.TB` · `His.BB` · `His.TB.SBV` ·
`His.BB.SBV` · `His.Other` · `No PT`

⁽ᵃ⁾ hidden.

`Input>>` và `Output>>` là sheet phân cách, không chứa dữ liệu — mọi thứ trước `Input>>` là
tham số/kiểm tra, giữa hai sheet đó là dữ liệu nạp vào, sau `Output>>` là dữ liệu đẩy lên DB53.

### 50 query Power Query

Nhóm theo nguồn:

- **CSV Murex/FNRP (Param → `Csv.Document`)**: `Tradingbook_MSB`, `Bankingbook_MSB`,
  `Tradingbook_SBV`, `Bankingbook_SBV`, `FNRP_Future`, `FNRP_Repo`, `RPBOD3`
- **Kondor xlsx**: `Outstanding pos_TB_MSB`, `Outstanding pos_BB_MSB`, `Outstanding pos_TB_SBV`,
  `Outstanding pos_BB_SBV`, `Deal_delete`, `Deal_fix`, `Mapping deal`, `Mapping sell Unrealized`,
  `Newdeals`, `Buy price`, `Sell unvalid`, `Buy unvalid`
- **DB53 / lịch sử**: `Data53_PnL`, `Data53_PV01`, `His53_BB_MSB`, `His53_TB_SBV`,
  `His53_BB_SBV`, `His pos`, `His pos BB`, `His NPT`, `His yield bq`, `His funding cost`
- **Thị trường**: `Yield`, `Yield_yesterday`, `Yield_lastmonth`, `Yield_lastyear`, `Yield_1Y`,
  `Outright`, `Repo`, `Repo TD`, `Aution`, `OMO`, `MM`, `Bond Last Tradingdate`
- **Nội bộ / tính toán**: `Param`, `VaR_series`, `VaR_yesterday`, `VaR Bond MHCC`,
  `On_off_ngoai bs`, `PTCK on BS`, `PTCK offBS`, `Est NIM`, `Giam sat yesterday`

`Param` đúng như memory mô tả:

```m
(TableName, RowNumber) =>
    let Source = Excel.CurrentWorkbook(){[Name=TableName]}[Content],
        value  = Source{RowNumber-1}[Values]
    in  value
```

Mọi query CSV đều gọi `Param("Input", n)` → đường dẫn dựng từ `RPT`. Đổi ngày báo cáo ở
`Runtool!C2` là toàn bộ nguồn đổi theo. Đây là điểm vào duy nhất.

---

## 2. `Runtool` — 12 ô KHÔNG LINK (danh sách đã đổi)

Vẫn đúng **12 ô**, nhưng **không còn là bộ ô mà memory ghi**. Memory ghi
`C17,C18,C19,C20,C22,C24,C25,C26,C39,C40,C41,C42`. File 18/08:

| Ô | Nhãn (cột B) | Nội dung | Nhận xét |
|---|---|---|---|
| `C2` | Report date | `18/08/2026` | **cố ý** — đây là ô input duy nhất, `RPT` |
| `C12` | Tool Giám sát deal | rỗng | chưa dùng |
| `C17` | *(trống)* | rỗng | dòng ngăn |
| `C18` | Latest update QLHS | `30/06/2026` | ngày gõ tay, **đã cũ 1,5 tháng so với RPT** |
| `C19` | List deals | UNC cố định `…/List deal ps 2025.xlsx` | năm gõ cứng |
| `C20` | *(trống)* | rỗng | dòng ngăn |
| `C24` | HLA level 2 FI Bond | rỗng | chưa dùng |
| `C25` | List deals thứ cấp | UNC `HNX_Outright_Repo.xlsm` | không theo `RPT` |
| `C26` | List deals sơ cấp | UNC `HNX_DS NY.xlsm` | không theo `RPT` |
| `C31` | Deal ps 2016-2024 | UNC cố định | lưu trữ, đúng là cố định |
| `C32` | Deal ps | UNC `List deal ps 2025.xlsx` | **trùng nội dung `C19`** |
| `C41` | Tool cap | rỗng | macro `ToolCap.Open_ToolCap` đọc ô này |

Ba ô đáng xử lý, không phải cả 12:

1. **`C19` và `C32` trùng nhau** — cùng trỏ `List deal ps 2025.xlsx`. VBA `Copy_to_listdeal` đọc
   `C19`, các chỗ khác đọc `C32`. Sang 2027 mà chỉ sửa một ô là hai nhánh chạy lệch nhau, im lặng.
   Nên để một ô là nguồn, ô kia `="…List deal ps "&YEAR(RPT)&".xlsx"`.
2. **`C18` = 30/06/2026** — mốc QLHS gõ tay, ai quên cập nhật thì khối NIM/QLHS tính theo mốc cũ
   mà không có cảnh báo nào.
3. **`C41` rỗng** — `Open_ToolCap` sẽ `Workbooks.Open("")` → lỗi runtime khi bấm nút.

`C17`, `C20`, `C12`, `C24` là chỗ trống có chủ đích (dòng ngăn / chỉ tiêu chưa triển khai) —
không cần link.

### Defined name hỏng

`BQ = 'His53'!#REF!` và `BV = 'His53'!#REF!` — hai name đã gãy tham chiếu, còn sót lại.
`_xleta.T = #NAME?` cũng vậy. Chưa công thức nào đang gọi, nhưng nên xoá cho sạch.

---

## 3. Sheet `Double check` — 51 kiểm tra, đang có 8 cờ CHECK

Đây là cửa kiểm số của file, không phải `Validate` (`Validate` là bản ghi phương pháp luận
FRP, 21 mục — tài liệu, không tính toán).

Công thức chung: `E = IF(I=0,"OK","CHECK")`, với `I = ROUND(G-H, 0)`.

### Lỗi phải sửa — `I9` có `+1` gắn cứng

```
I9 = ROUND(G9-H9,0)+1
```

Mọi dòng khác dùng `ROUND(G-H,0)` trơn. Riêng dòng 9 cộng thêm 1.

Dòng 9 là *"Unrealized ItD MtM PnL của nguồn RP_SEC_POS = Basel?"*:

| | Giá trị (tỷ VND) |
|---|---|
| `G9` (Phan tich_TB!M7 + Phan tich_BB!P7) | −3.386,3980955 |
| `H9` (BS!G6 + BS!L6 + BS!Q6) | −3.385,7795090 |
| Chênh thật | **−0,6185865 tỷ ≈ −618,6 triệu** |
| `ROUND(chênh, 0)` | −1 |
| `+1` | **0 → hiện "OK"** |

Chênh 618,6 triệu giữa nguồn RP_SEC_POS và Basel đang bị `+1` che, báo OK. Bỏ `+1` thì dòng
này ra CHECK. Đây không phải sai số làm tròn — dòng 8 (SBV vs MSB, cùng công thức `G`) ra chênh
đúng bằng 0, nên `G` không có vấn đề; lệch nằm ở phía `BS!G6+L6+Q6`.

**Đề xuất**: sửa `I9` về `=ROUND(G9-H9,0)`, rồi truy phần chênh ở `BS!G6/L6/Q6`. Nếu thực sự
có sai số nguồn chấp nhận được thì khai báo ngưỡng ra một ô riêng
(`=IF(ABS(G9-H9)<=Ngưỡng,"OK","CHECK")`), đừng cộng hằng số vào hiệu.

### Ngưỡng làm tròn không nhất quán

| Dòng | Công thức `I` | Dung sai ngầm |
|---|---|---|
| 8, 9, 10 | `ROUND(G-H,0)` | ±500 triệu VND (đơn vị tỷ) |
| 16–20, 24 | `G-H` | 0 tuyệt đối |
| 30 | `ROUND(G-H,6)` | ±100 đồng (có ghi chú ở `J30`) |

Ba mức dung sai khác nhau cho cùng một loại kiểm tra vị thế/PnL. Dòng 30 ghi rõ ngưỡng ở cột
`J`, dòng 8–10 thì không — người đọc không biết mình đang chấp nhận nửa tỷ đồng.

**Đề xuất**: đưa ngưỡng ra một cột riêng (`K`), `E = IF(ABS(G-H)<=K,"OK","CHECK")`, để dung sai
là dữ liệu đọc được chứ không nằm trong tham số của `ROUND`.

### 8 dòng đang CHECK ở bản 18/08

| # | Nhóm | Nội dung | Số |
|---|---|---|---|
| 4 | PnL | PnL bất thường (ngưỡng 15 tỷ) | −92,58 tỷ |
| 10 | Position | Deal phát sinh chênh MtM > 40bps | có deal gắn cờ ở `Deal ps!AK` |
| 19 | Position | Trạng thái đầu kỳ vs cuối kỳ — Trading | 17.513 vs 18.213, lệch **700** |
| 23 | Position | Trạng thái đầu kỳ vs cuối kỳ — Banking | 10.858,31 vs 10.158,31, lệch **700** |
| 28 | Position | Check FNRP, Basel, RPBOD3, GL | `BS!G19 ≠ 0` |
| 42 | Time | Deal TB bộ SBV giữ > 345 ngày | 2 mã |
| 43 | Time | Deal TB bộ MSB giữ > 345 ngày | 8 mã |
| 46 | Sensitivity | Daily PnL vượt VaR(RPT−1) | −92,58 vs −21,38 → vượt 71,2 tỷ |

Dòng 19 và 23 lệch **cùng một lượng 700, ngược dấu nhau** (Trading +700, Banking −700). Đó là
một lô 700 chuyển book TB→BB chưa được phản ánh ở số đầu kỳ, không phải hai lỗi riêng biệt.

Dòng 4 và 46 cùng bắt một sự kiện: khoản lỗ −92,58 tỷ trong ngày, vượt cả ngưỡng 15 tỷ lẫn
VaR hôm trước. Ngưỡng 15 tỷ ở dòng 4 gõ thẳng trong công thức
(`IF(ABS('Phan tich_TB'!M2)>15,…)`) — nên đưa về `Ref` như ngưỡng 40bps đã làm.

---

## 4. Những chỗ memory ghi sai / đã thay đổi

### `Ref!BF:BG` — danh sách override hiện **rỗng**

Memory: *"col S = adjusted, override 7 deal cụ thể qua `Ref!BF:BG`"*.

File 18/08: `Ref!BF1:BG1` chỉ còn header `Deal công nợ` / `Giá MtM tại ngày ps`, **không có dòng
dữ liệu nào**. Công thức `RP_SEC!S2` vẫn tra bảng đó:

```
= IFNA( (XLOOKUP(…[Bonds_ShortName], FNRP!D:D, FNRP!I:I,,0)
       - XLOOKUP(…[DEAL_BUY_ID], Ref!BF:BF, Ref!BG:BG,,0)) * …, …)
```

Bảng rỗng → `XLOOKUP` trả `#N/A` → `IFNA` nuốt → **`S` = `N`, không còn adjust gì**. Cơ chế
vẫn còn nguyên nhưng đang không hoạt động. Cần Jak xác nhận: 7 deal đó đã tất toán (đúng là
nên rỗng), hay danh sách bị xoá nhầm.

### `Ref!AG2 = 40` — đúng

Nhãn `AG1` = *"Ngưỡng check yield bất thường (points)"*, `AG2` = `40`. Khớp memory.

### `FNRP` không chỉ có GOVBOND

Cột `AJ Type` có 3 giá trị: **`TPCP`**, **`TPCQĐP`** (chính quyền địa phương), **`TPCPBL`**
(Chính phủ bảo lãnh). Câu *"FNRP: chỉ chứa GOVBOND"* trong memory là sai nếu hiểu GOVBOND =
TPCP thuần. Ba loại này khác nhau về rủi ro phát hành — đừng gộp khi phân tích issuer.

### Cột team thêm ở `FNRP` là **AG:AQ (11 cột)**, không phải AG:AR (12)

`AG Days` · `AH Days (year)` · `AI Tenor` · `AJ Type` · `AK Issuer` · `AL Quantity_edited` ·
`AM Duration_edited` · `AN Modified Duration` · `AO Conv` · `AP BPV_edited` · `AQ PV100`.

**`AR` trống hoàn toàn** (header rỗng, dữ liệu rỗng). Cột `Coupon%` mà memory liệt kê không tồn
tại — coupon nằm sẵn ở `N Cpn Rate` của nguồn, không cần cột thêm.

### `BB_tufund` là **DY1:EK77 (13 cột)**, không phải DY:EQ

`DY Deal ID` · `DZ Bond code` · `EA Folder` · `EB Amt` · `EC Tenor left` · `ED Yield` ·
`EE ItD` · `EF Yield RP SEC` · `EG Coupon Rate` · `EH Đáo hạn` · `EI % amt FNRP` ·
`EJ Market Value FNRP` · `EK PV01`.

**Không có cột `Last Coupon Date`, không có `AI Full`.** Phần việc đang treo trong memory
(derive last coupon date từ anniversary của maturity, tính accrued interest cộng dồn) **chưa có
mặt trong file này** — hoặc chưa từng lưu, hoặc đã bỏ. Kèm theo đó, cái caveat "giả định coupon
annual chưa xác nhận được vì thiếu field frequency" hiện **không còn treo trên file 01** vì
không còn công thức nào phụ thuộc vào giả định đó.

`DY:EB` (Deal ID, Bond code, Folder, Amt) là **giá trị gõ tay**, không phải công thức. `EC`,
`ED` cũng vậy. Từ `EE` trở đi mới là `XLOOKUP` sang `Outstanding_pos_BB_MSB` và
`Bankingbook_MSB`. Nghĩa là 6 cột đầu của bảng này không có đường truy nguồn — sửa deal ở Kondor
thì bảng này không tự đổi.

### PQ `Tradingbook_MSB` đang là **10 bước**, không phải 4

Memory ghi *"đã tối ưu: 9 bước → 4 bước"*. Bản 18/08:

`Source` → `Changed Type` → `Renamed Columns` → `Removed Other Columns` → `Changed Type2` →
`Renamed Columns1` → `Changed Type1` → `Reordered Columns` → `Changed Type3` → `Filtered Rows`

Bốn bước đổi kiểu và hai bước đổi tên rời rạc, cộng một `Reordered Columns` không cần thiết
(thứ tự cột do `Removed Other Columns` quyết định rồi). `Bankingbook_MSB` cùng nguồn chỉ có
6 bước. Tối ưu đã mất — hoặc chưa bao giờ lưu vào bản này.

**Đề xuất**: gộp về `Source` → `Changed Type` (một lần, đủ 87 cột) → `Renamed Columns` (một
lần) → `Removed Other Columns` → `Filtered Rows`. Nhớ quy tắc top-down: bước sau chỉ được tham
chiếu bước khai báo trên nó.

### `Outstanding_pos` loại 19328–19330 — vẫn đúng

Có ở **2 query**, không phải một:

```m
#"Outstanding pos_BB_MSB"  … [BondsDeals_Id] <> 19328/19329/19330 … and [Folders_ShortName] = "AFS-ALM"
#"Outstanding pos_BB_SBV"  … [BondsDeals_Id] <> 19328/19329/19330 … and [Folders_ShortName] <> "AFS-ITB"
```

Hai bản TB (`_TB_MSB`, `_TB_SBV`) **không** loại. Ba deal này gõ thẳng trong M code — khi senior
duyệt xong phải sửa hai chỗ, sót một chỗ là hai book lệch nhau.

**Đề xuất**: đưa danh sách deal loại trừ vào một table trên `Ref`, query đọc bằng
`Excel.CurrentWorkbook()` rồi `List.Contains` — sửa một chỗ, cả hai query theo.

### Ngày `C3–C6` — `C5` không dùng XLOOKUP

- `C3 Yesterday` — array formula
- `C4 Last month` — `XLOOKUP(EOMONTH(RPT,-1), Ref!BB:BB, …, XLOOKUP(…,,-1), 0)` (XLOOKUP lồng,
  fallback match mode −1)
- `C5 Last year` — `WORKDAY(EOMONTH(DATE(YEAR(RPT),1,1),-1), 0, 0)` — **không phải XLOOKUP**.
  Đây chính là chỗ memory ghi *"C5: đã fix bug WORKDAY"* — bug đã fix, nhưng nó không nằm trong
  "pattern XLOOKUP" như memory mô tả.
- `C6 −14 days` — `XLOOKUP(RPT-14, 'Yield 1Y'!G:G, …,, -1)`

---

## 5. VBA — 8 module, 16 sub

`a_LoadData` · `b_CopyCheck` · `c_Upload53` · `d_ToolPnL` · `e_ToolFIBond` · `f_ToolReport` ·
`ToolCap` · `XuatScriptPowerQuery`

### `Copy_to_listdeal` — bản fix `ThisWorkbook` **không có trong file này**

Memory ghi *"đã fix — đổi `Workbooks(Bond_tool)` → `ThisWorkbook` tại 5 chỗ"*. File 18/08 vẫn là
`Workbooks(Bond_tool)` ở **7 chỗ** trong `b_CopyCheck` (dòng 9, 24, 39, 63, 74, 93, 100, 115).

`Bond_tool` = `Runtool!C22` = `="01.RptTool_Bond_"&TEXT(RPT,"yyyy.mm.dd")&".xlsm"`. Macro chỉ
chạy khi file đang mở **đúng tên đó**. Đổi tên file, mở bản copy, hay `RPT` lệch tên file →
`Workbooks(…)` không tìm thấy → **runtime error 9**. `a_LoadData` cùng file thì dùng
`ThisWorkbook` đúng cách (dòng 19, 47) — hai module đang xử lý cùng một vấn đề theo hai kiểu.

**Đề xuất**: thay 7 chỗ đó bằng `ThisWorkbook`. Ô `C22` vẫn giữ vì query khác dùng.

### `Keo_cong_thuc` — sai chính tả hằng số

```vb
Application.Calculation = xlCalculationAutomat
```

Thiếu `ic`. Không có `Option Explicit` nên VBA coi đây là biến rỗng → gán `0` vào
`Application.Calculation` → **error 1004** khi nhánh `dieukien > 4` chạy. Đúng phải là
`xlCalculationAutomatic`. Dòng tương ứng ở `Refresh_query` (`= xlAutomatic`) thì đúng giá trị.

### Chuỗi kết nối DB53 nhúng thẳng trong code

`c_Upload53` mở kết nối SQL Server ở **8 sub**, mỗi sub lặp lại nguyên chuỗi
`Provider=SQLOLEDB; Data Source=…; Initial Catalog=…; User ID=…; Password=…` — **tài khoản và
mật khẩu viết thẳng trong VBA, không mã hoá**.

Hai vấn đề tách bạch:

1. **Bảo mật** — bất kỳ ai mở file đều đọc được thông tin đăng nhập DB. File này đang nằm trên
   file server dùng chung. Đây là việc cần báo lên, không phải việc sửa bằng code.
2. **Bảo trì** — 8 bản sao của cùng một chuỗi. Đổi mật khẩu là sửa 8 chỗ.

**Đề xuất (chỉ giải quyết vế 2)**: tách ra `Private Function MoKetNoi() As Object` dùng chung.
Vế 1 phải theo chính sách của IT/bảo mật.

---

## 6. Chuỗi dữ liệu — bản đã verify

```
Murex CSV (RPT-dated UNC)
   │  Param("Input", n) — mọi đường dẫn dựng từ Runtool!C2
   ├─► PQ Tradingbook_MSB / Bankingbook_MSB / _SBV ──► FNRP
   ├─► PQ Outstanding pos_* (Kondor xlsx)         ──► RP_SEC
   └─► PQ RPBOD3 / FNRP_Repo / Market            ──► Deal ps / Funding / Market
             │
             ▼
     Phan tich_TB (25 vùng) · Phan tich_BB (3 mục, 10 vùng)
             │
             ├─► Double check (51 kiểm tra)  ← cửa kiểm số, phải xanh trước khi đẩy
             ▼
     His.TB · His.BB · His.TB.SBV · His.BB.SBV · His.Other · His.DataTT
             │  VBA c_Upload53 (8 sub, ADODB)
             ▼
     DB53 (SQL Server, DB GS_RRTT)
             │
             ▼
     File 02 → HTML
```

`Phan tich_TB` chia 3 mục: `01. TRADING BOOK_MSB` (A:CQ) · `02. BOND FUTURE` (CR:DG) ·
`03. TRADING BOOK_SBV` (DH:EO) — tổng 25 vùng, khớp memory.

`Phan tich_BB` chia 3 mục: `01. Banking book_MSB` (A:AK, 6 vùng) · `02. Banking book_SBV`
(AL:AZ, 3 vùng) · `03. Repo & Reverse repo` (BA:BE, 1 vùng) — **10 vùng, không phải 12**.

Điểm silent-failure mà memory cảnh báo vẫn đúng nguyên: File 02 chỉ đọc DB53. VBA upload hỏng
thì File 02 vẫn chạy, vẫn ra báo cáo, chỉ là số cũ. `Double check` dòng 37–40 có kiểm
`His.BB` / `His.TB.SBV` / `His.BB.SBV` / `His.Other` đẩy lên 53 có lỗi không — **nhưng thiếu
dòng kiểm cho `His.TB`**. Book lớn nhất lại là book duy nhất không có dòng kiểm upload.

**Đề xuất**: thêm một dòng kiểm `His.TB` vào `Double check`, cùng dạng với dòng 37.

---

## 7. Sheet `Log` — 258 bản ghi, dừng ở 16/07/2026

Log trải từ 10/05/2012 đến 16/07/2026, phân nhóm `Methodology` / `Manual` / `Data` / `Deal` /
`Tính toán`. Bản ghi cuối là 16/07 — **một tháng trước ngày chốt của file này**.

Những thay đổi phát hiện ở mục 4 (danh sách override `Ref!BF:BG` rỗng đi, PQ `Tradingbook_MSB`
quay lại 10 bước, `BB_tufund` không có cột last coupon) đều **không có dòng log nào**. Log là cơ
chế truy vết duy nhất của file — nếu số liệu đổi mà log không ghi thì tháng sau không ai dựng
lại được vì sao.

---

## 8. Việc cần Jak quyết

| # | Việc | Cần gì từ Jak |
|---|---|---|
| 1 | `I9` có `+1` che chênh 618,6 triệu | Xác nhận bỏ `+1`, rồi truy `BS!G6/L6/Q6` |
| 2 | `Ref!BF:BG` rỗng | 7 deal đã tất toán, hay bị xoá nhầm? |
| 3 | `BB_tufund` không có cột last coupon / AI Full | Bỏ hẳn, hay cần dựng lại? |
| 4 | 19328–19330 | Senior đã duyệt chưa? Duyệt rồi thì sửa **2** query |
| 5 | `C19` ≡ `C32` | Chọn ô nào làm nguồn |
| 6 | `C18` = 30/06 | Mốc QLHS này còn đúng không |
| 7 | Chuỗi kết nối DB53 lộ mật khẩu | Báo IT/bảo mật — ngoài phạm vi sửa file |

Mục 1–6 sửa được ngay khi có xác nhận. Đúng quy tắc làm việc: **sửa từng bước một, confirm rồi
mới sang bước kế tiếp**, không dùng helper column.

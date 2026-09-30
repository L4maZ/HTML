# Đợt 5 — Bảng đề xuất tên chỉ số đầy đủ (chờ Jak duyệt)

Quét 30/09/2026 trên toàn bộ chữ hiển thị của báo cáo (tab, tiêu đề, tên cột, tên dòng, nhãn KPI, chú giải chart, trục chart), nạp `Key_20260929.xlsx`. Có **139 nhãn** chứa từ viết tắt hoặc thuật ngữ, gom lại còn các nhóm dưới đây.

**Cách áp dụng khi duyệt xong:** đổi tên **chỉ ở lớp hiển thị** (một bảng `LABEL_FULL` trong HTML). Tên gốc từ Excel (`Face Value`, `ItD Unrealized MtM PnL`, `PV01`…) vẫn dùng làm khoá tra cứu nên không gãy số liệu. Nhãn nào không có trong bảng thì hiện nguyên như cũ. Bảng nằm trong HTML nên sửa tên sau này không cần đụng code.

Cột "Duyệt": ghi **OK**, hoặc ghi tên bạn muốn, hoặc **giữ** nếu muốn để nguyên.

## A. ItD chỉ ghi vắn tắt (mục bạn đã nêu)

| # | Hiện tại | Ở đâu | Đề xuất | Duyệt |
|---|---|---|---|---|
| A1 | `Trading · ItD` / `Banking · ItD` | Trạng thái → Cơ cấu theo thời gian nắm giữ (cột) | `Trading · ItD Unrealized PnL` / `Banking · ItD Unrealized PnL` | |
| A2 | `ItD 29/09/2026`, `ItD 28/09/2026`, `ItD 28/08/2026` | Tab ItD Unrealized PnL (cột, 2 bảng) | `ItD Unrealized PnL` xuống dòng rồi ngày: `ItD Unrealized PnL` / `29/09/2026` | |
| A3 | `ItD Trading`, `ItD Banking`, trục `ItD · tỷ` | Trạng thái → chart nhóm thời gian nắm giữ | `ItD Unrealized PnL · Trading`, `… · Banking`, trục `ItD Unrealized PnL · tỷ` | |
| A4 | `Trạng thái & ItD theo nhóm thời gian nắm giữ` | Trạng thái (tiêu đề chart) | `Trạng thái & ItD Unrealized PnL theo nhóm thời gian nắm giữ` | |
| A5 | `ItD toàn book Trading` / `ItD toàn book Banking` | Tab ItD Unrealized PnL (KPI) | `ItD Unrealized PnL toàn book Trading` / `… Banking` | |
| A6 | `Lỗ tiềm ẩn (ItD)` | Sử dụng vốn (chú giải chart) | `Lỗ tiềm ẩn (ItD Unrealized PnL)` | |
| A7 | `ItD hiện tại`, `tác động ItD · tỷ VND` | QTRR theo kịch bản | `ItD Unrealized PnL hiện tại`, `tác động lên ItD Unrealized PnL · tỷ VND` | |
| A8 | `ItD Unrealized PnL · GBond Trading`, `Tổng GBond` | Highlight, tab ItD (KPI) | `… · GovBond Trading`, `Tổng GovBond` (thống nhất `GovBond`, hiện đang lẫn `GBond`) | |
| A9 | `ItD Unrealized MtM PnL` | Indicator (nhãn từ Excel) | **giữ nguyên** (đã đầy đủ, đúng nhãn Excel) | |

## B. Kỳ so sánh

| # | Hiện tại | Ở đâu | Đề xuất | Duyệt |
|---|---|---|---|---|
| B1 | `+/− DtD`, `+/− MtD`, `+/− YtD`, `DtD`, `MtD`, `YtD` | Hầu hết các bảng | **Giữ** (đúng ký hiệu trên File 02) và thêm **một dòng chú giải** dưới tiêu đề trang Indicator: `DtD: so với ngày làm việc trước · MtD: so với cuối tháng trước · YtD: so với đầu năm` | |
| B2 | `sàn ≥ 60% · +0.11 đpt DtD` | FI Bond & CD (KPI) | `… · +0.11 điểm % so với hôm trước` (`đpt` là viết tắt khó hiểu) | |
| B3 | `Daily PnL`, `MtD PnL`, `YtD PnL`, và `Ytd PnL` (Banking) | Indicator | thống nhất viết `YtD PnL` (sửa `Ytd` của Banking); tuỳ chọn thêm tiếng Việt: `Daily PnL (trong ngày)` | |
| B4 | `Ngày BC so với 15/9/2026` | QTRR theo kịch bản | `Ngày báo cáo so với 15/09/2026` (đủ chữ, đủ số 0) | |

## C. Bảng cân đối và ghi nhận

| # | Hiện tại | Ở đâu | Đề xuất | Duyệt |
|---|---|---|---|---|
| C1 | `MtM (clean)` | Trạng thái → Cấu trúc theo lớp ghi nhận | `Giá trị thị trường (MtM, clean)` | |
| C2 | `PT/CK còn lại` | cùng bảng | `Phụ trội/chiết khấu còn lại (PT/CK)` — **xác nhận PT/CK = phụ trội/chiết khấu** | |
| C3 | `On BS (nội bảng, đã thanh toán)`, `Off BS (đã valid, chưa thanh toán)`, `Ngoài BS (BO chưa xác nhận)` | cùng bảng | `Nội bảng (đã thanh toán)`, `Ngoại bảng (đã valid, chưa thanh toán)`, `Ngoài bảng cân đối (Back Office chưa xác nhận)` | |
| C4 | `YTM onBS` | Indicator | `YTM nội bảng` | |

## D. Dòng chỉ tiêu ở Indicator (nhãn tiếng Anh từ Excel)

| # | Hiện tại | Đề xuất | Duyệt |
|---|---|---|---|
| D1 | `Face Value` | `Face Value (mệnh giá)` | |
| D2 | `Average Price/1bond` | `Giá bình quân mỗi bond` | |
| D3 | `Bond futures` | `Hợp đồng tương lai trái phiếu (Bond futures)` | |
| D4 | `Theo TT41` | `Theo Thông tư 41 (TT41)` | |
| D5 | `a) Book AFS`, `b) Book HTM` | `a) Book AFS (sẵn sàng để bán)`, `b) Book HTM (nắm giữ đến ngày đáo hạn)` — hoặc **giữ** nếu desk đều quen AFS/HTM | |
| D6 | `VaR` | `VaR (giá trị chịu rủi ro)` | |
| D7 | `CVaR/ES` và `(Expected Shortfall)` đang tách thành hai nhãn ở hai dòng | gộp thành một nhãn `CVaR / ES (Expected Shortfall)` | |
| D8 | `Time holding (days)` | `Thời gian nắm giữ bình quân (ngày)` | |
| D9 | `Duration (Years)` | `Duration (năm)` | |
| D10 | `Number of bondID held > 1Y` | `Số bond nắm giữ trên 1 năm` | |
| D11 | `Coupon rate`, `Repo rate` | `Lãi suất Coupon`, `Lãi suất Repo` | |
| D12 | `Concentration_tenor > 20Y`, `Concentration_TPCQDP` | `Tỷ trọng kỳ hạn trên 20 năm`, `Tỷ trọng TPCQDP` | |
| D13 | `YTM` | `YTM (lợi suất đến khi đáo hạn)` | |

## E. Cơ cấu theo tổ chức phát hành

| # | Hiện tại | Đề xuất | Duyệt |
|---|---|---|---|
| E1 | `TPCP` | `TPCP (Chính phủ)` | |
| E2 | `TPCPBL` | `TPCPBL (Chính phủ bảo lãnh)` | |
| E3 | `TPCQDP` | `TPCQDP (Chính quyền địa phương)` | |

## F. Sử dụng vốn và Lãi/lỗ

| # | Hiện tại | Đề xuất | Duyệt |
|---|---|---|---|
| F1 | `… BC KQKD (= a + b)` | `… Báo cáo kết quả kinh doanh (BC KQKD) (= a + b)` | |
| F2 | `NIM` (dòng b, legend, dòng `2. Lãi/lỗ về NIM`) | `NIM (chênh lệch lãi suất ròng)` ở lần đầu mỗi bảng; các chỗ sau giữ `NIM` | |
| F3 | `Vốn yêu cầu TT41` (legend) | `Vốn yêu cầu theo Thông tư 41` | |
| F4 | `KB1: Giữ DM đến khi đáo hạn Repo` (KB2, KB3) | `Kịch bản 1: Giữ danh mục đến khi đáo hạn Repo` (DM = danh mục) | |
| F5 | `PnL theo phương pháp QLHS` (2 tiêu đề) | `PnL theo phương pháp quản lý hiệu suất (QLHS)` — **xác nhận QLHS = quản lý hiệu suất** | |
| F6 | `Kịch bản VaR` | `Kịch bản VaR (giá trị chịu rủi ro)` | |

## G. Thị trường, FI Bond & CD

| # | Hiện tại | Đề xuất | Duyệt |
|---|---|---|---|
| G1 | `KLGD (tỷ VND)` | `Khối lượng giao dịch (tỷ VND)` | |
| G2 | `GTTT sơ cấp`, `GTGD thứ cấp` | `Giá trị trúng thầu sơ cấp`, `Giá trị giao dịch thứ cấp` | |
| G3 | `giao dịch TPDN` | `giao dịch trái phiếu doanh nghiệp (TPDN)` | |
| G4 | `Tỷ lệ đầu tư an toàn (% VTC)`, `Đầu tư an toàn / VTC`, `trần 7% VTC` | thay `VTC` bằng `vốn tự có` — **xác nhận VTC = vốn tự có** | |
| G5 | `3. CD ( chưa điều chỉnh)` (sai khoảng trắng), `CD` | `3. CD (chưa điều chỉnh)`; chỗ đầu tiên: `CD (chứng chỉ tiền gửi)` | |
| G6 | `4. ITB` | **cần bạn cho biết ITB là gì** (transcript ghi "International bond book"?) rồi ghi đầy đủ | |
| G7 | `FI Bond`, `FI Bond & CD` | **cần bạn cho biết FI = gì** (Financial Institution?), ghi ở lần đầu: `FI Bond (…)` | |
| G8 | `Coupon BQ danh mục` | `Coupon bình quân danh mục` | |
| G9 | `Spread vs GovBond (bps)` | `Chênh lệch so với GovBond (bps)` | |
| G10 | `VIRA`, `Dự báo từ khảo sát VIRA (Average)` | giữ tên riêng; **nếu có tên đầy đủ của VIRA** cho tôi để ghi chú một lần | |

## Cần bạn quyết trước khi tôi làm

1. **B1** — giữ `DtD/MtD/YtD` và thêm dòng chú giải, hay đổi hết sang tiếng Việt?
2. **Kiểu viết**: "Việt hoá + viết tắt trong ngoặc" (như trên) ở mọi chỗ, hay chỉ ở lần đầu mỗi bảng?
3. Cột tiêu đề có ngày (A2): xuống dòng hai tầng (đề xuất) hay để một dòng dài?
4. Trả lời các mục **cần xác nhận**: C2 (PT/CK), F5 (QLHS), G4 (VTC), G6 (ITB), G7 (FI), G10 (VIRA).

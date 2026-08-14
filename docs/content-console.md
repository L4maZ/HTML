# Content Console pattern

Loại sản phẩm thứ hai, khác hẳn dashboard tài chính: **công cụ đọc / học / tra cứu** dựng từ
nội dung dài, không có time-series và không có ECharts.

Hai file thuộc loại này: `cfa_l1_console.html` (11,175 dòng, 990KB) và `QC.RR.022_lampd3.html`.

## Khác gì dashboard

| | Dashboard | Content console |
|---|---|---|
| Nguồn dữ liệu | Excel upload → RAW object | Nội dung viết tay trong file |
| Thư viện | ECharts inline | Không, hoặc KaTeX |
| Đơn vị nội dung | Chart + KPI + bảng | Block prose có cấu trúc |
| State | Không lưu | `localStorage` (tiến độ, highlight) — công cụ cá nhân, không phải deliverable của bank |
| Điều hướng | Sidebar trang | Screen system nhiều tầng + phím tắt |

## DSL nội dung

Nội dung là **dữ liệu, không phải HTML**. Author viết object, renderer lo phần markup.

```js
var T = (id,name,range,weight,readings) => ({id,name,range,weight,readings:readings||[]});
var R = (num,title,modules) => ({num,title,modules});

var QM = T('qm','Quantitative Methods','1–11','6–9%',[ /* readings */ ]);
```

`renderBlock(b)` switch trên `b.t`:

| `t` | Nghĩa |
|---|---|
| `h` / `sh` | Heading cấp 1 / 2 |
| `p` | Đoạn văn |
| `ul` | Bullet list |
| `f` | Công thức đơn (`tex`, `cap`, `alt`) |
| `fm` | Nhóm công thức |
| `ex` | Ví dụ có `body` + `answer` |
| `note` | Ghi chú, `warn:true` thì đổi màu |
| `tbl` | Bảng (`head`, `rows`, `cap`, `num` = căn phải cột số) |
| `deep` | `<details>` gập lại — phần đào sâu, mặc định ẩn |

Lợi ích: thêm loại block mới chỉ cần thêm một `case`; toàn bộ nội dung cũ không đụng tới.

## KaTeX degrade — cùng triết lý với rule "không bao giờ undefined"

CDN KaTeX bị chặn thì công thức vẫn **đọc được**, không mất nội dung:

```js
function renderMath(root){
  if(!window.katex) return;                    // không có KaTeX → thoát êm, .fraw giữ nguyên
  root.querySelectorAll('.fml[data-tex]:not([data-done])').forEach(el=>{
    const tgt = el.querySelector('.fraw');
    try{
      katex.render(el.getAttribute('data-tex'), tgt, {throwOnError:false, displayMode:true});
      tgt.classList.remove('fraw');
      el.setAttribute('data-done','1');
    }catch(e){ el.setAttribute('data-done','fail'); }   // hỏng 1 công thức không giết cả trang
  });
}
```

Markup luôn chứa sẵn text thô (`alt` hoặc chính `tex`) trong `.fraw`. KaTeX chỉ **nâng cấp**
thứ đã đọc được, không phải điều kiện để hiển thị. Đây chính là rule fallback của key system,
áp cho công thức.

## Screen system

Phẳng hơn dashboard: mọi màn hình là `.screen`, bật bằng class `.on`.

```js
function go(id){
  document.querySelectorAll('.screen').forEach(s=>s.classList.remove('on'));
  const el = document.getElementById('scr-'+id);
  if(el) el.classList.add('on');
  window.scrollTo(0,0);
  if(id==='hub-study')    renderStudyGrid();      // render lazy theo màn hình
  if(id==='hub-practice') renderBankGrid();
  if(id==='hub-tools')  { renderHighlightList(); renderProgressList(); }
  if(id==='start')        renderStartStats();
}
```

Phím tắt: `Esc` lùi một cấp; ở màn start có `S` / `P` / `T`. Luôn bỏ qua khi focus đang ở
`INPUT` / `TEXTAREA`.

## Tone Ink/Gold/Paper

Tone thứ tư, chỉ dùng cho content console — nền tối, chữ kem, accent vàng, panel màu giấy:

| Token | Hex |
|---|---|
| `--ink` / `--ink2` / `--ink3` | `#10162a` / `#161d36` / `#1e2745` |
| `--paper` / `--paper2` | `#f6f1e4` / `#ede6d3` |
| `--line` | `#d9cca0` |
| `--gold` / `--goldl` | `#b6862f` / `#d6a94a` |
| `--teal` / `--tealb` | `#1f6f6b` / `#2c8f89` |
| `--text` / `--muted` | `#1c2333` / `#5a5645` |

Semantic giữ nguyên như dashboard: `--red:#9B2C2C`, `--green:#276749`, `--amber:#a56a1b`.

Ba họ chữ có vai trò rõ ràng:

```css
--sans:  -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;  /* thân bài */
--serif: Georgia,"Times New Roman",serif;        /* heading, tiêu đề màn hình */
--mono:  ui-monospace,"SF Mono",Menlo,Consolas,monospace;  /* nhãn, mã, phím tắt */
```

Cả ba đều là font hệ thống → **không phụ thuộc mạng**, khác dashboard vốn tải DM Sans.

## localStorage ở đây là hợp lệ

```js
const LS = { prog:'cfa1.progress.v1', hl:'cfa1.highlights.v1', quiz:'cfa1.quizstate.v1' };
function lsGet(k, d){ try{ const v=localStorage.getItem(k); return v?JSON.parse(v):d; }catch(e){ return d; } }
function lsSet(k, v){ try{ localStorage.setItem(k, JSON.stringify(v)); }catch(e){} }
```

Rule cấm localStorage áp cho **báo cáo giao cho bank** (dữ liệu phải đi theo file export).
Công cụ học cá nhân thì tiến độ và highlight nằm ở localStorage là đúng chỗ — miễn là
`try/catch` và có giá trị mặc định, mất state thì app vẫn chạy.

Đặt tên key có version (`.v1`) để đổi cấu trúc sau này không vỡ dữ liệu cũ.

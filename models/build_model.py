# -*- coding: utf-8 -*-
"""IPO advisory model - FICTIONAL company, illustrative data only."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

NAVY   = "1F3864"
GOLD   = "BF9000"
LIGHT  = "DDEBF7"
INPUT  = "FFF2CC"
GREY   = "F2F2F2"

H1 = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
H2 = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BOLD = Font(name="Calibri", size=11, bold=True)
NORM = Font(name="Calibri", size=11)
ITAL = Font(name="Calibri", size=9, italic=True, color="808080")
INP  = Font(name="Calibri", size=11, color="0000FF")   # blue = hardcoded input
FRM  = Font(name="Calibri", size=11, color="000000")   # black = formula

fill_navy = PatternFill("solid", fgColor=NAVY)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_lite = PatternFill("solid", fgColor=LIGHT)
fill_inp  = PatternFill("solid", fgColor=INPUT)
fill_grey = PatternFill("solid", fgColor=GREY)
thin = Side(style="thin", color="BFBFBF")
box  = Border(left=thin, right=thin, top=thin, bottom=thin)

YEARS = ["FY23A","FY24A","FY25A","FY26E","FY27E","FY28E"]
COLS  = ["C","D","E","F","G","H"]   # year columns

def title(ws, text, sub=""):
    ws["A1"] = text; ws["A1"].font = H1; ws["A1"].fill = fill_navy
    for c in range(1, 14):
        ws.cell(row=1, column=c).fill = fill_navy
    ws["A2"] = sub or "Illustrative - fictional data. Not based on any real company."
    ws["A2"].font = ITAL
    ws.freeze_panes = "C5"

def hdr(ws, row):
    ws.cell(row=row, column=1, value="Line item").font = H2
    ws.cell(row=row, column=2, value="Unit").font = H2
    for i, y in enumerate(YEARS):
        c = ws.cell(row=row, column=3+i, value=y); c.font = H2; c.alignment = Alignment(horizontal="center")
    for c in range(1, 9):
        ws.cell(row=row, column=c).fill = fill_navy
        ws.cell(row=row, column=c).border = box

def put(ws, row, label, unit, vals, fmt="#,##0.0", is_input=True, bold=False, section=False):
    ws.cell(row=row, column=1, value=label).font = BOLD if (bold or section) else NORM
    if section:
        for c in range(1, 9):
            ws.cell(row=row, column=c).fill = fill_lite
        return
    ws.cell(row=row, column=2, value=unit).font = NORM
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=3+i, value=v)
        c.number_format = fmt
        c.font = INP if is_input else FRM
        if is_input: c.fill = fill_inp
        if bold: c.font = Font(name="Calibri", size=11, bold=True, color="0000FF" if is_input else "000000")
        c.border = box

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

# ---------------------------------------------------------------- README
ws = wb.active; ws.title = "README"
title(ws, "PROJECT LOTUS - IPO ADVISORY MODEL (FICTIONAL)")
rows = [
 ("Target (fictional)", "NovaPay Digital JSC - Vietnam e-wallet + merchant payment gateway"),
 ("Adviser (fictional)", "Meridian Partners Securities - proposed Lead Manager / Bookrunner"),
 ("Proposed venue", "HOSE"),
 ("Currency", "USD millions unless stated otherwise"),
 ("Fiscal year", "FY ends 31 December"),
 ("", ""),
 ("DISCLAIMER", "Every company name, transaction, KPI and valuation figure in this file is INVENTED"),
 ("", "for an illustrative pitch exercise. No real issuer, bank, exchange or deal is referenced."),
 ("", "Regulatory points are described generically as 'per prevailing listing regulations'."),
 ("", ""),
 ("Colour convention", "BLUE font on yellow fill = hardcoded assumption. BLACK font = formula."),
 ("", ""),
 ("Sheet", "Contents"),
 ("Operating_KPI", "TPV, take rate, MAU, merchants - 3Y historical + 3Y forecast"),
 ("Financials", "Revenue build, gross margin, EBITDA, net income, net cash"),
 ("Comps", "5 fictional listed payment/fintech peers - EV/Rev, EV/EBITDA, P/E"),
 ("Precedents", "4 fictional fintech M&A / IPO precedent transactions"),
 ("DCF", "Unlevered FCF, WACC build, terminal value, equity bridge"),
 ("Sensitivity", "DCF equity value vs WACC and terminal growth"),
 ("Football_Field", "Valuation ranges by methodology + IPO discount -> offer range"),
 ("Offer_Structure", "Primary/secondary split, proceeds, free float, use of proceeds"),
 ("IPO_Timeline", "Indicative execution calendar"),
 ("Slide_Numbers", "Every number that appears on a slide, with its source cell"),
]
r = 4
for a, b in rows:
    ws.cell(row=r, column=1, value=a).font = BOLD
    ws.cell(row=r, column=2, value=b).font = NORM
    r += 1
widths(ws, {"A": 22, "B": 95})

# ---------------------------------------------------------------- KPI
ws = wb.create_sheet("Operating_KPI")
title(ws, "OPERATING KPIs")
hdr(ws, 4)
put(ws, 5,  "Total payment volume (TPV)", "USD bn", [8.0,11.2,15.1,19.6,24.9,31.1], "#,##0.0", True, bold=True)
ws.cell(row=6, column=1, value="TPV growth y/y").font = NORM
ws.cell(row=6, column=2, value="%").font = NORM
for i in range(1,6):
    c = ws.cell(row=6, column=4+i-1+0)
for i in range(1,6):
    col = COLS[i]; prev = COLS[i-1]
    c = ws[f"{col}6"]; c.value = f"={col}5/{prev}5-1"; c.number_format = "0.0%"; c.border = box
put(ws, 7,  "Take rate (blended)", "%", [0.0062,0.0066,0.0070,0.0072,0.0074,0.0075], "0.00%", True)
put(ws, 8,  "Monthly active users (MAU)", "m", [6.2,8.1,10.3,12.4,14.5,16.4], "#,##0.0", True)
put(ws, 9,  "Active merchants", "'000", [118,172,245,320,405,490], "#,##0", True)
ws.cell(row=10, column=1, value="TPV per MAU").font = NORM
ws.cell(row=10, column=2, value="USD/yr").font = NORM
for col in COLS:
    c = ws[f"{col}10"]; c.value = f"={col}5*1000/{col}8"; c.number_format = "#,##0"; c.border = box
ws.cell(row=11, column=1, value="TPV per merchant").font = NORM
ws.cell(row=11, column=2, value="USD'000/yr").font = NORM
for col in COLS:
    c = ws[f"{col}11"]; c.value = f"={col}5*1000000/{col}9/1000"; c.number_format = "#,##0"; c.border = box
ws["A13"] = "Illustrative - fictional data"; ws["A13"].font = ITAL
widths(ws, {"A": 32, "B": 12, **{c: 11 for c in COLS}})

# ---------------------------------------------------------------- Financials
ws = wb.create_sheet("Financials")
title(ws, "FINANCIAL SUMMARY (USD m)")
hdr(ws, 4)
ws.cell(row=5, column=1, value="Payment revenue (TPV x take rate)").font = NORM
ws.cell(row=5, column=2, value="USD m").font = NORM
for col in COLS:
    c = ws[f"{col}5"]; c.value = f"=Operating_KPI!{col}5*1000*Operating_KPI!{col}7"; c.number_format="#,##0.0"; c.border=box
put(ws, 6, "Value-added services revenue", "USD m", [4.0,6.5,10.5,15.5,22.0,30.0], "#,##0.0", True)
ws.cell(row=7, column=1, value="Total revenue").font = BOLD
ws.cell(row=7, column=2, value="USD m").font = BOLD
for col in COLS:
    c = ws[f"{col}7"]; c.value = f"=SUM({col}5:{col}6)"; c.number_format="#,##0.0"; c.font=BOLD; c.border=box; c.fill=fill_grey
ws.cell(row=8, column=1, value="Revenue growth y/y").font = NORM
ws.cell(row=8, column=2, value="%").font = NORM
for i in range(1,6):
    col=COLS[i]; prev=COLS[i-1]
    c=ws[f"{col}8"]; c.value=f"={col}7/{prev}7-1"; c.number_format="0.0%"; c.border=box
put(ws, 9, "Gross margin", "%", [0.41,0.44,0.47,0.49,0.51,0.52], "0.0%", True)
ws.cell(row=10, column=1, value="Gross profit").font = NORM
ws.cell(row=10, column=2, value="USD m").font = NORM
for col in COLS:
    c=ws[f"{col}10"]; c.value=f"={col}7*{col}9"; c.number_format="#,##0.0"; c.border=box
put(ws, 11, "EBITDA margin", "%", [0.030,0.080,0.130,0.160,0.190,0.210], "0.0%", True)
ws.cell(row=12, column=1, value="EBITDA").font = BOLD
ws.cell(row=12, column=2, value="USD m").font = BOLD
for col in COLS:
    c=ws[f"{col}12"]; c.value=f"={col}7*{col}11"; c.number_format="#,##0.0"; c.font=BOLD; c.border=box; c.fill=fill_grey
put(ws, 13, "Depreciation & amortisation", "USD m", [3.5,4.6,6.0,7.5,9.3,11.3], "#,##0.0", True)
ws.cell(row=14, column=1, value="EBIT").font = NORM
ws.cell(row=14, column=2, value="USD m").font = NORM
for col in COLS:
    c=ws[f"{col}14"]; c.value=f"={col}12-{col}13"; c.number_format="#,##0.0"; c.border=box
put(ws, 15, "Net interest & other income", "USD m", [0.6,0.9,1.4,1.8,2.4,3.2], "#,##0.0", True)
ws.cell(row=16, column=1, value="Profit before tax").font = NORM
ws.cell(row=16, column=2, value="USD m").font = NORM
for col in COLS:
    c=ws[f"{col}16"]; c.value=f"={col}14+{col}15"; c.number_format="#,##0.0"; c.border=box
put(ws, 17, "Effective tax rate", "%", [0.20]*6, "0.0%", True)
ws.cell(row=18, column=1, value="Tax expense").font = NORM
ws.cell(row=18, column=2, value="USD m").font = NORM
for col in COLS:
    c=ws[f"{col}18"]; c.value=f"=-MAX(0,{col}16)*{col}17"; c.number_format="#,##0.0"; c.border=box
ws.cell(row=19, column=1, value="Net income").font = BOLD
ws.cell(row=19, column=2, value="USD m").font = BOLD
for col in COLS:
    c=ws[f"{col}19"]; c.value=f"={col}16+{col}18"; c.number_format="#,##0.0"; c.font=BOLD; c.border=box; c.fill=fill_grey
ws.cell(row=20, column=1, value="Net margin").font = NORM
ws.cell(row=20, column=2, value="%").font = NORM
for col in COLS:
    c=ws[f"{col}20"]; c.value=f"={col}19/{col}7"; c.number_format="0.0%"; c.border=box
put(ws, 22, "Capital expenditure", "USD m", [7.0,9.0,11.0,12.5,15.0,18.0], "#,##0.0", True)
put(ws, 23, "Increase in net working capital", "USD m", [1.5,2.0,2.6,3.0,3.5,4.0], "#,##0.0", True)
put(ws, 24, "Net cash (cash less debt), year end", "USD m", [18.0,24.0,36.0,52.0,78.0,118.0], "#,##0.0", True, bold=True)
ws["A26"] = "Illustrative - fictional data"; ws["A26"].font = ITAL
widths(ws, {"A": 34, "B": 12, **{c: 11 for c in COLS}})

# ---------------------------------------------------------------- Comps
ws = wb.create_sheet("Comps")
title(ws, "TRADING COMPARABLES (FICTIONAL LISTED PEERS)")
heads = ["Peer (fictional)","Region","EV (USD m)","FY26E Rev (USD m)","FY26E EBITDA (USD m)","FY26E NI (USD m)","EV/Revenue","EV/EBITDA","P/E"]
for i,h in enumerate(heads):
    c = ws.cell(row=4, column=1+i, value=h); c.font=H2; c.fill=fill_navy; c.alignment=Alignment(wrap_text=True, horizontal="center"); c.border=box
peers = [
 ("Meridian Pay Holdings","SE Asia", 11020, 1900, 355, 262),
 ("Sundara Digital Payments","SE Asia", 5290, 1150, 200, 151),
 ("Kalisa Fintech Group","SE Asia", 2340, 600, 106, 77),
 ("Northwind Payments Inc.","Global", 24960, 3900, 745, 555),
 ("Arcadia Merchant Services","Global", 8500, 1700, 309, 233),
]
r=5
for n,reg,ev,rev,eb,ni in peers:
    ws.cell(row=r,column=1,value=n).font=NORM
    ws.cell(row=r,column=2,value=reg).font=NORM
    for j,v in enumerate([ev,rev,eb,ni]):
        c=ws.cell(row=r,column=3+j,value=v); c.number_format="#,##0"; c.font=INP; c.fill=fill_inp; c.border=box
    ws.cell(row=r,column=7).value=f"=C{r}/D{r}"; ws.cell(row=r,column=7).number_format='0.0"x"'
    ws.cell(row=r,column=8).value=f"=C{r}/E{r}"; ws.cell(row=r,column=8).number_format='0.0"x"'
    ws.cell(row=r,column=9).value=f"=(C{r}*0.97)/F{r}"; ws.cell(row=r,column=9).number_format='0.0"x"'
    for cc in (7,8,9): ws.cell(row=r,column=cc).border=box
    r+=1
stats=[("Minimum","MIN"),("Median","MEDIAN"),("Mean","AVERAGE"),("Maximum","MAX")]
for nm,fn in stats:
    ws.cell(row=r,column=1,value=nm).font=BOLD
    for cc in (7,8,9):
        L=get_column_letter(cc)
        c=ws.cell(row=r,column=cc,value=f"={fn}({L}5:{L}9)"); c.number_format='0.0"x"'; c.font=BOLD; c.border=box; c.fill=fill_grey
    r+=1
ws.cell(row=r+1,column=1,value="Applied to NovaPay FY26E").font=BOLD
ws.cell(row=r+2,column=1,value="EV/Revenue range (min-max)").font=NORM
ws.cell(row=r+2,column=3,value="=G10*Financials!F7").number_format="#,##0"
ws.cell(row=r+2,column=4,value="=G13*Financials!F7").number_format="#,##0"
ws.cell(row=r+3,column=1,value="  -> implied equity value (EV + net cash)").font=NORM
ws.cell(row=r+3,column=3,value=f"=C{r+2}+Financials!F24").number_format="#,##0"
ws.cell(row=r+3,column=4,value=f"=D{r+2}+Financials!F24").number_format="#,##0"
ws.cell(row=r+4,column=1,value="EV/EBITDA range (min-max)").font=NORM
ws.cell(row=r+4,column=3,value="=H10*Financials!F12").number_format="#,##0"
ws.cell(row=r+4,column=4,value="=H13*Financials!F12").number_format="#,##0"
ws.cell(row=r+5,column=1,value="  -> implied equity value (EV + net cash)").font=NORM
ws.cell(row=r+5,column=3,value=f"=C{r+4}+Financials!F24").number_format="#,##0"
ws.cell(row=r+5,column=4,value=f"=D{r+4}+Financials!F24").number_format="#,##0"
ws.cell(row=r+7,column=1,value="Illustrative - fictional data").font=ITAL
widths(ws, {"A":28,"B":10,"C":14,"D":17,"E":19,"F":16,"G":11,"H":11,"I":11})

# ---------------------------------------------------------------- Precedents
ws = wb.create_sheet("Precedents")
title(ws, "PRECEDENT TRANSACTIONS / IPOs (FICTIONAL)")
heads=["Year","Target (fictional)","Acquirer / Event (fictional)","Type","EV (USD m)","LTM Rev (USD m)","LTM EBITDA (USD m)","EV/Revenue","EV/EBITDA"]
for i,h in enumerate(heads):
    c=ws.cell(row=4,column=1+i,value=h); c.font=H2; c.fill=fill_navy; c.alignment=Alignment(wrap_text=True,horizontal="center"); c.border=box
tx=[(2023,"Lotus Wallet Co.","Arcadia Merchant Services","M&A",540,120,22.5),
    (2024,"Meridian Pay Holdings","IPO (regional exchange)","IPO",7930,1300,264.0),
    (2024,"Orion Gateway Ltd.","Sable Pay (merger)","M&A",1890,350,70.0),
    (2025,"Vantage Digital Pay","IPO (regional exchange)","IPO",3600,500,106.0)]
r=5
for y,t,a,ty,ev,rev,eb in tx:
    ws.cell(row=r,column=1,value=y).font=INP
    ws.cell(row=r,column=2,value=t).font=NORM
    ws.cell(row=r,column=3,value=a).font=NORM
    ws.cell(row=r,column=4,value=ty).font=NORM
    for j,v in enumerate([ev,rev,eb]):
        c=ws.cell(row=r,column=5+j,value=v); c.number_format="#,##0.0"; c.font=INP; c.fill=fill_inp; c.border=box
    ws.cell(row=r,column=8).value=f"=E{r}/F{r}"; ws.cell(row=r,column=8).number_format='0.0"x"'; ws.cell(row=r,column=8).border=box
    ws.cell(row=r,column=9).value=f"=E{r}/G{r}"; ws.cell(row=r,column=9).number_format='0.0"x"'; ws.cell(row=r,column=9).border=box
    r+=1
for nm,fn in stats:
    ws.cell(row=r,column=2,value=nm).font=BOLD
    for cc in (8,9):
        L=get_column_letter(cc)
        c=ws.cell(row=r,column=cc,value=f"={fn}({L}5:{L}8)"); c.number_format='0.0"x"'; c.font=BOLD; c.fill=fill_grey; c.border=box
    r+=1
ws.cell(row=r+1,column=2,value="Applied to NovaPay FY26E revenue").font=BOLD
ws.cell(row=r+2,column=2,value="Implied EV range (min-max EV/Rev)").font=NORM
ws.cell(row=r+2,column=5,value="=H9*Financials!F7").number_format="#,##0"
ws.cell(row=r+2,column=6,value="=H12*Financials!F7").number_format="#,##0"
ws.cell(row=r+3,column=2,value="Implied equity value (EV + net cash)").font=NORM
ws.cell(row=r+3,column=5,value=f"=E{r+2}+Financials!F24").number_format="#,##0"
ws.cell(row=r+3,column=6,value=f"=F{r+2}+Financials!F24").number_format="#,##0"
ws.cell(row=r+5,column=1,value="Illustrative - fictional data").font=ITAL
widths(ws,{"A":8,"B":24,"C":30,"D":9,"E":13,"F":16,"G":18,"H":11,"I":11})

# ---------------------------------------------------------------- DCF
ws = wb.create_sheet("DCF")
title(ws, "DISCOUNTED CASH FLOW")
DYEARS=["FY26E","FY27E","FY28E","FY29E","FY30E","FY31E","FY32E","FY33E"]
DC=["C","D","E","F","G","H","I","J"]
ws.cell(row=4,column=1,value="WACC build").font=H2; ws.cell(row=4,column=1).fill=fill_navy
wacc_rows=[("Risk-free rate",0.045),("Equity risk premium",0.070),("Country / size premium",0.025),
           ("Beta (levered)",1.15),("Cost of equity",None),("After-tax cost of debt",0.060),
           ("Debt / total capital",0.05),("WACC",None)]
r=5
for nm,v in wacc_rows:
    ws.cell(row=r,column=1,value=nm).font=BOLD if v is None else NORM
    c=ws.cell(row=r,column=2)
    if v is None:
        if nm=="Cost of equity": c.value="=B5+B8*B6+B7"
        else: c.value="=B9*(1-B11)+B10*B11"
        c.font=BOLD; c.fill=fill_grey
    else:
        c.value=v; c.font=INP; c.fill=fill_inp
    c.number_format="0.00" if nm.startswith("Beta") else "0.0%"
    c.border=box; r+=1
ws.cell(row=13,column=1,value="Terminal growth rate (g)").font=BOLD
ws.cell(row=13,column=2,value=0.040).number_format="0.0%"; ws["B13"].font=INP; ws["B13"].fill=fill_inp; ws["B13"].border=box

ws.cell(row=15,column=1,value="Unlevered free cash flow (USD m)").font=H2; ws.cell(row=15,column=1).fill=fill_navy
for i,y in enumerate(DYEARS):
    c=ws.cell(row=15,column=3+i,value=y); c.font=H2; c.fill=fill_navy; c.alignment=Alignment(horizontal="center")
def drow(r,label,vals=None,formula=None,fmt="#,##0.0",bold=False,inp=False):
    ws.cell(row=r,column=1,value=label).font=BOLD if bold else NORM
    for i,col in enumerate(DC):
        c=ws[f"{col}{r}"]
        c.value = vals[i] if vals is not None else formula.replace("{c}",col)
        c.number_format=fmt; c.border=box
        if inp: c.font=INP; c.fill=fill_inp
        elif bold: c.font=BOLD; c.fill=fill_grey
# EBIT: FY26-28 link to Financials; FY29-33 grow
ebit_ext=[None,None,None,1.40,1.34,1.28,1.20,1.14]
ws.cell(row=16,column=1,value="EBIT").font=NORM
for i,col in enumerate(DC):
    c=ws[f"{col}16"]
    if i<3: c.value=f"=Financials!{COLS[3+i]}14"
    else:   c.value=f"={DC[i-1]}16*{ebit_ext[i]}"
    c.number_format="#,##0.0"; c.border=box
drow(17,"Tax on EBIT (20%)",formula="=-{c}16*0.20")
ws.cell(row=18,column=1,value="D&A").font=NORM
for i,col in enumerate(DC):
    c=ws[f"{col}18"]
    c.value=f"=Financials!{COLS[3+i]}13" if i<3 else f"={DC[i-1]}18*1.10"
    c.number_format="#,##0.0"; c.border=box
ws.cell(row=19,column=1,value="Capital expenditure").font=NORM
for i,col in enumerate(DC):
    c=ws[f"{col}19"]
    c.value=f"=-Financials!{COLS[3+i]}22" if i<3 else f"={DC[i-1]}19*1.06"
    c.number_format="#,##0.0"; c.border=box
ws.cell(row=20,column=1,value="Change in net working capital").font=NORM
for i,col in enumerate(DC):
    c=ws[f"{col}20"]
    c.value=f"=-Financials!{COLS[3+i]}23" if i<3 else f"={DC[i-1]}20*1.08"
    c.number_format="#,##0.0"; c.border=box
drow(21,"Unlevered free cash flow",formula="=SUM({c}16:{c}20)",bold=True)
ws.cell(row=22,column=1,value="Discount period (yrs, mid-year)").font=NORM
for i,col in enumerate(DC):
    c=ws[f"{col}22"]; c.value=i+0.5; c.number_format="0.0"; c.border=box
drow(23,"Discount factor",formula="=1/(1+$B$12)^{c}22",fmt="0.000")
drow(24,"PV of FCF",formula="={c}21*{c}23",bold=True)

ws.cell(row=26,column=1,value="PV of forecast FCF (FY26-FY33)").font=NORM
ws["C26"]="=SUM(C24:J24)"; ws["C26"].number_format="#,##0"; ws["C26"].border=box
ws.cell(row=27,column=1,value="Terminal value (Gordon growth, on FY33 FCF)").font=NORM
ws["C27"]="=J21*(1+$B$13)/($B$12-$B$13)"; ws["C27"].number_format="#,##0"; ws["C27"].border=box
ws.cell(row=28,column=1,value="PV of terminal value").font=NORM
ws["C28"]="=C27*J23"; ws["C28"].number_format="#,##0"; ws["C28"].border=box
ws.cell(row=29,column=1,value="Enterprise value").font=BOLD
ws["C29"]="=C26+C28"; ws["C29"].number_format="#,##0"; ws["C29"].font=BOLD; ws["C29"].fill=fill_grey; ws["C29"].border=box
ws.cell(row=30,column=1,value="Plus: net cash (FY26E)").font=NORM
ws["C30"]="=Financials!F24"; ws["C30"].number_format="#,##0"; ws["C30"].border=box
ws.cell(row=31,column=1,value="Equity value (DCF, base case)").font=BOLD
ws["C31"]="=C29+C30"; ws["C31"].number_format="#,##0"; ws["C31"].font=BOLD; ws["C31"].fill=fill_gold; ws["C31"].border=box
ws.cell(row=33,column=1,value="Terminal value as % of EV").font=NORM
ws["C33"]="=C28/C29"; ws["C33"].number_format="0.0%"; ws["C33"].border=box
ws.cell(row=35,column=1,value="Illustrative - fictional data").font=ITAL
widths(ws,{"A":38,"B":12,**{c:11 for c in DC}})

# ---------------------------------------------------------------- Sensitivity
ws = wb.create_sheet("Sensitivity")
title(ws, "DCF EQUITY VALUE SENSITIVITY (USD m)")
ws.cell(row=4,column=1,value="Rows = WACC   |   Columns = terminal growth rate").font=BOLD
waccs=[0.115,0.125,0.135,0.145,0.155]
gs=[0.030,0.035,0.040,0.045,0.050]
ws.cell(row=6,column=1,value="WACC \\ g").font=H2; ws.cell(row=6,column=1).fill=fill_navy
for j,g in enumerate(gs):
    c=ws.cell(row=6,column=2+j,value=g); c.number_format="0.0%"; c.font=H2; c.fill=fill_navy; c.alignment=Alignment(horizontal="center"); c.border=box
for i,w in enumerate(waccs):
    rr=7+i
    c=ws.cell(row=rr,column=1,value=w); c.number_format="0.0%"; c.font=BOLD; c.fill=fill_lite; c.border=box
    for j,g in enumerate(gs):
        # replicate DCF with substituted WACC/g using array-free formula
        pv_terms="+".join([f"DCF!{DC[k]}21/(1+$A{rr})^DCF!{DC[k]}22" for k in range(8)])
        tv=f"(DCF!J21*(1+{chr(66+j)}$6)/($A{rr}-{chr(66+j)}$6))/(1+$A{rr})^DCF!J22"
        cc=ws.cell(row=rr,column=2+j,value=f"={pv_terms}+{tv}+Financials!F24")
        cc.number_format="#,##0"; cc.border=box
        if abs(w-0.135)<1e-9 and abs(g-0.040)<1e-9:
            cc.fill=fill_gold; cc.font=BOLD
ws.cell(row=13,column=1,value="Gold cell = base case (WACC 13.5%, g 4.0%) and ties to DCF!C31").font=ITAL
ws.cell(row=15,column=1,value="Illustrative - fictional data").font=ITAL
widths(ws,{"A":12,"B":14,"C":14,"D":14,"E":14,"F":14})

# ---------------------------------------------------------------- Football field
ws = wb.create_sheet("Football_Field")
title(ws, "VALUATION SUMMARY - FOOTBALL FIELD (EQUITY VALUE, USD m)")
heads=["Methodology","Metric / basis","Low (USD m)","High (USD m)","Midpoint"]
for i,h in enumerate(heads):
    c=ws.cell(row=4,column=1+i,value=h); c.font=H2; c.fill=fill_navy; c.border=box; c.alignment=Alignment(wrap_text=True,horizontal="center")
ws["A5"]="Trading comps - EV/Revenue"; ws["B5"]="FY26E revenue x 3.9x-6.4x"
ws["C5"]="=Comps!C12"; ws["D5"]="=Comps!D12"
ws["A6"]="Trading comps - EV/EBITDA"; ws["B6"]="FY26E EBITDA x 22.0x-33.5x"
ws["C6"]="=Comps!C14"; ws["D6"]="=Comps!D14"
ws["A7"]="Precedent transactions"; ws["B7"]="FY26E revenue x 4.5x-7.2x"
ws["C7"]="=Precedents!E12"; ws["D7"]="=Precedents!F12"
ws["A8"]="DCF"; ws["B8"]="WACC 12.5%-14.5%, g 3.5%-4.5%"
ws["C8"]="=Sensitivity!C10"; ws["D8"]="=Sensitivity!E8"
for r in range(5,9):
    ws.cell(row=r,column=1).font=BOLD
    ws.cell(row=r,column=2).font=NORM
    for cc in (3,4):
        ws.cell(row=r,column=cc).number_format="#,##0"; ws.cell(row=r,column=cc).border=box
    e=ws.cell(row=r,column=5,value=f"=AVERAGE(C{r}:D{r})"); e.number_format="#,##0"; e.border=box
ws["A10"]="Concluded reference equity value range"; ws["A10"].font=BOLD
ws["B10"]="Central overlap of the four methodologies"
ws["C10"]="=MEDIAN(C5:C8)"; ws["D10"]="=MEDIAN(D5:D8)"; ws["E10"]="=AVERAGE(C10:D10)"
for cc in (3,4,5):
    c=ws.cell(row=10,column=cc); c.number_format="#,##0"; c.font=BOLD; c.fill=fill_grey; c.border=box
ws["A12"]="IPO discount to concluded midpoint (low end / high end)"; ws["A12"].font=BOLD
ws["C12"]=0.20; ws["D12"]=0.10
for cc in (3,4):
    c=ws.cell(row=12,column=cc); c.number_format="0%"; c.font=INP; c.fill=fill_inp; c.border=box
ws["A13"]="Pre-money equity value at IPO (post-discount)"; ws["A13"].font=BOLD
ws["C13"]="=$E$10*(1-C12)"; ws["D13"]="=$E$10*(1-D12)"; ws["E13"]="=AVERAGE(C13:D13)"
for cc in (3,4,5):
    c=ws.cell(row=13,column=cc); c.number_format="#,##0"; c.font=BOLD; c.fill=fill_gold; c.border=box
ws["A15"]="Implied FY26E EV/Revenue at offer (pre-money)"; 
ws["C15"]="=(C13-Financials!F24)/Financials!F7"; ws["D15"]="=(D13-Financials!F24)/Financials!F7"
ws["A16"]="Implied FY26E EV/EBITDA at offer (pre-money)"
ws["C16"]="=(C13-Financials!F24)/Financials!F12"; ws["D16"]="=(D13-Financials!F24)/Financials!F12"
for r in (15,16):
    for cc in (3,4):
        ws.cell(row=r,column=cc).number_format='0.0"x"'; ws.cell(row=r,column=cc).border=box
ws["A18"]="Illustrative - fictional data"; ws["A18"].font=ITAL
widths(ws,{"A":42,"B":34,"C":14,"D":14,"E":14})

# ---------------------------------------------------------------- Offer structure
ws = wb.create_sheet("Offer_Structure")
title(ws, "OFFER STRUCTURE & USE OF PROCEEDS")
def orow(r,label,val,fmt="#,##0.0",inp=False,bold=False,unit=""):
    ws.cell(row=r,column=1,value=label).font=BOLD if bold else NORM
    c=ws.cell(row=r,column=2,value=val); c.number_format=fmt; c.border=box
    if inp: c.font=INP; c.fill=fill_inp
    if bold: c.font=Font(name="Calibri",size=11,bold=True); c.fill=fill_grey
    ws.cell(row=r,column=3,value=unit).font=ITAL
orow(4,"Pre-IPO shares outstanding",320.0,"#,##0.0",inp=True,unit="m shares")
orow(5,"New primary shares issued",51.0,"#,##0.0",inp=True,unit="m shares")
orow(6,"Secondary (selling shareholder) shares",22.0,"#,##0.0",inp=True,unit="m shares")
orow(7,"Base offering size","=B5+B6","#,##0.0",bold=True,unit="m shares")
orow(8,"Primary share of offering","=B5/B7","0%",unit="% of base offer")
orow(9,"Over-allotment option (greenshoe)","=B7*0.15","#,##0.0",unit="m shares (15%)")
orow(10,"Post-IPO shares outstanding","=B4+B5","#,##0.0",bold=True,unit="m shares")
orow(12,"Offer price - low","=Football_Field!C13/B4","$#,##0.00",unit="USD / share")
orow(13,"Offer price - high","=Football_Field!D13/B4","$#,##0.00",unit="USD / share")
orow(14,"Offer price - midpoint","=AVERAGE(B12:B13)","$#,##0.00",bold=True,unit="USD / share")
orow(16,"Gross proceeds - total (low)","=B7*B12","#,##0",unit="USD m")
orow(17,"Gross proceeds - total (high)","=B7*B13","#,##0",unit="USD m")
orow(18,"Primary gross proceeds (midpoint)","=B5*B14","#,##0",bold=True,unit="USD m - to the company")
orow(19,"Secondary gross proceeds (midpoint)","=B6*B14","#,##0",unit="USD m - to selling holders")
orow(20,"Underwriting & offering costs","=B18*0.05","#,##0",unit="USD m (5% of primary)")
orow(21,"Net primary proceeds (midpoint)","=B18-B20","#,##0",bold=True,unit="USD m")
orow(23,"Post-money market cap at midpoint","=B10*B14","#,##0",bold=True,unit="USD m")
orow(24,"Market cap at low / high","=B10*B12","#,##0",unit="USD m (low)")
orow(25,"","=B10*B13","#,##0",unit="USD m (high)")
orow(26,"Free float - base offering","=B7/B10","0.0%",bold=True,unit="% of post-IPO shares")
orow(27,"Free float - incl. greenshoe","=(B7+B9)/B10","0.0%",unit="% of post-IPO shares")
ws.cell(row=29,column=1,value="Use of primary proceeds").font=H2; ws.cell(row=29,column=1).fill=fill_navy
for i,h in enumerate(["Category","% of net proceeds","USD m"]):
    c=ws.cell(row=30,column=1+i,value=h); c.font=BOLD; c.fill=fill_lite; c.border=box
uop=[("Merchant acquisition & POS / QR terminal rollout",0.35),
     ("Technology platform, risk & fraud infrastructure",0.25),
     ("Licensing capital & regulatory reserve for new products",0.15),
     ("Strategic M&A and partnership investments",0.15),
     ("Working capital and general corporate purposes",0.10)]
r=31
for nm,p in uop:
    ws.cell(row=r,column=1,value=nm).font=NORM
    c=ws.cell(row=r,column=2,value=p); c.number_format="0%"; c.font=INP; c.fill=fill_inp; c.border=box
    c2=ws.cell(row=r,column=3,value=f"=$B$21*B{r}"); c2.number_format="#,##0.0"; c2.border=box
    r+=1
ws.cell(row=r,column=1,value="Total").font=BOLD
c=ws.cell(row=r,column=2,value=f"=SUM(B31:B{r-1})"); c.number_format="0%"; c.font=BOLD; c.fill=fill_grey; c.border=box
c=ws.cell(row=r,column=3,value=f"=SUM(C31:C{r-1})"); c.number_format="#,##0.0"; c.font=BOLD; c.fill=fill_grey; c.border=box
ws.cell(row=r+2,column=1,value="Illustrative - fictional data").font=ITAL
widths(ws,{"A":50,"B":16,"C":26})

# ---------------------------------------------------------------- Timeline
ws = wb.create_sheet("IPO_Timeline")
title(ws, "INDICATIVE IPO EXECUTION TIMELINE")
for i,h in enumerate(["Phase","Weeks from kick-off","Duration (weeks)","Key workstreams"]):
    c=ws.cell(row=4,column=1+i,value=h); c.font=H2; c.fill=fill_navy; c.border=box; c.alignment=Alignment(wrap_text=True,horizontal="center")
tl=[("1. Preparation & readiness","W0 - W10",10,"Financial and legal due diligence, audit of 3 years, corporate restructuring, ESOP clean-up, governance and independent board build-out"),
    ("2. Documentation & filing","W10 - W20",10,"Prospectus drafting, valuation report, listing application to the exchange per prevailing listing regulations"),
    ("3. Regulatory review","W20 - W30",10,"Responses to regulator comments, approval of the registration file, finalisation of the offering structure"),
    ("4. Marketing & roadshow","W30 - W34",4,"Analyst presentation, pre-deal investor education, management roadshow across domestic and regional accounts"),
    ("5. Bookbuilding & pricing","W34 - W36",2,"Order book build, price discovery within the range, allocation to anchor and institutional investors"),
    ("6. Listing & aftermarket","W36 - W40",4,"Settlement, listing on HOSE, stabilisation via greenshoe, index inclusion work, research coverage initiation")]
r=5
for ph,wk,d,ws_ in tl:
    ws.cell(row=r,column=1,value=ph).font=BOLD
    ws.cell(row=r,column=2,value=wk).font=NORM
    c=ws.cell(row=r,column=3,value=d); c.number_format="0"; c.font=INP; c.fill=fill_inp; c.border=box
    ws.cell(row=r,column=4,value=ws_).font=NORM
    ws.cell(row=r,column=4).alignment=Alignment(wrap_text=True,vertical="top")
    ws.row_dimensions[r].height=42
    r+=1
ws.cell(row=r,column=1,value="Total elapsed").font=BOLD
ws.cell(row=r,column=3,value="=SUM(C5:C10)").font=BOLD
ws.cell(row=r,column=4,value="weeks of sequential + overlapping workstreams; target listing ~9-10 months from mandate").font=ITAL
ws.cell(row=r+2,column=1,value="Illustrative - fictional data").font=ITAL
widths(ws,{"A":30,"B":20,"C":18,"D":72})

# ---------------------------------------------------------------- Slide numbers
ws = wb.create_sheet("Slide_Numbers")
title(ws, "SLIDE-TO-MODEL TRACE (every number used on a slide)")
for i,h in enumerate(["Slide","Figure as shown on slide","Source cell","Live value"]):
    c=ws.cell(row=4,column=1+i,value=h); c.font=H2; c.fill=fill_navy; c.border=box
trace=[
 (2,"FY25A revenue (USD m)","Financials!E7"),
 (2,"FY25A EBITDA (USD m)","Financials!E12"),
 (2,"Pre-money equity value low (USD m)","Football_Field!C13"),
 (2,"Pre-money equity value high (USD m)","Football_Field!D13"),
 (2,"Primary gross proceeds, midpoint (USD m)","Offer_Structure!B18"),
 (2,"Free float incl. greenshoe (%)","Offer_Structure!B27"),
 (4,"FY25A TPV (USD bn)","Operating_KPI!E5"),
 (5,"FY23A TPV (USD bn)","Operating_KPI!C5"),
 (5,"FY28E TPV (USD bn)","Operating_KPI!H5"),
 (5,"FY25A MAU (m)","Operating_KPI!E8"),
 (5,"FY28E MAU (m)","Operating_KPI!H8"),
 (5,"FY25A merchants ('000)","Operating_KPI!E9"),
 (5,"FY28E merchants ('000)","Operating_KPI!H9"),
 (5,"FY25A take rate (%)","Operating_KPI!E7"),
 (5,"FY28E take rate (%)","Operating_KPI!H7"),
 (6,"FY23A revenue (USD m)","Financials!C7"),
 (6,"FY28E revenue (USD m)","Financials!H7"),
 (6,"FY25A EBITDA margin (%)","Financials!E11"),
 (6,"FY28E EBITDA margin (%)","Financials!H11"),
 (6,"FY28E net income (USD m)","Financials!H19"),
 (7,"Comps EV/Rev low - high","Comps!G10 / Comps!G13"),
 (7,"Comps EV/EBITDA low - high","Comps!H10 / Comps!H13"),
 (7,"Precedents EV/Rev low - high","Precedents!H9 / Precedents!H12"),
 (7,"DCF equity low - high (USD m)","Football_Field!C8 / Football_Field!D8"),
 (7,"Concluded range (USD m)","Football_Field!C10 / Football_Field!D10"),
 (7,"IPO discount range (%)","Football_Field!D12 / Football_Field!C12"),
 (7,"Offer-price equity range (USD m)","Football_Field!C13 / Football_Field!D13"),
 (7,"Offer price per share (USD)","Offer_Structure!B12 / Offer_Structure!B13"),
 (8,"Base offering shares (m)","Offer_Structure!B7"),
 (8,"Primary % of offering","Offer_Structure!B8"),
 (8,"Net primary proceeds (USD m)","Offer_Structure!B21"),
 (8,"Secondary proceeds (USD m)","Offer_Structure!B19"),
 (8,"Post-money market cap (USD m)","Offer_Structure!B23"),
 (8,"Free float base / incl. shoe","Offer_Structure!B26 / Offer_Structure!B27"),
 (8,"Use of proceeds %","Offer_Structure!B31:B35"),
 (9,"Timeline phases / weeks","IPO_Timeline!A5:C10"),
 ("A1","Full comps table","Comps!A4:I13"),
 ("A2","Full precedents table","Precedents!A4:I12"),
 ("A3","DCF build & WACC","DCF!A4:J33"),
 ("A4","Sensitivity grid","Sensitivity!A6:F11"),
]
r=5
for s,f,src in trace:
    ws.cell(row=r,column=1,value=s).font=BOLD
    ws.cell(row=r,column=2,value=f).font=NORM
    ws.cell(row=r,column=3,value=src).font=Font(name="Consolas",size=10)
    if ":" not in src and "/" not in src:
        c=ws.cell(row=r,column=4,value=f"={src}"); c.border=box
    r+=1
ws.cell(row=r+1,column=1,value="Rule: no number may appear on a slide unless it is listed here.").font=ITAL
widths(ws,{"A":8,"B":42,"C":42,"D":16})

out="/home/user/HTML/models/NovaPay_IPO_Model_Illustrative.xlsx"
wb.save(out)
print("saved", out)

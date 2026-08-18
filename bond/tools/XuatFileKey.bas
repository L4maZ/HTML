Attribute VB_Name = "z_XuatFileKey"
Option Explicit

' ============================================================================
'  Sinh Key_YYYYMMDD.xlsx tu File 02.
'
'  CHI CO MOT MACRO DUY NHAT:  XuatFileKey
'  Lan chay dau tien no tu dat 17 Named Range cho o nhan dinh, cac lan sau
'  khong dat lai (de Excel tu doi vung khi o nhan dinh dich len xuong).
'  Muon dat lai: xoa name HTML_* trong Formulas > Name Manager roi chay lai.
'
'  Nguyen tac toc do: moi vung nguon doc MOT lan vao mang, moi sheet dich
'  ghi MOT lan bang mang. Khong doc/ghi tung o qua COM.
' ============================================================================

Private Const SCHEMA As String = "bond.key.v3"
Private Const SH_S2 As String = "Linked (1)"
Private Const SH_LK As String = "Linked"
Private Const SH_CD As String = "Chart data"
Private Const SH_RP As String = "Report"
Private Const SH_RT As String = "Run Tool"
Private Const SH_VS As String = "VIRA scenarios"
Private Const SH_VOL As String = "Volatility"
Private Const VOL_ROWS As Long = 800
Private Const NM_PREFIX As String = "HTML_"

Private Const S2_ROWS As Long = 160
Private Const S2_COLS As Long = 14
Private Const LK_ROWS As Long = 330
Private Const LK_COLS As Long = 195

' --- vung dem ghi ---
Private mB() As Variant
Private mW As Long
Private mN As Long
Private mCap As Long
Private mFull As Boolean

' --- nguon da doc san ---
Private gS2 As Variant
Private gLK As Variant

' Doc dinh dang chu: dem so lan hoi Excel, co tran an toan
Private gStep As String
Private mFmtCalls As Long
Private Const FMT_BUDGET As Long = 20000
Private Const FMT_MINSEG As Long = 3

' ============================ tien ich chuoi ============================

Private Function Txt(v As Variant) As String
    If IsObject(v) Then
        Txt = ""
    ElseIf IsError(v) Then
        Txt = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        Txt = ""
    Else
        Txt = Trim$(CStr(v))
    End If
End Function

Private Function OneLine(v As Variant) As String
    Dim s As String
    s = Txt(v)
    s = Replace$(s, vbCrLf, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbCr, " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    OneLine = Trim$(s)
End Function

Private Function KeepLines(v As Variant) As String
    Dim s As String, parts As Variant, i As Long, out As String, ln As String
    s = Txt(v)
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    parts = Split(s, vbLf)
    For i = LBound(parts) To UBound(parts)
        ln = parts(i)
        Do While InStr(ln, "  ") > 0
            ln = Replace$(ln, "  ", " ")
        Loop
        parts(i) = Trim$(ln)
    Next i
    out = Join(parts, vbLf)
    Do While Left$(out, 1) = vbLf
        out = Mid$(out, 2)
    Loop
    Do While Len(out) > 0 And Right$(out, 1) = vbLf
        out = Left$(out, Len(out) - 1)
    Loop
    KeepLines = out
End Function

Private Sub AddMap(m As Object, ByVal base As String, codes As Variant)
    Dim i As Long
    For i = LBound(codes) To UBound(codes)
        m(ChrW$(CLng(codes(i)))) = base
    Next i
End Sub

Private Function FoldMap() As Object
    Static m As Object
    If m Is Nothing Then
        Set m = CreateObject("Scripting.Dictionary")
    AddMap m, "a", Array(224, 225, 7841, 7843, 227, 226, 7847, 7845, 7853, 7849, 7851, 259, 7857, 7855, 7863, 7859, 7861)
    AddMap m, "e", Array(232, 233, 7865, 7867, 7869, 234, 7873, 7871, 7879, 7875, 7877)
    AddMap m, "i", Array(236, 237, 7883, 7881, 297)
    AddMap m, "o", Array(242, 243, 7885, 7887, 245, 244, 7891, 7889, 7897, 7893, 7895, 417, 7901, 7899, 7907, 7903, 7905)
    AddMap m, "u", Array(249, 250, 7909, 7911, 361, 432, 7915, 7913, 7921, 7917, 7919)
    AddMap m, "y", Array(7923, 253, 7925, 7927, 7929)
    AddMap m, "d", Array(273)
    End If
    Set FoldMap = m
End Function

Private Function Fold(ByVal s As String) As String
    Dim i As Long, ch As String, m As Object, out As String
    Set m = FoldMap()
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If m.Exists(ch) Then out = out & m(ch) Else out = out & ch
    Next i
    Fold = out
End Function

Private Function Slug(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, up As Boolean, first As Boolean
    s = Fold(s)
    first = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            If up And Not first Then ch = UCase$(ch)
            out = out & ch
            up = False
            first = False
        Else
            up = True
        End If
    Next i
    If out = "" Then out = "unnamed"
    Slug = out
End Function

' Ngay tu o doc bang .Value (con kieu Date)
Private Function VnDate(v As Variant) As String
    On Error Resume Next
    If IsObject(v) Then
        VnDate = ""
    ElseIf IsError(v) Then
        VnDate = ""
    ElseIf IsDate(v) Then
        VnDate = Format$(v, "dd/mm/yyyy")
    Else
        VnDate = Txt(v)
    End If
    On Error GoTo 0
End Function

' Ngay tu mang doc bang .Value2 (chi con so serial)
Private Function VnSerial(v As Variant) As String
    On Error Resume Next
    If IsObject(v) Then
        VnSerial = ""
    ElseIf IsError(v) Then
        VnSerial = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        VnSerial = ""
    ElseIf VarType(v) = vbDate Then
        VnSerial = Format$(v, "dd/mm/yyyy")
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then
            VnSerial = Format$(CDate(CDbl(v)), "dd/mm/yyyy")
        Else
            VnSerial = Txt(v)
        End If
    Else
        VnSerial = Txt(v)
    End If
    On Error GoTo 0
End Function

' ============================ tien ich mang ============================

' Doc an toan tu mang 2 chieu, ngoai vung thi tra ve rong
Private Function G(a As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If Not IsArray(a) Then Exit Function
    If r < LBound(a, 1) Or r > UBound(a, 1) Then Exit Function
    If c < LBound(a, 2) Or c > UBound(a, 2) Then Exit Function
    G = a(r, c)
End Function

Private Sub BufNew(ByVal w As Long, ByVal cap As Long)
    mW = w
    mCap = cap
    mN = 0
    mFull = False
    ReDim mB(1 To cap, 1 To w)
End Sub

' Mo mot dong moi trong vung dem. Tra ve False neu day.
Private Function BufRow() As Boolean
    If mN >= mCap Then
        mFull = True
        BufRow = False
        Exit Function
    End If
    mN = mN + 1
    BufRow = True
End Function

Private Sub BufSet(ByVal c As Long, v As Variant)
    If mN >= 1 And c >= 1 And c <= mW Then mB(mN, c) = v
End Sub

Private Sub BufFlush(sh As Worksheet)
    Dim o() As Variant, r As Long, c As Long
    If mN = 0 Then Exit Sub
    ReDim o(1 To mN, 1 To mW)
    For r = 1 To mN
        For c = 1 To mW
            o(r, c) = mB(r, c)
        Next c
    Next r
    sh.Range("A2").Resize(mN, mW).Value = o
End Sub

' Tim dong chua nhan trong mot cot cua mang nguon
Private Function FindRowA(a As Variant, ByVal col As Long, ByVal label As String, _
                          ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim r As Long, want As String
    want = Fold(Trim$(label))
    If r2 > UBound(a, 1) Then r2 = UBound(a, 1)
    For r = r1 To r2
        If Fold(Txt(a(r, col))) = want Then
            FindRowA = r
            Exit Function
        End If
    Next r
    FindRowA = 0
End Function

' ==================== Named Range cho o nhan dinh ====================

' Dat lai 17 Named Range. Binh thuong khong can goi: XuatFileKey tu goi khi
' chua co name. Muon dat lai thi xoa name trong Formulas > Name Manager,
' lan chay sau macro tu tao lai.
Private Sub TaoNameNhanDinh()
    DatName False
End Sub

Private Function NameKeys() As Variant
    NameKeys = Array("market", "assessment", "noteTB", "noteTPCP", "noteFI", "noteFV", _
                     "itdTB", "itdBB", "note10d", "noteVar", "noteBB10d", "noteBBVar", _
                     "noteRealized", "noteScenario", "noteSign", "noteRating", "noteVira")
End Function

Private Sub DatName(ByVal quiet As Boolean)
    Dim rp As Worksheet, i As Long, n As Long
    Dim keys As Variant, addrs As Variant
    keys = NameKeys()
    addrs = Array("B4", "H4", "Q10", "Q59", "Q66", "B110", _
                  "B84", "L84", "B166", "B184", "B203", "B221", _
                  "M95", "M110", "B120", "B254", "N223")
    Set rp = ThisWorkbook.Sheets(SH_RP)
    For i = LBound(keys) To UBound(keys)
        On Error Resume Next
        ThisWorkbook.Names(NM_PREFIX & keys(i)).Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:=NM_PREFIX & keys(i), _
                               RefersTo:="='" & SH_RP & "'!" & rp.Range(addrs(i)).Address
        n = n + 1
    Next i
    If Not quiet Then
        MsgBox "Da dat " & n & " Named Range cho o nhan dinh." & vbCrLf & _
               "Tu gio o co dich len xuong thi Excel tu cap nhat, macro van doc dung.", vbInformation
    End If
End Sub

Private Function NamesReady() As Boolean
    Dim nm As Object
    On Error Resume Next
    Set nm = ThisWorkbook.Names(NM_PREFIX & "market")
    On Error GoTo 0
    NamesReady = Not (nm Is Nothing)
End Function

' ================= doc dinh dang chu trong o nhan dinh =================
' Duong nhanh: hoi dinh dang CA O bang 3 loi goi. O nao dong nhat (phan lon)
' thi xong ngay. Chi o co nhieu dinh dang moi phai quet chia doi.

' Dung chuoi GOC (chua Trim) vi Characters(st, ln) dem theo chuoi goc.
' Bat ky truc trac nao khi doc dinh dang cung chi lam mat dinh dang,
' khong duoc lam chet ca macro.
Private Function RichMarkup(rg As Range) As String
    Dim s As String, fb As Variant, fi As Variant, fc As Variant
    Dim v As Variant
    On Error GoTo Plain
    v = rg.Value2
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        RichMarkup = ""
        Exit Function
    End If
    s = CStr(v)
    If Len(s) = 0 Then
        RichMarkup = ""
        Exit Function
    End If
    fb = rg.Font.Bold
    fi = rg.Font.Italic
    fc = rg.Font.Color
    If Not (IsNull(fb) Or IsNull(fi) Or IsNull(fc)) Then
        RichMarkup = AddFill(rg, Wrap(s, (fb = True), (fi = True), (fc <> 0)))
        Exit Function
    End If
    RichMarkup = AddFill(rg, Emit(rg, 1, Len(s), s))
    Exit Function
Plain:
    RichMarkup = s
End Function

Private Function Uniform(rg As Range, ByVal st As Long, ByVal ln As Long, _
                         ByRef bb As Boolean, ByRef ii As Boolean, ByRef rr As Boolean) As Boolean
    Dim f As Object
    On Error GoTo NotUniform
    mFmtCalls = mFmtCalls + 1
    Set f = rg.Characters(st, ln).Font
    If IsNull(f.Bold) Or IsNull(f.Italic) Or IsNull(f.Color) Then
        Uniform = False
    Else
        bb = (f.Bold = True)
        ii = (f.Italic = True)
        rr = (f.Color <> 0)
        Uniform = True
    End If
    Exit Function
NotUniform:
    Uniform = False
End Function

Private Function Wrap(ByVal t As String, ByVal bb As Boolean, ByVal ii As Boolean, _
                      ByVal rr As Boolean) As String
    If rr Then t = "[r]" & t & "[/r]"
    If bb Then t = "[b]" & t & "[/b]"
    If ii Then t = "[i]" & t & "[/i]"
    Wrap = t
End Function

Private Function Emit(rg As Range, ByVal st As Long, ByVal ln As Long, ByRef s As String) As String
    Dim bb As Boolean, ii As Boolean, rr As Boolean, half As Long
    If ln <= 0 Then
        Emit = ""
        Exit Function
    End If
    If Uniform(rg, st, ln, bb, ii, rr) Then
        Emit = Wrap(Mid$(s, st, ln), bb, ii, rr)
        Exit Function
    End If
    ' Doan qua ngan hoac da het ngan sach: lay dinh dang cua ky tu dau
    If ln <= FMT_MINSEG Or mFmtCalls > FMT_BUDGET Then
        If ln = 1 Then
            Emit = Mid$(s, st, 1)
        ElseIf Uniform(rg, st, 1, bb, ii, rr) Then
            Emit = Wrap(Mid$(s, st, ln), bb, ii, rr)
        Else
            Emit = Mid$(s, st, ln)
        End If
        Exit Function
    End If
    half = ln \ 2
    Emit = Emit(rg, st, half, s) & Emit(rg, st + half, ln - half, s)
End Function

Private Function AddFill(rg As Range, ByVal s As String) As String
    On Error Resume Next
    If rg.Interior.ColorIndex <> xlColorIndexNone And rg.Interior.Color <> 16777215 Then
        s = "[h]" & s & "[/h]"
    End If
    On Error GoTo 0
    AddFill = s
End Function

Private Function NameVal(ByVal k As String) As String
    Dim r As Range
    On Error Resume Next
    Set r = ThisWorkbook.Names(NM_PREFIX & k).RefersToRange
    On Error GoTo 0
    If r Is Nothing Then
        NameVal = ""
    Else
        NameVal = KeepLines(RichMarkup(r.Cells(1, 1)))
    End If
End Function

' ============================ xuat file key ============================

Public Sub XuatFileKey()
    Dim s2 As Worksheet, lk As Worksheet, cd As Worksheet, rt As Worksheet
    Dim vs As Worksheet, vo As Worksheet
    Dim wb As Workbook, sh As Worksheet
    Dim i As Long, r As Long, c As Long
    Dim dest As String, t0 As Single, tRead As Single, tBody As Single, tText As Single
    Dim oCalc As Long, oEvt As Boolean, oUpd As Boolean, oAlert As Boolean
    Dim started As Boolean, msgLog As String, missLab As String

    t0 = Timer
    On Error GoTo Fail

    gStep = "mo cac sheet nguon"
    Set s2 = ThisWorkbook.Sheets(SH_S2)
    Set lk = ThisWorkbook.Sheets(SH_LK)
    Set cd = ThisWorkbook.Sheets(SH_CD)
    Set rt = ThisWorkbook.Sheets(SH_RT)
    Set vs = ThisWorkbook.Sheets(SH_VS)
    Set vo = ThisWorkbook.Sheets(SH_VOL)

    gStep = "dat Named Range cho o nhan dinh"
    If Not NamesReady() Then DatName True

    oCalc = Application.Calculation
    oEvt = Application.EnableEvents
    oUpd = Application.ScreenUpdating
    oAlert = Application.DisplayAlerts
    started = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "File key 1/6 - doc du lieu nguon..."

    ' ---- doc mot lan toan bo vung nguon ----
    gStep = "doc sheet Linked (1)"
    gS2 = s2.Range(s2.Cells(1, 1), s2.Cells(S2_ROWS, S2_COLS)).Value2
    gStep = "doc sheet Linked"
    gLK = lk.Range(lk.Cells(1, 1), lk.Cells(LK_ROWS, LK_COLS)).Value2
    tRead = Timer - t0

    gStep = "tao workbook tam"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Application.StatusBar = "File key 2/6 - khoi 2 va cac bang so..."

    ' ---------------- META ----------------
    gStep = "sheet META"
    Set sh = wb.Sheets(1)
    sh.Name = "META"
    sh.Range("A1:C1").Value = Array("Key", "Value", "Ghi chu")
    Dim mk As Variant, mv As Variant
    mk = Array("schema", "asOf", "dateYest", "dateLastMonth", "dateLastQuarter", _
               "dateLastYear", "generatedAt")
    mv = Array(SCHEMA, VnDate(rt.Range("B2").Value), VnDate(rt.Range("B3").Value), _
               VnDate(rt.Range("B4").Value), VnDate(s2.Range("H3").Value), _
               VnDate(rt.Range("B7").Value), Format$(Now, "dd/mm/yyyy hh:nn"))
    BufNew 3, 20
    For i = LBound(mk) To UBound(mk)
        If Not BufRow() Then GoTo Overflow
        BufSet 1, mk(i)
        BufSet 2, mv(i)
    Next i
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- DATA_S2 : do theo nhan khoi ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "DATA_S2"
    sh.Range("A1:Q1").Value = Array("KeyID", "Book", "Code", "Section", "STT", "Nhom", _
        "ChiTieu", "Sub", "Today", "DtD", "Yesterday", "LastMonth", "LastQuarter", _
        "LastYear", "Limit", "Used", "Light")

    gStep = "DATA_S2 - do nhan khoi"
    Dim bcode As Variant, blab As Variant, bkey As Variant
    bkey = Array("TB_INT", "BB_INT", "TB_SBV", "BB_SBV", "OTHER", "FIBOND")
    bcode = Array("2.1.", "2.2.", "2.3.", "2.4.", "2.5.", "2.6.")
    blab = Array("trading book noi bo", "banking book noi bo", "trading book sbv", _
                 "banking book sbv", "khac", "fi bond & cd")

    Dim starts(0 To 5) As Long, ends(0 To 5) As Long
    For i = 0 To 5
        starts(i) = FindRowA(gS2, 2, CStr(blab(i)), 1, 130)
        If starts(i) = 0 Then
            missLab = CStr(blab(i))
            wb.Close False
            Set wb = Nothing
            GoTo NoLabel
        End If
    Next i
    For i = 0 To 5
        If i < 5 Then
            ends(i) = starts(i + 1) - 1
        Else
            ends(i) = starts(i) + 30
        End If
        If ends(i) > S2_ROWS Then ends(i) = S2_ROWS
    Next i

    Dim dict As Object, lab As String, base As String, kk As String
    gStep = "DATA_S2 - doc 52 chi tieu"
    Set dict = CreateObject("Scripting.Dictionary")
    BufNew 17, 600
    For i = 0 To 5
        For r = starts(i) + 1 To ends(i)
            lab = OneLine(G(gS2, r, 3))
            If lab <> "" Then
                base = Slug(lab)
                kk = bkey(i) & "|" & base
                If dict.Exists(kk) Then
                    dict(kk) = dict(kk) + 1
                    base = base & CStr(dict(kk))
                Else
                    dict.Add kk, 1
                End If
                If Not BufRow() Then GoTo Overflow
                BufSet 1, "S2." & bkey(i) & "." & base
                BufSet 2, bkey(i)
                BufSet 3, bcode(i)
                BufSet 4, OneLine(G(gS2, starts(i), 2))
                BufSet 5, G(gS2, r, 1)
                BufSet 6, OneLine(G(gS2, r, 2))
                BufSet 7, lab
                BufSet 8, IIf(Left$(lab, 2) = "a)" Or Left$(lab, 2) = "b)", 1, 0)
                For c = 4 To 12
                    BufSet 5 + c, G(gS2, r, c)
                Next c
            End If
        Next r
    Next i
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- POS ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "POS"
    sh.Range("A1:O1").Value = Array("KeyID", "Book", "Tenor", "Face", "FaceYest", "FaceLM", _
        "DtD", "MtD", "PV01", "Itd", "ItdYest", "ItdLM", "ItdMtD", "ItdYtD", "Daily")
    BufNew 15, 300
    gStep = "POS - Trading"
    DumpGrid "POS.TB.", "TB", 14, 65, 75, 13
    gStep = "POS - Banking"
    DumpGrid "POS.BB.", "BB", 28, 80, 90, 13
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- CURVE ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "CURVE"
    sh.Range("A1:J1").Value = Array("KeyID", "Type", "Tenor", "V1", "V2", "V3", "V4", _
        "DtD", "MtD", "YtD")
    BufNew 10, 300
    gStep = "CURVE - yield"
    DumpCurve "YIELD", 41, 80, 89
    gStep = "CURVE - repo"
    DumpCurve "REPO", 95, 171, 178
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- GRID ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "GRID"
    sh.Range("A1:R1").Value = Array("KeyID", "Block", "BlockName", "Label", _
        "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14")
    BufNew 18, 2000
    gStep = "GRID - MIX"
    DumpBlock "MIX", "3.3 Co cau theo to chuc phat hanh", 147, 238, 249, 7
    gStep = "GRID - HOLD"
    DumpBlock "HOLD", "3.4 Co cau theo thoi gian nam giu", 164, 274, 284, 5
    gStep = "GRID - PNL"
    DumpBlock "PNL", "3.5 Unrealized & Realized PnL", 154, 253, 259, 4
    gStep = "GRID - CAPITAL"
    DumpBlock "CAPITAL", "3.8 Muc do su dung von", 158, 262, 270, 6
    gStep = "GRID - BS"
    DumpBlock "BS", "3.6 Ghi nhan PnL theo lop bang can doi", 103, 184, 189, 14
    gStep = "GRID - FIONBS"
    DumpBlock "FIONBS", "7.1 FI Bond tren bang can doi", 169, 288, 294, 5
    gStep = "GRID - PNLSCEN"
    DumpBlock "PNLSCEN", "3.7 Phan tich kich ban PnL", 140, 230, 234, 7
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- SCEN ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "SCEN"
    sh.Range("A1:J1").Value = Array("KeyID", "Book", "Kind", "Scenario", "Sub", "Tenor", _
        "PV01", "Itd", "YieldBps", "ItdChange")
    BufNew 10, 1200
    gStep = "SCEN - TB recent"
    DumpScen "TB", "recent", 51, 95, 105, 92
    gStep = "SCEN - TB var"
    DumpScen "TB", "var", 51, 110, 120, 107
    gStep = "SCEN - BB recent"
    DumpScen "BB", "recent", 51, 140, 150, 137
    gStep = "SCEN - BB var"
    DumpScen "BB", "var", 51, 125, 135, 122
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- VIRA4 ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "VIRA4"
    sh.Range("A1:F1").Value = Array("Month", "VIRA", "Big4", "MarketMaker", "Top3", "Actual")
    Dim av As Variant
    gStep = "VIRA4 - doc VIRA scenarios"
    av = vs.Range(vs.Cells(1, 1), vs.Cells(40, 6)).Value2
    BufNew 6, 60
    For r = 2 To UBound(av, 1)
        If Txt(av(r, 1)) <> "" And Txt(av(r, 2)) <> "" Then
            If Not BufRow() Then GoTo Overflow
            BufSet 1, VnSerial(av(r, 1))
            For c = 2 To 6
                BufSet c, av(r, c)
            Next c
        End If
    Next r
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- VIRA ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "VIRA"
    sh.Range("A1:H1").Value = Array("KeyID", "Scope", "Label", "PV01", "Itd", _
        "Scenario", "YieldBps", "ItdAfter")
    BufNew 8, 600
    gStep = "VIRA - theo book"
    DumpVira "BOOK", 66, 154, 155, 151
    gStep = "VIRA - theo ky han"
    DumpVira "BB_TENOR", 78, 154, 164, 151
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- RATING ----------------
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "RATING"
    sh.Range("A1:J1").Value = Array("KeyID", "Issuer", "Amount", "Rating", "ReviewDate", _
        "Pct", "CumPct", "Fitch", "Moody", "SP")
    BufNew 10, 200
    gStep = "RATING"
    Dim iss As String
    For r = 203 To 223
        iss = OneLine(G(gLK, r, 131))
        If iss <> "" Then
            If Not BufRow() Then GoTo Overflow
            BufSet 1, "RATING." & Slug(iss)
            BufSet 2, iss
            For c = 1 To 8
                If c = 3 Then
                    BufSet 2 + c, VnSerial(G(gLK, r, 131 + c))
                Else
                    BufSet 2 + c, G(gLK, r, 131 + c)
                End If
            Next c
        End If
    Next r
    If mFull Then GoTo Overflow
    BufFlush sh
    tBody = Timer - t0

    ' ---------------- VOL ----------------
    Application.StatusBar = "File key 3/6 - Volatility..."
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "VOL"
    Dim va As Variant
    gStep = "VOL - doc Volatility"
    va = vo.Range(vo.Cells(1, 1), vo.Cells(VOL_ROWS + 1, 11)).Value2
    BufNew 11, VOL_ROWS + 10
    ' hang tieu de ghi rieng vi BufFlush bat dau tu A2
    Dim vh(1 To 11) As Variant
    vh(1) = "Date"
    For c = 1 To 10
        vh(c + 1) = OneLine(va(1, c + 1))
    Next c
    sh.Range("A1").Resize(1, 11).Value = vh
    For r = 2 To UBound(va, 1)
        If Txt(va(r, 1)) = "" Then Exit For
        If Not BufRow() Then GoTo Overflow
        BufSet 1, VnSerial(va(r, 1))
        For c = 1 To 10
            BufSet c + 1, va(r, c + 1)
        Next c
    Next r
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- TEXT ----------------
    Application.StatusBar = "File key 4/6 - nhan dinh va ghi chu..."
    Dim tText0 As Single
    tText0 = Timer
    gStep = "TEXT - doc nhan dinh"
    mFmtCalls = 0
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "TEXT"
    sh.Range("A1:D1").Value = Array("KeyID", "Mo ta", "Nguon", "Value")
    Dim tk As Variant, td As Variant
    tk = NameKeys()
    td = Array("Thong tin thi truong", "Tuan thu han muc & danh gia", "Ghi chu Trading Book", _
               "Ghi chu ty le TPCP", "Ghi chu co cau FI Bond", "Ghi chu Fair value", _
               "Lo MtM Trading theo ky han", "Lo MtM Banking theo ky han", _
               "Kich ban 10 ngay Trading", "Kich ban VaR Trading", _
               "Kich ban 1 thang Banking", "Kich ban VaR Banking", _
               "Ghi chu Realized PnL", "Ghi chu kich ban PnL", "Quy uoc dau", _
               "Ghi chu xep hang FI Bond", "Ghi chu VIRA")
    BufNew 4, 40
    For i = LBound(tk) To UBound(tk)
        If Not BufRow() Then GoTo Overflow
        BufSet 1, "txt." & tk(i)
        BufSet 2, td(i)
        BufSet 3, NM_PREFIX & tk(i)
        gStep = "TEXT - o nhan dinh " & NM_PREFIX & tk(i)
        BufSet 4, NameVal(CStr(tk(i)))
    Next i
    If mFull Then GoTo Overflow
    BufFlush sh
    sh.Columns(4).WrapText = True
    tText = Timer - tText0

    ' ---------------- TS ----------------
    Application.StatusBar = "File key 5/6 - chuoi bieu do..."
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    sh.Name = "TS"
    sh.Range("A1:D1").Value = Array("Series", "Field", "Date", "Value")
    gStep = "TS - chuoi bieu do"
    DumpTS cd
    If mFull Then GoTo Overflow
    BufFlush sh

    ' ---------------- luu ----------------
    Application.StatusBar = "File key 6/6 - dang luu..."
    dest = ThisWorkbook.Path & Application.PathSeparator & _
           "Key_" & Format$(rt.Range("B2").Value, "yyyymmdd") & ".xlsx"
    gStep = "luu file key"
    wb.SaveAs Filename:=dest, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Nothing

    msgLog = "Doc nguon " & Format$(tRead, "0.0") & "s" & _
          "  |  bang so " & Format$(tBody - tRead, "0.0") & "s" & _
          "  |  nhan dinh " & Format$(tText, "0.0") & "s (" & mFmtCalls & " luot doc dinh dang)" & _
          "  |  tong " & Format$(Timer - t0, "0.0") & "s"

    RestoreApp started, oCalc, oEvt, oUpd, oAlert
    MsgBox "Da xuat file key:" & vbCrLf & dest & vbCrLf & vbCrLf & msgLog, vbInformation
    Exit Sub

Overflow:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    RestoreApp started, oCalc, oEvt, oUpd, oAlert
    MsgBox "Du lieu vuot suc chua vung dem cua macro." & vbCrLf & _
           "Bao lai cho nguoi viet macro de nang han muc.", vbCritical
    Exit Sub

NoLabel:
    RestoreApp started, oCalc, oEvt, oUpd, oAlert
    MsgBox "Khong tim thay nhan khoi: " & missLab & vbCrLf & _
           "Tim trong cot B cua sheet " & SH_S2 & ", 130 dong dau." & vbCrLf & _
           "Kiem tra lai chu cua dong do. Macro dung, khong xuat file sai.", vbCritical
    Exit Sub

Fail:
    Dim em As String
    em = "Loi " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    RestoreApp started, oCalc, oEvt, oUpd, oAlert
    MsgBox "Macro dung o buoc:" & vbCrLf & "   " & gStep & vbCrLf & vbCrLf & em & vbCrLf & vbCrLf & _
           "Chup nguyen hop thoai nay gui lai.", vbCritical
End Sub

Private Sub RestoreApp(ByVal started As Boolean, ByVal oCalc As Long, _
                       ByVal oEvt As Boolean, ByVal oUpd As Boolean, ByVal oAlert As Boolean)
    On Error Resume Next
    If started Then
        Application.Calculation = oCalc
        Application.EnableEvents = oEvt
        Application.ScreenUpdating = oUpd
        Application.DisplayAlerts = oAlert
    End If
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ======================= cac ham do khoi (doc tu mang) =======================

Private Sub DumpGrid(ByVal pfx As String, ByVal book As String, ByVal c0 As Long, _
                     ByVal r0 As Long, ByVal r1 As Long, ByVal w As Long)
    Dim r As Long, c As Long, t As String
    For r = r0 To r1
        t = OneLine(G(gLK, r, c0))
        If t <> "" Then
            If Not BufRow() Then Exit Sub
            BufSet 1, pfx & Slug(t)
            BufSet 2, book
            BufSet 3, t
            For c = 1 To w - 1
                BufSet 3 + c, G(gLK, r, c0 + c)
            Next c
        End If
    Next r
End Sub

Private Sub DumpCurve(ByVal nm As String, ByVal c0 As Long, _
                      ByVal r0 As Long, ByVal r1 As Long)
    Dim r As Long, t As String
    For r = r0 To r1
        t = OneLine(G(gLK, r, c0))
        If t <> "" Then
            If Not BufRow() Then Exit Sub
            BufSet 1, "CURVE." & nm & "." & Slug(t)
            BufSet 2, nm
            BufSet 3, t
            BufSet 4, G(gLK, r, c0 + 1)
            BufSet 5, G(gLK, r, c0 + 2)
            BufSet 6, G(gLK, r, c0 + 3)
            BufSet 7, G(gLK, r, c0 + 4)
            BufSet 8, G(gLK, r, c0 + 6)
            BufSet 9, G(gLK, r, c0 + 7)
            BufSet 10, G(gLK, r, c0 + 8)
        End If
    Next r
End Sub

Private Sub DumpBlock(ByVal nm As String, ByVal title As String, ByVal c0 As Long, _
                      ByVal r0 As Long, ByVal r1 As Long, ByVal w As Long)
    Dim r As Long, c As Long, t As String
    For r = r0 To r1
        t = OneLine(G(gLK, r, c0))
        If t <> "" Then
            If Not BufRow() Then Exit Sub
            BufSet 1, "GRID." & nm & "." & Slug(t)
            BufSet 2, nm
            BufSet 3, title
            BufSet 4, t
            For c = 1 To w - 1
                BufSet 4 + c, G(gLK, r, c0 + c)
            Next c
        End If
    Next r
End Sub

Private Sub DumpScen(ByVal book As String, ByVal kind As String, ByVal c0 As Long, _
                     ByVal r0 As Long, ByVal r1 As Long, ByVal hdr As Long)
    Dim r As Long, i As Long, tenor As String
    Dim nm(0 To 5) As String, sb(0 To 5) As String
    For i = 0 To 5
        nm(i) = OneLine(G(gLK, hdr, c0 + 3 + i * 2))
        sb(i) = OneLine(G(gLK, hdr + 2, c0 + 3 + i * 2))
        If nm(i) = "" Then nm(i) = "KB" & CStr(i + 1)
    Next i
    For r = r0 To r1
        tenor = OneLine(G(gLK, r, c0))
        If tenor <> "" Then
            For i = 0 To 5
                If Not BufRow() Then Exit Sub
                BufSet 1, "SCEN." & book & "." & kind & "." & Slug(nm(i)) & "." & Slug(tenor)
                BufSet 2, book
                BufSet 3, kind
                BufSet 4, nm(i)
                BufSet 5, sb(i)
                BufSet 6, tenor
                BufSet 7, G(gLK, r, c0 + 1)
                BufSet 8, G(gLK, r, c0 + 2)
                BufSet 9, G(gLK, r, c0 + 3 + i * 2)
                BufSet 10, G(gLK, r, c0 + 4 + i * 2)
            Next i
        End If
    Next r
End Sub

Private Sub DumpVira(ByVal scope As String, ByVal c0 As Long, ByVal r0 As Long, _
                     ByVal r1 As Long, ByVal hdr As Long)
    Dim r As Long, i As Long, lab As String
    Dim nm(0 To 3) As String
    For i = 0 To 3
        nm(i) = OneLine(G(gLK, hdr, c0 + 3 + i * 2))
        If nm(i) = "" Then nm(i) = "VIRA" & CStr(i + 1)
    Next i
    For r = r0 To r1
        lab = OneLine(G(gLK, r, c0))
        If lab <> "" Then
            For i = 0 To 3
                If Not BufRow() Then Exit Sub
                BufSet 1, "VIRA." & LCase$(scope) & "." & Slug(lab) & "." & Slug(nm(i))
                BufSet 2, scope
                BufSet 3, lab
                BufSet 4, G(gLK, r, c0 + 1)
                BufSet 5, G(gLK, r, c0 + 2)
                BufSet 6, nm(i)
                BufSet 7, G(gLK, r, c0 + 3 + i * 2)
                BufSet 8, G(gLK, r, c0 + 4 + i * 2)
            Next i
        End If
    Next r
End Sub

' TS: tieu de nam ca o hang 1 cua Chart data. Doc mot lan, ghi mot lan.
Private Sub DumpTS(cd As Worksheet)
    Dim maxc As Long, maxr As Long, c As Long, i As Long, r As Long
    Dim heads() As String, isAxis() As Boolean, src As Variant
    Dim serName As String, axisCol As Long, endCol As Long, dtxt As String

    maxc = cd.Cells(1, cd.Columns.Count).End(xlToLeft).Column
    ' Moi chuoi mot do dai khac nhau, cot B chi la truc cua khoi dau tien.
    ' Phai lay dong cuoi cua CA sheet, neu khong cac chuoi dai bi cat cut.
    maxr = 0
    On Error Resume Next
    maxr = cd.UsedRange.Row + cd.UsedRange.Rows.Count - 1
    On Error GoTo 0
    If maxr < 2 Then maxr = 2
    If maxc < 2 Then Exit Sub
    src = cd.Range(cd.Cells(1, 1), cd.Cells(maxr, maxc)).Value2

    ReDim heads(1 To maxc)
    ReDim isAxis(1 To maxc)
    For c = 1 To maxc
        heads(c) = OneLine(src(1, c))
        Select Case Fold(heads(c))
            Case "date", "ngay", "thang", "ky han", "stt": isAxis(c) = True
        End Select
    Next c

    BufNew 4, 60000

    For c = 1 To maxc - 1
        If heads(c) <> "" And Not isAxis(c) And isAxis(c + 1) Then
            serName = Slug(heads(c))
            axisCol = c + 1
            endCol = maxc
            For i = c + 2 To maxc - 1
                If heads(i) <> "" And isAxis(i + 1) Then
                    endCol = i - 1
                    Exit For
                End If
            Next i
            For r = 2 To maxr
                If Txt(src(r, axisCol)) = "" Then Exit For
                dtxt = VnSerial(src(r, axisCol))
                For i = axisCol + 1 To endCol
                    If heads(i) <> "" And Not isAxis(i) Then
                        If Txt(src(r, i)) <> "" Then
                            If Not BufRow() Then Exit Sub
                            BufSet 1, serName
                            BufSet 2, heads(i)
                            BufSet 3, dtxt
                            BufSet 4, src(r, i)
                        End If
                    End If
                Next i
            Next r
        End If
    Next c
End Sub

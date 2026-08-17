Attribute VB_Name = "z_XuatFileKey"
Option Explicit

' Sinh Key_YYYYMMDD.xlsx tu File 02, cung cau truc voi ban Python.
' Chay mot lan:  TaoNameNhanDinh   (dat Named Range cho 17 o nhan dinh)
' Chay moi ky:   XuatFileKey

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

' ---------- tien ich ----------

Private Function Txt(v As Variant) As String
    If IsError(v) Then
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
    Dim s As String, parts As Variant, i As Long, out As String
    s = Txt(v)
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    parts = Split(s, vbLf)
    For i = LBound(parts) To UBound(parts)
        Dim ln As String
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
    Do While Right$(out, 1) = vbLf
        out = Left$(out, Len(out) - 1)
    Loop
    KeepLines = out
End Function

Private Function Fold(ByVal s As String) As String
    Dim src As Variant, dst As Variant, i As Long, j As Long
    src = Array("àáạảãâầấậẩẫăằắặẳẵ", "èéẹẻẽêềếệểễ", "ìíịỉĩ", _
                "òóọỏõôồốộổỗơờớợởỡ", "ùúụủũưừứựửữ", "ỳýỵỷỹ", "đ")
    dst = Array("a", "e", "i", "o", "u", "y", "d")
    s = LCase$(s)
    For i = LBound(src) To UBound(src)
        For j = 1 To Len(src(i))
            s = Replace$(s, Mid$(src(i), j, 1), dst(i))
        Next j
    Next i
    Fold = s
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

Private Function VnDate(v As Variant) As String
    If IsDate(v) Then VnDate = Format$(v, "dd/mm/yyyy") Else VnDate = Txt(v)
End Function

' Tim dong chua nhan trong mot cot, trong pham vi r1..r2
Private Function FindRow(ws As Worksheet, col As Long, ByVal label As String, _
                         ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim r As Long, want As String
    want = Fold(Trim$(label))
    For r = r1 To r2
        If Fold(Txt(ws.Cells(r, col).Value)) = want Then
            FindRow = r
            Exit Function
        End If
    Next r
    FindRow = 0
End Function

' ---------- dat Named Range cho o nhan dinh ----------

Public Sub TaoNameNhanDinh()
    Dim rp As Worksheet, i As Long, n As Long
    Dim keys As Variant, addrs As Variant
    keys = Array("market", "assessment", "noteTB", "noteTPCP", "noteFI", "noteFV", _
                 "itdTB", "itdBB", "note10d", "noteVar", "noteBB10d", "noteBBVar", _
                 "noteRealized", "noteScenario", "noteSign", "noteRating", "noteVira")
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
    MsgBox "Da dat " & n & " Named Range cho o nhan dinh." & vbCrLf & _
           "Tu gio o co dich len xuong thi Excel tu cap nhat, macro van doc dung.", vbInformation
End Sub

' Doc rich text runs cua mot o -> danh dau [b] [i] [r] [h]
Private Function RichMarkup(rg As Range) As String
    Dim s As String, i As Long, n As Long
    Dim curB As Boolean, curI As Boolean, curR As Boolean
    Dim nB As Boolean, nI As Boolean, nR As Boolean
    Dim out As String, mixed As Boolean

    s = Txt(rg.Value)
    If s = "" Then RichMarkup = "": Exit Function
    n = Len(s)

    mixed = IsNull(rg.Font.Bold) Or IsNull(rg.Font.Italic) Or IsNull(rg.Font.Color)
    If Not mixed Then
        out = s
        If rg.Font.Color <> 0 And Not IsNull(rg.Font.Color) Then out = "[r]" & out & "[/r]"
        If rg.Font.Bold = True Then out = "[b]" & out & "[/b]"
        If rg.Font.Italic = True Then out = "[i]" & out & "[/i]"
        RichMarkup = AddFill(rg, out)
        Exit Function
    End If

    curB = False: curI = False: curR = False
    For i = 1 To n
        With rg.Characters(i, 1).Font
            nB = (.Bold = True)
            nI = (.Italic = True)
            nR = (.Color <> 0)
        End With
        If nB <> curB Or nI <> curI Or nR <> curR Then
            If curI Then out = out & "[/i]"
            If curB Then out = out & "[/b]"
            If curR Then out = out & "[/r]"
            If nR Then out = out & "[r]"
            If nB Then out = out & "[b]"
            If nI Then out = out & "[i]"
            curB = nB: curI = nI: curR = nR
        End If
        out = out & Mid$(s, i, 1)
    Next i
    If curI Then out = out & "[/i]"
    If curB Then out = out & "[/b]"
    If curR Then out = out & "[/r]"

    RichMarkup = AddFill(rg, out)
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

' ---------- xuat file key ----------

Public Sub XuatFileKey()
    Dim s2 As Worksheet, lk As Worksheet, cd As Worksheet, rt As Worksheet
    Dim wb As Workbook, sh As Worksheet
    Dim i As Long, r As Long, c As Long, n As Long
    Dim dest As String, t0 As Single

    t0 = Timer
    Set s2 = ThisWorkbook.Sheets(SH_S2)
    Set lk = ThisWorkbook.Sheets(SH_LK)
    Set cd = ThisWorkbook.Sheets(SH_CD)
    Set rt = ThisWorkbook.Sheets(SH_RT)

    If ThisWorkbook.Names.Count = 0 Or NameVal("market") = "" Then
        If MsgBox("Chua thay Named Range nhan dinh. Chay TaoNameNhanDinh truoc?" & vbCrLf & _
                  "Chon No de van xuat (phan nhan dinh se rong).", vbYesNo + vbQuestion) = vbYes Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)

    ' --- META ---
    Set sh = wb.Sheets(1): sh.Name = "META"
    sh.Range("A1:C1").Value = Array("Key", "Value", "Ghi chu")
    Dim mk As Variant, mv As Variant
    mk = Array("schema", "asOf", "dateYest", "dateLastMonth", "dateLastQuarter", _
               "dateLastYear", "rptTitle", "rptSubtitle", "unitNote", "generatedAt")
    mv = Array(SCHEMA, VnDate(rt.Range("B2").Value), VnDate(rt.Range("B3").Value), _
               VnDate(rt.Range("B4").Value), VnDate(s2.Range("H3").Value), _
               VnDate(rt.Range("B7").Value), "Báo cáo rủi ro thị trường", _
               "Báo cáo Desk Bond", "Đơn vị: tỷ VND, trừ khi ghi khác. Giá trị âm là lỗ.", _
               Format$(Now, "dd/mm/yyyy hh:nn"))
    For i = LBound(mk) To UBound(mk)
        sh.Cells(i + 2, 1).Value = mk(i)
        sh.Cells(i + 2, 2).Value = mv(i)
    Next i

    ' --- DATA_S2: do theo nhan khoi ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "DATA_S2"
    sh.Range("A1:Q1").Value = Array("KeyID", "Book", "Code", "Section", "STT", "Nhom", _
        "ChiTieu", "Sub", "Today", "DtD", "Yesterday", "LastMonth", "LastQuarter", _
        "LastYear", "Limit", "Used", "Light")
    Dim bcode As Variant, blab As Variant, bkey As Variant
    bkey = Array("TB_INT", "BB_INT", "TB_SBV", "BB_SBV", "OTHER", "FIBOND")
    bcode = Array("2.1.", "2.2.", "2.3.", "2.4.", "2.5.", "2.6.")
    blab = Array("TRADING BOOK NỘI BỘ", "BANKING BOOK NỘI BỘ", "TRADING BOOK SBV", _
                 "BANKING BOOK SBV", "KHÁC", "FI Bond & CD")

    Dim starts(0 To 5) As Long, ends(0 To 5) As Long
    For i = 0 To 5
        starts(i) = FindRow(s2, 2, CStr(blab(i)), 1, 120)
        If starts(i) = 0 Then
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            wb.Close False
            MsgBox "Khong tim thay nhan khoi: " & blab(i) & vbCrLf & _
                   "Kiem tra lai chu trong cot B cua sheet " & SH_S2 & ".", vbCritical
            Exit Sub
        End If
    Next i
    For i = 0 To 5
        If i < 5 Then ends(i) = starts(i + 1) - 1 Else ends(i) = starts(i) + 30
    Next i

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    n = 2
    For i = 0 To 5
        For r = starts(i) + 1 To ends(i)
            Dim lab As String
            lab = OneLine(s2.Cells(r, 3).Value)
            If lab <> "" Then
                Dim base As String, kk As String
                base = Slug(lab)
                kk = bkey(i) & "|" & base
                If dict.Exists(kk) Then
                    dict(kk) = dict(kk) + 1
                    base = base & CStr(dict(kk))
                Else
                    dict.Add kk, 1
                End If
                sh.Cells(n, 1).Value = "S2." & bkey(i) & "." & base
                sh.Cells(n, 2).Value = bkey(i)
                sh.Cells(n, 3).Value = bcode(i)
                sh.Cells(n, 4).Value = blab(i)
                sh.Cells(n, 5).Value = s2.Cells(r, 1).Value
                sh.Cells(n, 6).Value = OneLine(s2.Cells(r, 2).Value)
                sh.Cells(n, 7).Value = lab
                sh.Cells(n, 8).Value = IIf(Left$(lab, 2) = "a)" Or Left$(lab, 2) = "b)", 1, 0)
                For c = 4 To 12
                    sh.Cells(n, 5 + c).Value = s2.Cells(r, c).Value
                Next c
                n = n + 1
            End If
        Next r
    Next i

    ' --- POS ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "POS"
    sh.Range("A1:O1").Value = Array("KeyID", "Book", "Tenor", "Face", "FaceYest", "FaceLM", _
        "DtD", "MtD", "PV01", "Itd", "ItdYest", "ItdLM", "ItdMtD", "ItdYtD", "Daily")
    n = 2
    n = DumpGrid(lk, sh, n, "POS.TB.", "TB", 14, 65, 75, 13)
    n = DumpGrid(lk, sh, n, "POS.BB.", "BB", 28, 80, 90, 13)

    ' --- CURVE ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "CURVE"
    sh.Range("A1:J1").Value = Array("KeyID", "Type", "Tenor", "V1", "V2", "V3", "V4", _
        "DtD", "MtD", "YtD")
    n = 2
    n = DumpCurve(lk, sh, n, "YIELD", 41, 80, 89)
    n = DumpCurve(lk, sh, n, "REPO", 95, 171, 178)

    ' --- GRID ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "GRID"
    sh.Range("A1:R1").Value = Array("KeyID", "Block", "BlockName", "Label", _
        "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14")
    n = 2
    n = DumpBlock(lk, sh, n, "MIX", "3.3 Cơ cấu theo tổ chức phát hành", 147, 238, 249, 7)
    n = DumpBlock(lk, sh, n, "HOLD", "3.4 Cơ cấu theo thời gian nắm giữ", 164, 274, 284, 5)
    n = DumpBlock(lk, sh, n, "PNL", "3.5 Unrealized & Realized PnL", 154, 253, 259, 4)
    n = DumpBlock(lk, sh, n, "CAPITAL", "3.8 Mức độ sử dụng vốn", 158, 262, 270, 6)
    n = DumpBlock(lk, sh, n, "BS", "3.6 Ghi nhận PnL theo lớp bảng cân đối", 103, 184, 189, 14)
    n = DumpBlock(lk, sh, n, "FIONBS", "7.1 FI Bond trên bảng cân đối", 169, 288, 294, 5)
    n = DumpBlock(lk, sh, n, "PNLSCEN", "3.7 Phân tích kịch bản PnL", 140, 230, 234, 7)

    ' --- SCEN ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "SCEN"
    sh.Range("A1:J1").Value = Array("KeyID", "Book", "Kind", "Scenario", "Sub", "Tenor", _
        "PV01", "Itd", "YieldBps", "ItdChange")
    n = 2
    n = DumpScen(lk, sh, n, "TB", "recent", 51, 95, 105, 92)
    n = DumpScen(lk, sh, n, "TB", "var", 51, 110, 120, 107)
    n = DumpScen(lk, sh, n, "BB", "recent", 51, 140, 150, 137)
    n = DumpScen(lk, sh, n, "BB", "var", 51, 125, 135, 122)

    ' --- VIRA4 ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "VIRA4"
    sh.Range("A1:F1").Value = Array("Month", "VIRA", "Big4", "MarketMaker", "Top3", "Actual")
    Dim vs As Worksheet
    Set vs = ThisWorkbook.Sheets(SH_VS)
    n = 2
    For r = 2 To 30
        If Txt(vs.Cells(r, 1).Value) <> "" And Txt(vs.Cells(r, 2).Value) <> "" Then
            sh.Cells(n, 1).Value = VnDate(vs.Cells(r, 1).Value)
            For c = 2 To 6
                sh.Cells(n, c).Value = vs.Cells(r, c).Value
            Next c
            n = n + 1
        End If
    Next r

    ' --- VIRA ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "VIRA"
    sh.Range("A1:H1").Value = Array("KeyID", "Scope", "Label", "PV01", "Itd", _
        "Scenario", "YieldBps", "ItdAfter")
    n = 2
    n = DumpVira(lk, sh, n, "BOOK", 66, 154, 155, 151)
    n = DumpVira(lk, sh, n, "BB_TENOR", 78, 154, 164, 151)

    ' --- RATING ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "RATING"
    sh.Range("A1:J1").Value = Array("KeyID", "Issuer", "Amount", "Rating", "ReviewDate", _
        "Pct", "CumPct", "Fitch", "Moody", "SP")
    n = 2
    For r = 203 To 223
        Dim iss As String
        iss = OneLine(lk.Cells(r, 131).Value)
        If iss <> "" Then
            sh.Cells(n, 1).Value = "RATING." & Slug(iss)
            sh.Cells(n, 2).Value = iss
            For c = 1 To 8
                If c = 3 Then
                    sh.Cells(n, 2 + c).Value = VnDate(lk.Cells(r, 131 + c).Value)
                Else
                    sh.Cells(n, 2 + c).Value = lk.Cells(r, 131 + c).Value
                End If
            Next c
            n = n + 1
        End If
    Next r

    ' --- VOL ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "VOL"
    Dim vo As Worksheet
    Set vo = ThisWorkbook.Sheets(SH_VOL)
    sh.Cells(1, 1).Value = "Date"
    For c = 1 To 10
        sh.Cells(1, c + 1).Value = OneLine(vo.Cells(1, c + 1).Value)
    Next c
    n = 2
    For r = 2 To VOL_ROWS + 1
        If Txt(vo.Cells(r, 1).Value) = "" Then Exit For
        sh.Cells(n, 1).Value = VnDate(vo.Cells(r, 1).Value)
        For c = 1 To 10
            sh.Cells(n, c + 1).Value = vo.Cells(r, c + 1).Value
        Next c
        n = n + 1
    Next r

    ' --- TEXT (doc qua Named Range) ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "TEXT"
    sh.Range("A1:D1").Value = Array("KeyID", "Mo ta", "Nguon", "Value")
    Dim tk As Variant, td As Variant
    tk = Array("market", "assessment", "noteTB", "noteTPCP", "noteFI", "noteFV", _
               "itdTB", "itdBB", "note10d", "noteVar", "noteBB10d", "noteBBVar", _
               "noteRealized", "noteScenario", "noteSign", "noteRating", "noteVira")
    td = Array("Thong tin thi truong", "Tuan thu han muc & danh gia", "Ghi chu Trading Book", _
               "Ghi chu ty le TPCP", "Ghi chu co cau FI Bond", "Ghi chu Fair value", _
               "Lo MtM Trading theo ky han", "Lo MtM Banking theo ky han", _
               "Kich ban 10 ngay Trading", "Kich ban VaR Trading", _
               "Kich ban 1 thang Banking", "Kich ban VaR Banking", _
               "Ghi chu Realized PnL", "Ghi chu kich ban PnL", "Quy uoc dau", _
               "Ghi chu xep hang FI Bond", "Ghi chu VIRA")
    For i = LBound(tk) To UBound(tk)
        sh.Cells(i + 2, 1).Value = "txt." & tk(i)
        sh.Cells(i + 2, 2).Value = td(i)
        sh.Cells(i + 2, 3).Value = NM_PREFIX & tk(i)
        sh.Cells(i + 2, 4).Value = NameVal(CStr(tk(i)))
    Next i
    sh.Columns(4).WrapText = True

    ' --- TS: ghi mot lan bang mang ---
    Set sh = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)): sh.Name = "TS"
    sh.Range("A1:D1").Value = Array("Series", "Field", "Date", "Value")
    Call DumpTS(cd, sh)

    ' --- luu ---
    dest = ThisWorkbook.Path & Application.PathSeparator & _
           "Key_" & Format$(rt.Range("B2").Value, "yyyymmdd") & ".xlsx"
    wb.SaveAs Filename:=dest, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Da xuat file key trong " & Format$(Timer - t0, "0.0") & " giay:" & vbCrLf & dest, vbInformation
End Sub

' ---------- cac ham do khoi ----------

Private Function DumpGrid(lk As Worksheet, sh As Worksheet, ByVal n As Long, _
    ByVal pfx As String, ByVal book As String, ByVal c0 As Long, _
    ByVal r0 As Long, ByVal r1 As Long, ByVal w As Long) As Long
    Dim r As Long, c As Long, t As String
    For r = r0 To r1
        t = OneLine(lk.Cells(r, c0).Value)
        If t <> "" Then
            sh.Cells(n, 1).Value = pfx & Slug(t)
            sh.Cells(n, 2).Value = book
            sh.Cells(n, 3).Value = t
            For c = 1 To w - 1
                sh.Cells(n, 3 + c).Value = lk.Cells(r, c0 + c).Value
            Next c
            n = n + 1
        End If
    Next r
    DumpGrid = n
End Function

Private Function DumpCurve(lk As Worksheet, sh As Worksheet, ByVal n As Long, _
    ByVal nm As String, ByVal c0 As Long, ByVal r0 As Long, ByVal r1 As Long) As Long
    Dim r As Long, t As String
    For r = r0 To r1
        t = OneLine(lk.Cells(r, c0).Value)
        If t <> "" Then
            sh.Cells(n, 1).Value = "CURVE." & nm & "." & Slug(t)
            sh.Cells(n, 2).Value = nm
            sh.Cells(n, 3).Value = t
            sh.Cells(n, 4).Value = lk.Cells(r, c0 + 1).Value
            sh.Cells(n, 5).Value = lk.Cells(r, c0 + 2).Value
            sh.Cells(n, 6).Value = lk.Cells(r, c0 + 3).Value
            sh.Cells(n, 7).Value = lk.Cells(r, c0 + 4).Value
            sh.Cells(n, 8).Value = lk.Cells(r, c0 + 6).Value
            sh.Cells(n, 9).Value = lk.Cells(r, c0 + 7).Value
            sh.Cells(n, 10).Value = lk.Cells(r, c0 + 8).Value
            n = n + 1
        End If
    Next r
    DumpCurve = n
End Function

Private Function DumpBlock(lk As Worksheet, sh As Worksheet, ByVal n As Long, _
    ByVal nm As String, ByVal title As String, ByVal c0 As Long, _
    ByVal r0 As Long, ByVal r1 As Long, ByVal w As Long) As Long
    Dim r As Long, c As Long, t As String
    For r = r0 To r1
        t = OneLine(lk.Cells(r, c0).Value)
        If t <> "" Then
            sh.Cells(n, 1).Value = "GRID." & nm & "." & Slug(t)
            sh.Cells(n, 2).Value = nm
            sh.Cells(n, 3).Value = title
            sh.Cells(n, 4).Value = t
            For c = 1 To w - 1
                sh.Cells(n, 4 + c).Value = lk.Cells(r, c0 + c).Value
            Next c
            n = n + 1
        End If
    Next r
    DumpBlock = n
End Function

Private Function DumpScen(lk As Worksheet, sh As Worksheet, ByVal n As Long, _
    ByVal book As String, ByVal kind As String, ByVal c0 As Long, _
    ByVal r0 As Long, ByVal r1 As Long, ByVal hdr As Long) As Long
    Dim r As Long, i As Long, tenor As String
    Dim nm(0 To 5) As String, sb(0 To 5) As String
    For i = 0 To 5
        nm(i) = OneLine(lk.Cells(hdr, c0 + 3 + i * 2).Value)
        sb(i) = OneLine(lk.Cells(hdr + 2, c0 + 3 + i * 2).Value)
        If nm(i) = "" Then nm(i) = "KB" & CStr(i + 1)
    Next i
    For r = r0 To r1
        tenor = OneLine(lk.Cells(r, c0).Value)
        If tenor <> "" Then
            For i = 0 To 5
                sh.Cells(n, 1).Value = "SCEN." & book & "." & kind & "." & Slug(nm(i)) & "." & Slug(tenor)
                sh.Cells(n, 2).Value = book
                sh.Cells(n, 3).Value = kind
                sh.Cells(n, 4).Value = nm(i)
                sh.Cells(n, 5).Value = sb(i)
                sh.Cells(n, 6).Value = tenor
                sh.Cells(n, 7).Value = lk.Cells(r, c0 + 1).Value
                sh.Cells(n, 8).Value = lk.Cells(r, c0 + 2).Value
                sh.Cells(n, 9).Value = lk.Cells(r, c0 + 3 + i * 2).Value
                sh.Cells(n, 10).Value = lk.Cells(r, c0 + 4 + i * 2).Value
                n = n + 1
            Next i
        End If
    Next r
    DumpScen = n
End Function

Private Function DumpVira(lk As Worksheet, sh As Worksheet, ByVal n As Long, _
    ByVal scope As String, ByVal c0 As Long, ByVal r0 As Long, _
    ByVal r1 As Long, ByVal hdr As Long) As Long
    Dim r As Long, i As Long, lab As String
    Dim nm(0 To 3) As String
    For i = 0 To 3
        nm(i) = OneLine(lk.Cells(hdr, c0 + 3 + i * 2).Value)
        If nm(i) = "" Then nm(i) = "VIRA" & CStr(i + 1)
    Next i
    For r = r0 To r1
        lab = OneLine(lk.Cells(r, c0).Value)
        If lab <> "" Then
            For i = 0 To 3
                sh.Cells(n, 1).Value = "VIRA." & LCase$(scope) & "." & Slug(lab) & "." & Slug(nm(i))
                sh.Cells(n, 2).Value = scope
                sh.Cells(n, 3).Value = lab
                sh.Cells(n, 4).Value = lk.Cells(r, c0 + 1).Value
                sh.Cells(n, 5).Value = lk.Cells(r, c0 + 2).Value
                sh.Cells(n, 6).Value = nm(i)
                sh.Cells(n, 7).Value = lk.Cells(r, c0 + 3 + i * 2).Value
                sh.Cells(n, 8).Value = lk.Cells(r, c0 + 4 + i * 2).Value
                n = n + 1
            Next i
        End If
    Next r
    DumpVira = n
End Function

' TS: gom vao mang roi ghi mot lan. Header nam ca o hang 1 cua Chart data.
Private Sub DumpTS(cd As Worksheet, sh As Worksheet)
    Dim maxc As Long, c As Long, i As Long, r As Long
    Dim heads() As String, isAxis() As Boolean
    Dim arr() As Variant, n As Long, cap As Long

    maxc = cd.Cells(1, cd.Columns.Count).End(xlToLeft).Column
    ReDim heads(1 To maxc)
    ReDim isAxis(1 To maxc)
    For c = 1 To maxc
        heads(c) = OneLine(cd.Cells(1, c).Value)
        Select Case heads(c)
            Case "Date", "Ngày", "Tháng", "Kỳ hạn", "STT": isAxis(c) = True
        End Select
    Next c

    cap = 20000
    ReDim arr(1 To cap, 1 To 4)
    n = 0

    For c = 1 To maxc - 1
        If heads(c) <> "" And Not isAxis(c) And isAxis(c + 1) Then
            Dim series As String, axisCol As Long, endCol As Long
            series = Slug(heads(c))
            axisCol = c + 1
            endCol = maxc
            For i = c + 2 To maxc - 1
                If heads(i) <> "" And isAxis(i + 1) Then
                    endCol = i - 1
                    Exit For
                End If
            Next i
            For r = 2 To cd.Rows.Count
                If Txt(cd.Cells(r, axisCol).Value) = "" Then Exit For
                Dim dtxt As String
                dtxt = VnDate(cd.Cells(r, axisCol).Value)
                For i = axisCol + 1 To endCol
                    If heads(i) <> "" And Not isAxis(i) Then
                        If Txt(cd.Cells(r, i).Value) <> "" Then
                            n = n + 1
                            If n > cap Then
                                cap = cap * 2
                                ReDim Preserve arr(1 To cap, 1 To 4)
                            End If
                            arr(n, 1) = series
                            arr(n, 2) = heads(i)
                            arr(n, 3) = dtxt
                            arr(n, 4) = cd.Cells(r, i).Value
                        End If
                    End If
                Next i
            Next r
        End If
    Next c

    If n > 0 Then sh.Range("A2").Resize(n, 4).Value = arr
End Sub

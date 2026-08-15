Attribute VB_Name = "z_XuatFileKey"
Option Explicit

Private Const SHEET_SRC As String = "Linked (1)"
Private Const SHEET_RPT As String = "Report"
Private Const COL_KEYID As Long = 14
Private Const SCHEMA As String = "bond.key.v1"

Private Function Fold(ByVal s As String) As String
    Dim src As Variant, dst As Variant, i As Long
    src = Array("àáạảãâầấậẩẫăằắặẳẵ", "èéẹẻẽêềếệểễ", "ìíịỉĩ", _
                "òóọỏõôồốộổỗơờớợởỡ", "ùúụủũưừứựửữ", "ỳýỵỷỹ", "đ")
    dst = Array("a", "e", "i", "o", "u", "y", "d")
    s = LCase$(s)
    For i = LBound(src) To UBound(src)
        Dim j As Long
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

Private Function BookAt(ByVal r As Long) As String
    Select Case True
        Case r >= 5 And r <= 30:  BookAt = "TB_MSB"
        Case r >= 32 And r <= 53: BookAt = "BB_MSB"
        Case r >= 55 And r <= 60: BookAt = "TB_SBV"
        Case r >= 62 And r <= 67: BookAt = "BB_SBV"
        Case r = 69:              BookAt = "OTHER"
        Case r >= 71 And r <= 74: BookAt = "FIBOND"
        Case Else:                BookAt = ""
    End Select
End Function

Private Function CodeOf(ByVal book As String) As String
    Select Case book
        Case "TB_MSB": CodeOf = "2.1."
        Case "BB_MSB": CodeOf = "2.2."
        Case "TB_SBV": CodeOf = "2.3."
        Case "BB_SBV": CodeOf = "2.4."
        Case "OTHER":  CodeOf = "2.5."
        Case "FIBOND": CodeOf = "2.6."
    End Select
End Function

Private Function SectionOf(ByVal book As String) As String
    Select Case book
        Case "TB_MSB": SectionOf = "TRADING BOOK NOI BO"
        Case "BB_MSB": SectionOf = "BANKING BOOK NOI BO"
        Case "TB_SBV": SectionOf = "TRADING BOOK SBV"
        Case "BB_SBV": SectionOf = "BANKING BOOK SBV"
        Case "OTHER":  SectionOf = "KHAC"
        Case "FIBOND": SectionOf = "FI Bond & CD"
    End Select
End Function

Public Sub TaoCotKeyID()
    Dim ws As Worksheet, r As Long, book As String, lab As String, base As String
    Dim seen As Object, n As Long
    Set ws = ThisWorkbook.Sheets(SHEET_SRC)
    Set seen = CreateObject("Scripting.Dictionary")

    ws.Cells(2, COL_KEYID).Value = "KeyID"
    For r = 4 To 80
        book = BookAt(r)
        lab = Trim$(CStr(ws.Cells(r, 3).Value))
        If book <> "" And lab <> "" Then
            base = Slug(lab)
            If seen.Exists(book & "|" & base) Then
                n = seen(book & "|" & base) + 1
                seen(book & "|" & base) = n
                base = base & CStr(n)
            Else
                seen.Add book & "|" & base, 1
            End If
            ws.Cells(r, COL_KEYID).Value = "S2." & book & "." & base
        End If
    Next r
    MsgBox "Da ghi cot KeyID vao " & SHEET_SRC & " cot N.", vbInformation
End Sub

Public Sub XuatFileKey()
    Dim src As Worksheet, rpt As Worksheet, rt As Worksheet
    Dim wb As Workbook, mt As Worksheet, dt As Worksheet, tx As Worksheet
    Dim r As Long, i As Long, book As String, keyid As String
    Dim stamp As String, dest As String

    Set src = ThisWorkbook.Sheets(SHEET_SRC)
    Set rpt = ThisWorkbook.Sheets(SHEET_RPT)
    Set rt = ThisWorkbook.Sheets("Run Tool")

    If Trim$(CStr(src.Cells(2, COL_KEYID).Value)) <> "KeyID" Then
        MsgBox "Chua co cot KeyID. Chay TaoCotKeyID truoc.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)

    Set mt = wb.Sheets(1): mt.Name = "META"
    mt.Range("A1:C1").Value = Array("Key", "Value", "Ghi chu")
    mt.Range("A2:B2").Value = Array("schema", SCHEMA)
    mt.Range("A3:B3").Value = Array("asOf", Format$(rt.Range("B2").Value, "dd/mm/yyyy"))
    mt.Range("A4:B4").Value = Array("dateYest", Format$(rt.Range("B3").Value, "dd/mm/yyyy"))
    mt.Range("A5:B5").Value = Array("dateLastMonth", Format$(rt.Range("B4").Value, "dd/mm/yyyy"))
    mt.Range("A6:B6").Value = Array("dateLastQuarter", Format$(src.Range("H3").Value, "dd/mm/yyyy"))
    mt.Range("A7:B7").Value = Array("dateLastYear", Format$(rt.Range("B7").Value, "dd/mm/yyyy"))
    mt.Range("A8:B8").Value = Array("rptTitle", "Bao cao rui ro thi truong")
    mt.Range("A9:B9").Value = Array("rptSubtitle", "Bao cao Desk Bond")
    mt.Range("A10:B10").Value = Array("generatedAt", Format$(Now, "dd/mm/yyyy hh:nn"))

    Set dt = wb.Sheets.Add(After:=mt): dt.Name = "DATA_S2"
    dt.Range("A1:Q1").Value = Array("KeyID", "Book", "Code", "Section", "STT", "Nhom", _
        "ChiTieu", "Sub", "Today", "DtD", "Yesterday", "LastMonth", "LastQuarter", _
        "LastYear", "Limit", "Used", "Light")

    i = 2
    For r = 4 To 80
        book = BookAt(r)
        keyid = Trim$(CStr(src.Cells(r, COL_KEYID).Value))
        If book <> "" And keyid <> "" Then
            dt.Cells(i, 1).Value = keyid
            dt.Cells(i, 2).Value = book
            dt.Cells(i, 3).Value = CodeOf(book)
            dt.Cells(i, 4).Value = SectionOf(book)
            dt.Cells(i, 5).Value = src.Cells(r, 1).Value
            dt.Cells(i, 6).Value = src.Cells(r, 2).Value
            dt.Cells(i, 7).Value = Replace$(CStr(src.Cells(r, 3).Value), vbLf, " ")
            dt.Cells(i, 8).Value = IIf(Left$(Trim$(CStr(src.Cells(r, 3).Value)), 2) = "a)" _
                                    Or Left$(Trim$(CStr(src.Cells(r, 3).Value)), 2) = "b)", 1, 0)
            dt.Cells(i, 9).Value = src.Cells(r, 4).Value
            dt.Cells(i, 10).Value = src.Cells(r, 5).Value
            dt.Cells(i, 11).Value = src.Cells(r, 6).Value
            dt.Cells(i, 12).Value = src.Cells(r, 7).Value
            dt.Cells(i, 13).Value = src.Cells(r, 8).Value
            dt.Cells(i, 14).Value = src.Cells(r, 9).Value
            dt.Cells(i, 15).Value = src.Cells(r, 10).Value
            dt.Cells(i, 16).Value = src.Cells(r, 11).Value
            dt.Cells(i, 17).Value = src.Cells(r, 12).Value
            i = i + 1
        End If
    Next r

    Set tx = wb.Sheets.Add(After:=dt): tx.Name = "TEXT"
    tx.Range("A1:D1").Value = Array("KeyID", "Mo ta", "Nguon", "Value")
    tx.Range("A2:D2").Value = Array("txt.market", "Thong tin thi truong", "Report!B4", rpt.Range("B4").Value)
    tx.Range("A3:D3").Value = Array("txt.assessment", "Tuan thu han muc & danh gia", "Report!H4", rpt.Range("H4").Value)
    tx.Range("A4:D4").Value = Array("txt.noteTB", "Ghi chu Trading Book", "Report!Q10", rpt.Range("Q10").Value)
    tx.Range("A5:D5").Value = Array("txt.noteTPCP", "Ghi chu ty le TPCP", "Report!Q59", rpt.Range("Q59").Value)
    tx.Range("A6:D6").Value = Array("txt.noteFI", "Ghi chu co cau FI Bond", "Report!Q66", rpt.Range("Q66").Value)
    tx.Range("A7:D7").Value = Array("txt.noteFV", "Ghi chu Fair value", "Report!B110", rpt.Range("B110").Value)

    stamp = Format$(rt.Range("B2").Value, "yyyymmdd")
    dest = ThisWorkbook.Path & Application.PathSeparator & "Key_" & stamp & ".xlsx"

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=dest, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False

    Application.ScreenUpdating = True
    MsgBox "Da xuat " & (i - 2) & " chi tieu:" & vbCrLf & dest, vbInformation
End Sub

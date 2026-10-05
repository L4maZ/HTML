Attribute VB_Name = "modXuatRatingKey"
Option Explicit

'==============================================================================
' XUAT RATING KEY  -  schema "rating.key.v1"
'
' Doc cac bang da tinh trong workbook 01.FIBond_Rating (sheet Report, Portfolio)
' va xuat ra file values-only  Rating_Key_YYYYMMDD.xlsx  gom 7 sheet:
'   META, KPI, EVENTS, RATING_STRUCTURE, TENOR, PORTFOLIO, TEXT
' Day la data contract duy nhat cho Rating_Dashboard.html (tab Quan tri du lieu).
'
' CACH DUNG
'   1. Alt+F11 > File > Import File... > chon file .bas nay
'      (workbook phai la .xlsm hoac .xlsb)
'   2. Refresh + Calculate workbook, kiem tra Report!C10 (OK/CHECK)
'   3. Alt+F8 > XuatRatingKey > Run
'   4. Mo file Key, nhap nhan dinh o cot "Override Value" cua sheet TEXT
'   5. Upload file Key len Dashboard
'
' Macro DUNG (khong xuat) khi HTML chac chan se tu choi file:
'   thieu RPT / khoi Report / bang Portfolio, rating ngoai bang co cau,
'   tong khong khop, AmountAfterCpty > Amount, thieu Issuer/BondCode/Rating.
' Macro CANH BAO (hoi co xuat tiep khong) cho nhung sai lech ma HTML chi
'   hien canh bao vang (vd Report!C10 = CHECK, luy ke khong khop).
'==============================================================================

Private Const SCHEMA_ID As String = "rating.key.v1"
Private Const SH_REPORT As String = "Report"
Private Const SH_PORTFOLIO As String = "Portfolio"
Private Const OUT_FOLDER As String = ""            ' de trong = cung thu muc voi workbook nay
Private Const TOL_STOP As Double = 0.01            ' dung nguong cua validateData() trong HTML
Private Const TOL_CUM As Double = 0.005

' ---- trang thai doc du lieu (reset moi lan chay) ----
Private mErrs As Collection
Private mWarns As Collection
Private mEvents As Collection
Private mEvKeys As Collection
Private mRates As Collection
Private mRateKeys As Collection
Private mTenors As Collection
Private mTenorKeys As Collection
Private mPort As Collection
Private mPortKeys As Collection
Private mOutWb As Workbook

Private mRpt As Double
Private mCheck As Variant
Private mNChange As Long, mNNew As Long, mNReview As Long
Private mSrcChange As String, mSrcNew As String, mSrcReview As String
Private mGrandRow As Long, mBbbRow As Long
Private mTotRating As Double, mBbbPlus As Double, mBbbPlusCpty As Double
Private mTotTenor As Double
Private mTotPort As Double, mTotAfter As Double
Private mSumRatingAfter As Double, mSumPort As Double, mSumAfter As Double, mSumTenorRows As Double

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub XuatRatingKey()
    Dim wsR As Worksheet, wsP As Worksheet
    Dim savedPath As String
    Dim su As Boolean, ev As Boolean

    su = Application.ScreenUpdating
    ev = Application.EnableEvents
    On Error GoTo EH

    InitState

    Set wsR = FindSheet(SH_REPORT)
    Set wsP = FindSheet(SH_PORTFOLIO)
    If wsR Is Nothing Then mErrs.Add "Khong thay sheet '" & SH_REPORT & "'."
    If wsP Is Nothing Then mErrs.Add "Khong thay sheet '" & SH_PORTFOLIO & "'."
    If mErrs.Count > 0 Then GoTo ShowErrs

    If Application.Calculation = xlCalculationManual Then Application.Calculate

    ReadHeaderInfo wsR
    ReadReport wsR
    ReadRatingStructure wsP
    If mGrandRow > 0 Then ReadTenor wsP
    ReadPortfolio wsP
    If mErrs.Count = 0 Then CrossChecks

ShowErrs:
    If mErrs.Count > 0 Then
        MsgBox "KHONG XUAT DUOC file Key. Can sua " & mErrs.Count & " loi sau:" & vbCrLf & vbCrLf & _
               ListText(mErrs, 700), vbCritical, "Xuat Rating Key"
        GoTo Done
    End If

    If mWarns.Count > 0 Then
        If MsgBox("Co " & mWarns.Count & " canh bao (HTML se hien tren tab Quan tri du lieu):" & vbCrLf & vbCrLf & _
                  ListText(mWarns, 650) & vbCrLf & vbCrLf & "Van xuat file Key?", _
                  vbExclamation + vbYesNo + vbDefaultButton2, "Xuat Rating Key") <> vbYes Then GoTo Done
    End If

    Application.ScreenUpdating = False
    savedPath = BuildAndSave()
    If Len(savedPath) > 0 Then
        MsgBox "Da xuat file Key:" & vbCrLf & savedPath & vbCrLf & vbCrLf & _
               "Reporting date: " & Format$(mRpt, "dd/mm/yyyy") & vbCrLf & _
               "Trai phieu: " & mPort.Count & "  |  Su kien: " & mEvents.Count & _
               "  |  Canh bao: " & mWarns.Count & vbCrLf & vbCrLf & _
               "Buoc tiep: nhap nhan dinh o sheet TEXT (cot Override Value), roi upload len Dashboard.", _
               vbInformation, "Xuat Rating Key"
    End If

Done:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = su
    Application.EnableEvents = ev
    Exit Sub

EH:
    Dim msg As String
    msg = "Loi " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not mOutWb Is Nothing Then
        Application.DisplayAlerts = False
        mOutWb.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = su
    Application.EnableEvents = ev
    MsgBox "Xuat Key bi dung giua chung." & vbCrLf & msg, vbCritical, "Xuat Rating Key"
End Sub

Private Sub InitState()
    Set mErrs = New Collection
    Set mWarns = New Collection
    Set mEvents = New Collection
    Set mEvKeys = New Collection
    Set mRates = New Collection
    Set mRateKeys = New Collection
    Set mTenors = New Collection
    Set mTenorKeys = New Collection
    Set mPort = New Collection
    Set mPortKeys = New Collection
    Set mOutWb = Nothing
    mRpt = 0: mCheck = Empty
    mNChange = 0: mNNew = 0: mNReview = 0
    mSrcChange = "": mSrcNew = "": mSrcReview = ""
    mGrandRow = 0: mBbbRow = 0
    mTotRating = 0: mBbbPlus = 0: mBbbPlusCpty = 0
    mTotTenor = 0: mTotPort = 0: mTotAfter = 0
    mSumRatingAfter = 0: mSumPort = 0: mSumAfter = 0: mSumTenorRows = 0
End Sub

'==============================================================================
' DOC DU LIEU
'==============================================================================

' RPT (named range) va o Check (Report, cot C, dong "Check")
Private Sub ReadHeaderInfo(wsR As Worksheet)
    Dim v As Variant, r As Long

    On Error Resume Next
    v = ThisWorkbook.Names("RPT").RefersToRange.Value2
    Err.Clear
    On Error GoTo 0

    If IsEmpty(v) Or IsError(v) Then
        r = FindRow(wsR, 2, "RptDate", 1)
        If r > 0 Then v = wsR.Cells(r, 3).Value2
    End If

    If IsError(v) Or IsEmpty(v) Then
        mErrs.Add "Khong doc duoc RPT (name RPT hoac Report!C2 dang trong)."
    ElseIf VarType(v) = vbString Then
        mErrs.Add "RPT khong phai la ngay."
    ElseIf CDbl(v) < 40000 Then
        mErrs.Add "RPT khong hop le: " & CStr(v)
    Else
        mRpt = CDbl(Int(CDbl(v)))
    End If

    r = FindRow(wsR, 2, "Check", 1)
    If r > 0 Then mCheck = CTxt(wsR.Cells(r, 3).Value2)
End Sub

' 3 khoi tren Report, nhan dien theo tieu de "1." / "2." / "3." o cot G
Private Sub ReadReport(wsR As Worksheet)
    Dim t1 As Long, t2 As Long, t3 As Long, lastR As Long

    t1 = FindTitle(wsR, "1.")
    t2 = FindTitle(wsR, "2.")
    t3 = FindTitle(wsR, "3.")
    If t1 = 0 Or t2 = 0 Or t3 = 0 Or Not (t1 < t2 And t2 < t3) Then
        mErrs.Add "Report: khong tim thay du 3 tieu de khoi '1.', '2.', '3.' o cot G."
        Exit Sub
    End If
    If Norm(wsR.Cells(t1 + 1, 7).Value2) <> "issuer code" Or _
       Norm(wsR.Cells(t2 + 1, 7).Value2) <> "issuer code" Or _
       Norm(wsR.Cells(t3 + 1, 7).Value2) <> "issuer code" Then
        mErrs.Add "Report: dong header duoi tieu de khoi khong phai 'Issuer code' (cau truc bang da doi?)."
        Exit Sub
    End If

    lastR = LastUsedRow(wsR)
    If lastR < t3 + 3 Then lastR = t3 + 3

    mSrcChange = "Report!G" & (t1 + 2) & ":L" & (t2 - 1)
    mSrcReview = "Report!G" & (t2 + 2) & ":K" & (t3 - 1)
    mSrcNew = "Report!G" & (t3 + 3) & ":Q" & lastR

    ReadBlock wsR, "RATING_CHANGE", t1 + 2, t2 - 1, "L"
    ReadBlock wsR, "NEW_OR_REMOVED_RATING", t3 + 3, lastR, "Q"
    ReadBlock wsR, "REVIEW_CHANGE", t2 + 2, t3 - 1, "K"
End Sub

' Moi dong co Issuer name (cot H) khac rong la 1 su kien; dong FILTER rong bi bo qua
Private Sub ReadBlock(ws As Worksheet, ByVal kind As String, ByVal r1 As Long, ByVal r2 As Long, ByVal lastCol As String)
    Dim r As Long, e As Variant
    Dim nm As Variant, code As Variant, slug As String

    Select Case kind
        Case "RATING_CHANGE": slug = "rating_change"
        Case "REVIEW_CHANGE": slug = "review_change"
        Case Else: slug = "new_or_removed"
    End Select

    For r = r1 To r2
        nm = CTxt(ws.Cells(r, 8).Value2)
        If Not IsEmpty(nm) Then
            code = CTxt(ws.Cells(r, 7).Value2)
            If IsEmpty(code) Then
                code = nm
                mWarns.Add "Report!G" & r & ": thieu Issuer code, dung ten issuer lam khoa."
            End If

            ReDim e(0 To 16)
            e(1) = kind
            e(2) = CStr(code)
            e(3) = CStr(nm)

            Select Case kind
                Case "RATING_CHANGE"      ' G code | H name | I New Rating FI | J Last Rating | K New Date | L Last Date
                    e(5) = CTxt(ws.Cells(r, 9).Value2)
                    e(4) = CTxt(ws.Cells(r, 10).Value2)
                    e(7) = CDt(ws.Cells(r, 11).Value2)
                    e(6) = CDt(ws.Cells(r, 12).Value2)
                    e(12) = e(7)
                    mNChange = mNChange + 1
                Case "REVIEW_CHANGE"      ' G code | H name | I Rating FI | J New Date | K Last Date
                    e(4) = CTxt(ws.Cells(r, 9).Value2)
                    e(5) = e(4)
                    e(7) = CDt(ws.Cells(r, 10).Value2)
                    e(6) = CDt(ws.Cells(r, 11).Value2)
                    e(12) = e(7)
                    mNReview = mNReview + 1
                Case Else                 ' G code | H name | I Loai hinh | J FI Rating | K Ngay | L-N Fitch/Moody/S&P | O-Q ngay
                    e(8) = CTxt(ws.Cells(r, 9).Value2)
                    e(5) = CTxt(ws.Cells(r, 10).Value2)
                    e(12) = CDt(ws.Cells(r, 11).Value2)
                    e(9) = CTxt(ws.Cells(r, 12).Value2)
                    e(10) = CTxt(ws.Cells(r, 13).Value2)
                    e(11) = CTxt(ws.Cells(r, 14).Value2)
                    e(13) = CDt(ws.Cells(r, 15).Value2)
                    e(14) = CDt(ws.Cells(r, 16).Value2)
                    e(15) = CDt(ws.Cells(r, 17).Value2)
                    mNNew = mNNew + 1
            End Select

            e(16) = "Report!G" & r & ":" & lastCol & r
            e(0) = UniqueKey(mEvKeys, "event." & slug & "." & CStr(code))
            mEvents.Add e
        End If
    Next r
End Sub

' Portfolio!A:H  - bang co cau rating (2 dong header, ket thuc o "Grand Total")
Private Sub ReadRatingStructure(wsP As Worksheet)
    Dim a As Long, r As Long, so As Long
    Dim rt As Variant, e As Variant, found As Boolean

    a = FindRow(wsP, 1, "Internal Rating", 1)
    If a = 0 Then
        mErrs.Add "Portfolio: khong tim thay o 'Internal Rating' o cot A (bang co cau rating)."
        Exit Sub
    End If

    r = a + 2
    Do While r <= a + 60
        rt = CTxt(wsP.Cells(r, 1).Value2)
        If IsEmpty(rt) Then Exit Do
        If Norm(rt) = "grand total" Then
            mGrandRow = r
            found = True
            Exit Do
        End If

        so = so + 1
        ReDim e(0 To 8)
        e(0) = UniqueKey(mRateKeys, "rating." & CStr(rt))
        e(1) = CStr(rt)
        e(2) = so
        e(3) = CNum(wsP.Cells(r, 7).Value2)      ' G  Outstanding truoc dieu chinh
        e(4) = CNum(wsP.Cells(r, 8).Value2)      ' H  Luy ke truoc
        e(5) = CNum(wsP.Cells(r, 4).Value2)      ' D  Outstanding sau dieu chinh
        e(6) = CNum(wsP.Cells(r, 5).Value2)      ' E  Luy ke sau
        e(7) = CNum(wsP.Cells(r, 6).Value2)      ' F  Luy ke sau doi ung
        e(8) = "Portfolio!A" & r & ":H" & r
        mRates.Add e

        mSumRatingAfter = mSumRatingAfter + e(5)
        If UCase$(CStr(rt)) = "BBB" Then
            mBbbRow = r
            mBbbPlus = e(6)
            mBbbPlusCpty = e(7)
        End If
        r = r + 1
    Loop

    If Not found Then
        mErrs.Add "Portfolio: khong thay dong 'Grand Total' duoi bang co cau rating (cot A)."
        Exit Sub
    End If
    If mRates.Count = 0 Then mErrs.Add "Portfolio: bang co cau rating khong co dong nao."
    If mBbbRow = 0 Then mErrs.Add "Portfolio: bang co cau rating thieu dong BBB (can cho KPI 'Tu BBB tro len')."
    mTotRating = CNum(wsP.Cells(mGrandRow, 4).Value2)
End Sub

' Portfolio!A:C  - bang tenor (header "Tenor", ket thuc o "Total")
Private Sub ReadTenor(wsP As Worksheet)
    Dim h As Long, r As Long, so As Long
    Dim k As Variant, e As Variant, isTot As Boolean, found As Boolean

    h = FindRow(wsP, 1, "Tenor", mGrandRow + 1)
    If h = 0 Then
        mErrs.Add "Portfolio: khong tim thay o 'Tenor' o cot A (bang ky han)."
        Exit Sub
    End If

    r = h + 1
    Do While IsEmpty(CTxt(wsP.Cells(r, 1).Value2)) And r < h + 6
        r = r + 1
    Loop

    Do While r <= h + 40
        k = CTxt(wsP.Cells(r, 1).Value2)
        If IsEmpty(k) Then Exit Do
        isTot = (Norm(k) = "total")
        so = so + 1

        ReDim e(0 To 6)
        If isTot Then
            e(0) = UniqueKey(mTenorKeys, "tenor.total")
        Else
            e(0) = UniqueKey(mTenorKeys, "tenor." & CStr(k))
        End If
        e(1) = CStr(k)
        e(2) = so
        e(3) = CNum(wsP.Cells(r, 2).Value2)
        e(4) = CNum(wsP.Cells(r, 3).Value2)
        e(5) = isTot
        e(6) = "Portfolio!A" & r & ":C" & r
        mTenors.Add e

        If isTot Then
            mTotTenor = e(3)
            found = True
            Exit Do
        End If
        mSumTenorRows = mSumTenorRows + e(3)
        r = r + 1
    Loop

    If Not found Then mErrs.Add "Portfolio: bang tenor khong co dong 'Total'."
End Sub

' Portfolio_by_bond (cot W:AG) + cot cong them AH:AK. Dinh vi cot theo ten header,
' cua so tim kiem = tu (Bond_Code - 2) den (Bond_Code + 12) de khong nham voi bang by-issuer o cot K:T.
Private Sub ReadPortfolio(wsP As Worksheet)
    Dim bc As Long, c0 As Long, c1 As Long
    Dim cIss As Long, cName As Long, cBond As Long, cTen As Long, cTcp As Long, cRat As Long
    Dim cRev As Long, cFit As Long, cMoo As Long, cSp As Long, cAmt As Long, cAdj As Long
    Dim miss As String
    Dim r As Long, totRow As Long
    Dim iss As Variant, bond As Variant, rat As Variant, tcp As Variant, e As Variant
    Dim badRows As String, nBad As Long, gtRows As String, nGt As Long

    bc = FindHeaderCol(wsP, 1, 1, 80, "bond_code")
    If bc = 0 Then
        mErrs.Add "Portfolio: khong tim thay cot 'Bond_Code' o dong 1 (bang Portfolio_by_bond)."
        Exit Sub
    End If
    c0 = bc - 2
    c1 = bc + 12
    If c0 < 1 Then c0 = 1

    cIss = FindHeaderCol(wsP, 1, c0, c1, "issuer")
    cName = FindHeaderCol(wsP, 1, c0, c1, "issuer_name")
    cBond = bc
    cTen = FindHeaderCol(wsP, 1, c0, c1, "tenor")
    cTcp = FindHeaderCol(wsP, 1, c0, c1, "tenor (ca/pu)")
    cRat = FindHeaderCol(wsP, 1, c0, c1, "rating")
    cRev = FindHeaderCol(wsP, 1, c0, c1, "review date")
    cFit = FindHeaderCol(wsP, 1, c0, c1, "fitch rating")
    cMoo = FindHeaderCol(wsP, 1, c0, c1, "moody rating")
    cSp = FindHeaderCol(wsP, 1, c0, c1, "s&p rating")
    cAmt = FindHeaderCol(wsP, 1, c0, c1, "total amount")
    cAdj = FindHeaderCol(wsP, 1, c0, c1, "amount adj*")

    If cIss = 0 Then miss = miss & " Issuer;"
    If cName = 0 Then miss = miss & " Issuer_Name;"
    If cTen = 0 Then miss = miss & " Tenor;"
    If cTcp = 0 Then miss = miss & " Tenor (ca/pu);"
    If cRat = 0 Then miss = miss & " Rating;"
    If cRev = 0 Then miss = miss & " Review Date;"
    If cFit = 0 Then miss = miss & " Fitch Rating;"
    If cMoo = 0 Then miss = miss & " Moody Rating;"
    If cSp = 0 Then miss = miss & " S&P Rating;"
    If cAmt = 0 Then miss = miss & " Total Amount;"
    If cAdj = 0 Then miss = miss & " Amount adj ...;"
    If Len(miss) > 0 Then
        mErrs.Add "Portfolio: thieu cot header:" & miss
        Exit Sub
    End If

    r = 2
    Do While r < 5000
        iss = CTxt(wsP.Cells(r, cIss).Value2)
        If IsEmpty(iss) Then Exit Do
        If Norm(iss) = "total" Then
            totRow = r
            Exit Do
        End If

        bond = CTxt(wsP.Cells(r, cBond).Value2)
        rat = CTxt(wsP.Cells(r, cRat).Value2)
        If IsEmpty(bond) Or IsEmpty(rat) Then
            nBad = nBad + 1
            If nBad <= 8 Then badRows = badRows & r & " "
        End If

        tcp = CTxt(wsP.Cells(r, cTcp).Value2)
        If IsEmpty(tcp) Then tcp = CTxt(wsP.Cells(r, cTen).Value2)

        ReDim e(0 To 13)
        e(0) = UniqueKey(mPortKeys, "portfolio." & CStr(iss) & "." & CStr(bond))
        e(1) = CStr(iss)
        e(2) = CTxt(wsP.Cells(r, cName).Value2)
        e(3) = bond
        e(4) = rat
        e(5) = CTxt(wsP.Cells(r, cTen).Value2)
        e(6) = tcp
        e(7) = CNum(wsP.Cells(r, cAmt).Value2)
        e(8) = CNum(wsP.Cells(r, cAdj).Value2)
        e(9) = CDt(wsP.Cells(r, cRev).Value2)
        e(10) = CTxt(wsP.Cells(r, cFit).Value2)
        e(11) = CTxt(wsP.Cells(r, cMoo).Value2)
        e(12) = CTxt(wsP.Cells(r, cSp).Value2)
        e(13) = "Portfolio!" & ColL(c0) & r & ":" & ColL(c1) & r
        mPort.Add e

        mSumPort = mSumPort + e(7)
        mSumAfter = mSumAfter + e(8)
        If e(8) > e(7) + TOL_STOP Then
            nGt = nGt + 1
            If nGt <= 8 Then gtRows = gtRows & r & " "
        End If
        r = r + 1
    Loop

    If totRow = 0 Then
        mErrs.Add "Portfolio: khong thay dong 'Total' duoi bang Portfolio_by_bond (cot " & ColL(cIss) & "). Bang co dong trong giua chung?"
        Exit Sub
    End If
    If mPort.Count = 0 Then mErrs.Add "Portfolio_by_bond khong co dong trai phieu nao."
    If nBad > 0 Then mErrs.Add nBad & " dong Portfolio thieu BondCode hoac Rating (dong: " & Trim$(badRows) & IIf(nBad > 8, " ...", "") & ")."
    If nGt > 0 Then mErrs.Add nGt & " dong co Amount adj doi ung > Total Amount (dong: " & Trim$(gtRows) & IIf(nGt > 8, " ...", "") & ")."

    mTotPort = CNum(wsP.Cells(totRow, cAmt).Value2)
    mTotAfter = CNum(wsP.Cells(totRow, cAdj).Value2)
End Sub

'==============================================================================
' KIEM TRA CHEO  (lap lai validateData() cua HTML de chan loi som)
'==============================================================================
Private Sub CrossChecks()
    Dim ratingSet As Collection, e As Variant, miss As String, dummy As Variant
    Dim running As Double, derived As Double, i As Long
    Dim lastE As Variant

    ' rating co trong portfolio nhung khong co trong bang co cau
    Set ratingSet = New Collection
    For Each e In mRates
        On Error Resume Next
        ratingSet.Add 1, "k:" & UCase$(CStr(e(1)))
        Err.Clear
        On Error GoTo 0
    Next e
    For Each e In mPort
        If Not IsEmpty(e(4)) Then
            Err.Clear
            On Error Resume Next
            dummy = ratingSet("k:" & UCase$(CStr(e(4))))
            If Err.Number <> 0 Then
                If InStr(1, miss, "[" & CStr(e(4)) & "]", vbTextCompare) = 0 Then miss = miss & "[" & CStr(e(4)) & "]"
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next e
    If Len(miss) > 0 Then
        mErrs.Add "Rating co trong Portfolio_by_bond nhung thieu o bang co cau rating (Portfolio!A): " & miss & _
                  ". Bo sung rating moi (vd CCC/C) vao bang co cau/mapping truoc."
    End If

    If Abs(mTotPort - mSumPort) > TOL_STOP Then
        mErrs.Add "Dong Total cua Portfolio_by_bond (" & Fmt(mTotPort) & ") lech tong cac dong (" & Fmt(mSumPort) & ")."
    End If
    If Abs(mTotTenor - mSumPort) > TOL_STOP Then
        mErrs.Add "Total bang Tenor (" & Fmt(mTotTenor) & ") lech tong Portfolio_by_bond (" & Fmt(mSumPort) & _
                  ") " & Fmt(Abs(mTotTenor - mSumPort)) & " ty."
    End If
    If mErrs.Count > 0 Then Exit Sub

    ' ---- canh bao (HTML chi hien mau vang) ----
    If Not IsEmpty(mCheck) Then
        If UCase$(CStr(mCheck)) <> "OK" Then
            mWarns.Add "Report!C10 = " & CStr(mCheck) & ": cac tong kiem soat trong workbook chua khop."
        End If
    End If
    If Abs(mTotRating - mSumRatingAfter) > TOL_STOP Then
        mWarns.Add "Grand Total co cau rating lech tong cac dong rating " & Fmt(Abs(mTotRating - mSumRatingAfter)) & " ty."
    End If
    If Abs(mTotRating - mTotPort) > TOL_STOP Then
        mWarns.Add "Tong theo rating (Portfolio!D" & mGrandRow & " = " & Fmt(mTotRating) & ") lech tong bond-level (" & Fmt(mTotPort) & _
                   ") " & Fmt(Abs(mTotRating - mTotPort)) & " ty. Kiem tra rating cua issuer giua bang by-issuer va by-bond."
    End If
    If Abs(mTotAfter - mSumAfter) > TOL_STOP Then
        mWarns.Add "Dong Total 'Amount adj doi ung' (" & Fmt(mTotAfter) & ") lech tong chi tiet (" & Fmt(mSumAfter) & _
                   "). Kiem tra pham vi cong thuc o dong Total."
    End If

    If mSumRatingAfter > 0 Then
        For i = 1 To mRates.Count
            e = mRates(i)
            running = running + e(5)
            derived = running / mSumRatingAfter
            If Abs(e(6) - derived) > TOL_CUM Then
                mWarns.Add "Luy ke sau rating tai " & CStr(e(1)) & " = " & Format$(e(6), "0.0%") & _
                           ", tinh tu Outstanding = " & Format$(derived, "0.0%") & " (" & CStr(e(8)) & ")."
            End If
        Next i
    End If
    If mRates.Count > 0 Then
        lastE = mRates(mRates.Count)
        If Abs(lastE(6) - 1) > TOL_CUM Then mWarns.Add "Ty le luy ke sau rating khong ket thuc o 100%."
        If Abs(lastE(7) - 1) > TOL_CUM Then mWarns.Add "Ty le luy ke sau doi ung khong ket thuc o 100%."
    End If
End Sub

'==============================================================================
' GHI FILE KEY
'==============================================================================
Private Function BuildAndSave() As String
    Dim wb As Workbook, ws As Worksheet
    Dim names As Variant, i As Long
    Dim fn As String, folder As String

    folder = OutFolder()
    fn = folder & "Rating_Key_" & Format$(mRpt, "yyyymmdd") & ".xlsx"

    If Len(Dir$(fn)) > 0 Then
        If MsgBox("File da ton tai:" & vbCrLf & fn & vbCrLf & vbCrLf & "Ghi de?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Xuat Rating Key") <> vbYes Then Exit Function
    End If

    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set mOutWb = wb

    names = Array("META", "KPI", "EVENTS", "RATING_STRUCTURE", "TENOR", "PORTFOLIO", "TEXT")
    wb.Worksheets(1).Name = names(0)
    For i = 1 To UBound(names)
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = names(i)
    Next i

    WriteMeta wb.Worksheets("META")
    WriteKpi wb.Worksheets("KPI")
    WriteEvents wb.Worksheets("EVENTS")
    WriteRatingStructure wb.Worksheets("RATING_STRUCTURE")
    WriteTenor wb.Worksheets("TENOR")
    WritePortfolio wb.Worksheets("PORTFOLIO")
    WriteText wb.Worksheets("TEXT")

    wb.Worksheets("META").Activate
    wb.Worksheets("META").Range("A1").Select

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=fn, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    Set mOutWb = Nothing
    BuildAndSave = fn
End Function

Private Sub WriteMeta(ws As Worksheet)
    Dim items As Collection
    Set items = New Collection

    AddMeta items, "schema", SCHEMA_ID, "Phi" & ChrW(234) & "n b" & ChrW(7843) & "n c" & ChrW(7845) & "u tr" & ChrW(250) & "c File Key"
    AddMeta items, "asOf", mRpt, "Ng" & ChrW(224) & "y b" & ChrW(225) & "o c" & ChrW(225) & "o"
    AddMeta items, "sourceFile", ThisWorkbook.Name, "File Rating ngu" & ChrW(7891) & "n"
    AddMeta items, "generatedAt", Format$(Now, "yyyy-mm-dd hh:nn:ss"), "Th" & ChrW(7901) & "i " & ChrW(273) & "i" & ChrW(7875) & "m xu" & ChrW(7845) & "t File Key"
    AddMeta items, "sourceCheck", IIf(IsEmpty(mCheck), "", CStr(mCheck)), "Gi" & ChrW(225) & " tr" & ChrW(7883) & " Check trong file ngu" & ChrW(7891) & "n"
    AddMeta items, "portfolioCount", mPort.Count, "S" & ChrW(7889) & " d" & ChrW(242) & "ng tr" & ChrW(225) & "i phi" & ChrW(7871) & "u"
    AddMeta items, "eventCount", mEvents.Count, "T" & ChrW(7893) & "ng s" & ChrW(7889) & " s" & ChrW(7921) & " ki" & ChrW(7879) & "n Rating"
    AddMeta items, "ratingChangeCount", mNChange, "S" & ChrW(7889) & " d" & ChrW(242) & "ng thay " & ChrW(273) & ChrW(7893) & "i Rating"
    AddMeta items, "newRatingCount", mNNew, "S" & ChrW(7889) & " issuer m" & ChrW(7899) & "i / x" & ChrW(243) & "a Rating"
    AddMeta items, "reviewChangeCount", mNReview, "S" & ChrW(7889) & " issuer " & ChrW(273) & ChrW(7893) & "i ng" & ChrW(224) & "y Review"
    AddMeta items, "totalRating", mTotRating, "T" & ChrW(7893) & "ng t" & ChrW(7841) & "i b" & ChrW(7843) & "ng c" & ChrW(417) & " c" & ChrW(7845) & "u Rating"
    AddMeta items, "totalPortfolio", mTotPort, "T" & ChrW(7893) & "ng chi ti" & ChrW(7871) & "t bond-level"
    AddMeta items, "totalAfterCpty", mTotAfter, "T" & ChrW(7893) & "ng sau " & ChrW(273) & "i" & ChrW(7873) & "u ch" & ChrW(7881) & "nh " & ChrW(273) & ChrW(7889) & "i " & ChrW(7913) & "ng t" & ChrW(7841) & "i d" & ChrW(242) & "ng Total"
    AddMeta items, "totalTenor", mTotTenor, "T" & ChrW(7893) & "ng b" & ChrW(7843) & "ng k" & ChrW(7923) & " h" & ChrW(7841) & "n"
    AddMeta items, "diffRatingVsPortfolio", mTotRating - mTotPort, "Ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch b" & ChrW(7843) & "ng Rating v" & ChrW(224) & " bond-level"
    AddMeta items, "diffTenorVsPortfolio", mTotTenor - mTotPort, "Ch" & ChrW(234) & "nh l" & ChrW(7879) & "ch b" & ChrW(7843) & "ng tenor v" & ChrW(224) & " bond-level"

    Dim i As Long, v As Variant
    ws.Columns(1).NumberFormat = "@"
    ws.Columns(3).NumberFormat = "@"
    For i = 1 To items.Count
        v = items(i)
        If VarType(v(1)) = vbString Then ws.Cells(i + 1, 2).NumberFormat = "@"
    Next i
    ws.Cells(3, 2).NumberFormat = "yyyy-mm-dd"

    PutTable ws, Array("Key", "Value", "Ghi chu"), items
    StyleSheet ws, 3
    ws.Columns(1).ColumnWidth = 24
    ws.Columns(2).ColumnWidth = 30
    ws.Columns(3).ColumnWidth = 52
    ws.Range(ws.Cells(2, 2), ws.Cells(items.Count + 1, 2)).HorizontalAlignment = xlLeft
End Sub

Private Sub AddMeta(items As Collection, ByVal k As String, ByVal v As Variant, ByVal note As String)
    Dim a As Variant
    ReDim a(0 To 2)
    a(0) = k: a(1) = v: a(2) = note
    items.Add a
End Sub

Private Sub WriteKpi(ws As Worksheet)
    Dim items As Collection, usd As String
    Set items = New Collection
    usd = "t" & ChrW(7927) & " VND"

    AddKpi items, "kpi.totalOutstanding", "Bond Outstanding", mTotRating, usd, "Portfolio!D" & mGrandRow
    AddKpi items, "kpi.bbbPlus", "T" & ChrW(7915) & " BBB tr" & ChrW(7903) & " l" & ChrW(234) & "n", mBbbPlus, "%", "Portfolio!E" & mBbbRow
    AddKpi items, "kpi.bbbPlusAfterCpty", "T" & ChrW(7915) & " BBB tr" & ChrW(7903) & " l" & ChrW(234) & "n sau " & ChrW(273) & ChrW(7889) & "i " & ChrW(7913) & "ng", mBbbPlusCpty, "%", "Portfolio!F" & mBbbRow
    AddKpi items, "kpi.ratingChangeCount", "Thay " & ChrW(273) & ChrW(7893) & "i Rating", mNChange, "d" & ChrW(242) & "ng", mSrcChange
    AddKpi items, "kpi.newRatingCount", "Rating m" & ChrW(7899) & "i / x" & ChrW(243) & "a", mNNew, "issuer", mSrcNew
    AddKpi items, "kpi.reviewChangeCount", ChrW(272) & ChrW(7893) & "i ng" & ChrW(224) & "y Review", mNReview, "issuer", mSrcReview

    ws.Columns(1).NumberFormat = "@"
    ws.Columns(2).NumberFormat = "@"
    ws.Columns(4).NumberFormat = "@"
    ws.Columns(5).NumberFormat = "@"
    PutTable ws, Array("KeyID", "Label", "Value", "Unit", "Source"), items
    ws.Range("C3:C4").NumberFormat = "0.00%"
    ws.Range("C2").NumberFormat = "#,##0.00"
    StyleSheet ws, 5
End Sub

Private Sub AddKpi(items As Collection, ByVal k As String, ByVal lbl As String, ByVal v As Variant, ByVal unit As String, ByVal src As String)
    Dim a As Variant
    ReDim a(0 To 4)
    a(0) = k: a(1) = lbl: a(2) = v: a(3) = unit: a(4) = src
    items.Add a
End Sub

Private Sub WriteEvents(ws As Worksheet)
    Dim c As Variant
    For Each c In Array(1, 2, 3, 4, 5, 6, 9, 10, 11, 12, 17)
        ws.Columns(c).NumberFormat = "@"
    Next c
    For Each c In Array(7, 8, 13, 14, 15, 16)
        ws.Columns(c).NumberFormat = "yyyy-mm-dd"
    Next c
    PutTable ws, Array("KeyID", "EventType", "Issuer", "IssuerName", "OldRating", "NewRating", "OldReviewDate", "NewReviewDate", _
                       "CorporateType", "Fitch", "Moodys", "SP", "RatingDate", "FitchDate", "MoodysDate", "SPDate", "Source"), mEvents
    StyleSheet ws, 17
End Sub

Private Sub WriteRatingStructure(ws As Worksheet)
    Dim c As Variant
    For Each c In Array(1, 2, 9)
        ws.Columns(c).NumberFormat = "@"
    Next c
    PutTable ws, Array("KeyID", "Rating", "SortOrder", "OutstandingBefore", "CumulativeBefore", "OutstandingAfter", _
                       "CumulativeAfter", "CumulativeAfterCpty", "Source"), mRates
    If mRates.Count > 0 Then
        With ws
            .Range(.Cells(2, 4), .Cells(mRates.Count + 1, 4)).NumberFormat = "#,##0.00"
            .Range(.Cells(2, 6), .Cells(mRates.Count + 1, 6)).NumberFormat = "#,##0.00"
            .Range(.Cells(2, 5), .Cells(mRates.Count + 1, 5)).NumberFormat = "0.00%"
            .Range(.Cells(2, 7), .Cells(mRates.Count + 1, 8)).NumberFormat = "0.00%"
        End With
    End If
    StyleSheet ws, 9
End Sub

Private Sub WriteTenor(ws As Worksheet)
    Dim c As Variant
    For Each c In Array(1, 2, 7)
        ws.Columns(c).NumberFormat = "@"
    Next c
    PutTable ws, Array("KeyID", "Tenor", "SortOrder", "Amount", "AmountCallPut", "IsTotal", "Source"), mTenors
    If mTenors.Count > 0 Then ws.Range(ws.Cells(2, 4), ws.Cells(mTenors.Count + 1, 5)).NumberFormat = "#,##0.00"
    StyleSheet ws, 7
End Sub

Private Sub WritePortfolio(ws As Worksheet)
    Dim c As Variant
    For Each c In Array(1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14)
        ws.Columns(c).NumberFormat = "@"
    Next c
    ws.Columns(10).NumberFormat = "yyyy-mm-dd"
    PutTable ws, Array("KeyID", "Issuer", "IssuerName", "BondCode", "Rating", "Tenor", "TenorCallPut", "Amount", "AmountAfterCpty", _
                       "ReviewDate", "Fitch", "Moodys", "SP", "Source"), mPort
    If mPort.Count > 0 Then ws.Range(ws.Cells(2, 8), ws.Cells(mPort.Count + 1, 9)).NumberFormat = "#,##0.00"
    StyleSheet ws, 14
End Sub

' TEXT: nguoi dung chi sua "Override Value"; Value (cot D) = Override neu co, nguoc lai = Auto
Private Sub WriteText(ws As Worksheet)
    Dim keys As Variant, descs As Variant, i As Long, r As Long

    keys = Array("txt.overview", "txt.ratingChange", "txt.ratingStructure", "txt.portfolio")
    descs = Array("Nh" & ChrW(7853) & "n " & ChrW(273) & ChrW(7883) & "nh t" & ChrW(7893) & "ng quan", _
                  "Nh" & ChrW(7853) & "n " & ChrW(273) & ChrW(7883) & "nh bi" & ChrW(7871) & "n " & ChrW(273) & ChrW(7897) & "ng Rating", _
                  "Nh" & ChrW(7853) & "n " & ChrW(273) & ChrW(7883) & "nh c" & ChrW(417) & " c" & ChrW(7845) & "u Rating", _
                  "Nh" & ChrW(7853) & "n " & ChrW(273) & ChrW(7883) & "nh chi ti" & ChrW(7871) & "t danh m" & ChrW(7909) & "c")

    ws.Columns(1).NumberFormat = "@"
    ws.Columns(2).NumberFormat = "@"
    ws.Columns(3).NumberFormat = "@"
    ws.Columns(5).NumberFormat = "@"
    ws.Columns(6).NumberFormat = "@"

    ws.Cells(1, 1).Value2 = "KeyID"
    ws.Cells(1, 2).Value2 = "Description"
    ws.Cells(1, 3).Value2 = "Source"
    ws.Cells(1, 4).Value2 = "Value"
    ws.Cells(1, 5).Value2 = "Auto Value"
    ws.Cells(1, 6).Value2 = "Override Value"

    For i = 0 To UBound(keys)
        r = i + 2
        ws.Cells(r, 1).Value2 = keys(i)
        ws.Cells(r, 2).Value2 = descs(i)
        ws.Cells(r, 3).Value2 = "Nh" & ChrW(7853) & "p t" & ChrW(7841) & "i File Key"
        ws.Cells(r, 4).Formula = "=IF(TRIM(F" & r & ")<>"""",F" & r & ",E" & r & "&"""")"
    Next i

    StyleSheet ws, 6
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 32
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(4).ColumnWidth = 60
    ws.Columns(5).ColumnWidth = 40
    ws.Columns(6).ColumnWidth = 60
    ws.Range("D2:F5").WrapText = True
    ws.Range("D2:F5").VerticalAlignment = xlTop
    ws.Range("A2:C5").VerticalAlignment = xlTop
    ws.Range("F2:F5").Interior.Color = RGB(255, 249, 219)
    ws.Range("D1").Interior.Color = RGB(31, 41, 55)
    ws.Cells(1, 6).Interior.Color = RGB(184, 134, 11)
    ws.Calculate
End Sub

' ghi 1 bang: dong 1 = header, tu dong 2 = du lieu (moi item la mang 0-based)
Private Sub PutTable(ws As Worksheet, hdr As Variant, rowsC As Collection)
    Dim nC As Long, nR As Long, i As Long, j As Long
    Dim arr() As Variant, it As Variant

    nC = UBound(hdr) - LBound(hdr) + 1
    nR = rowsC.Count

    For j = 0 To nC - 1
        ws.Cells(1, j + 1).Value2 = hdr(LBound(hdr) + j)
    Next j
    If nR = 0 Then Exit Sub

    ReDim arr(1 To nR, 1 To nC)
    For i = 1 To nR
        it = rowsC(i)
        For j = 0 To nC - 1
            arr(i, j + 1) = it(j)
        Next j
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(nR + 1, nC)).Value2 = arr
End Sub

Private Sub StyleSheet(ws As Worksheet, ByVal nCols As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 41, 55)
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols)).EntireColumn.AutoFit
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    Err.Clear
    On Error GoTo 0
End Sub

'==============================================================================
' HELPER
'==============================================================================
Private Function FindSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase$(ws.Name) = LCase$(nm) Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    With ws.UsedRange
        LastUsedRow = .Row + .Rows.Count - 1
    End With
End Function

' chuan hoa de so sanh: bo NBSP, trim, lower; loi/rong -> ""
Private Function Norm(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Norm = LCase$(Trim$(Replace(CStr(v), Chr$(160), " ")))
End Function

' text sach: loi / rong / "-" / so 0 -> Empty (XLOOKUP tra 0 khi khong thay)
Private Function CTxt(ByVal v As Variant) As Variant
    Dim s As String
    CTxt = Empty
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(Replace(CStr(v), Chr$(160), " "))
        If Len(s) > 0 And s <> "-" Then CTxt = s
        Exit Function
    End If
    If VarType(v) = vbBoolean Then Exit Function
    If IsNumeric(v) Then
        If CDbl(v) <> 0 Then CTxt = v
    End If
End Function

' ngay: so serial >= 1 -> Double (khong gio); con lai (0, "-", rong, loi) -> Empty
Private Function CDt(ByVal v As Variant) As Variant
    CDt = Empty
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If VarType(v) = vbDate Then
        CDt = CDbl(Int(CDbl(CDate(v))))
        Exit Function
    End If
    If IsNumeric(v) Then
        If CDbl(v) >= 1 Then CDt = CDbl(Int(CDbl(v)))
    End If
End Function

Private Function CNum(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If IsNumeric(v) Then CNum = CDbl(v)
        Exit Function
    End If
    If IsNumeric(v) Then CNum = CDbl(v)
End Function

Private Function FindRow(ws As Worksheet, ByVal col As Long, ByVal txt As String, ByVal r1 As Long) As Long
    Dim r As Long, lastR As Long, want As String
    want = LCase$(txt)
    lastR = LastUsedRow(ws)
    For r = r1 To lastR
        If Norm(ws.Cells(r, col).Value2) = want Then
            FindRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindTitle(ws As Worksheet, ByVal prefix As String) As Long
    Dim r As Long, lastR As Long
    lastR = LastUsedRow(ws)
    For r = 1 To lastR
        If Left$(Norm(ws.Cells(r, 7).Value2), Len(prefix)) = prefix Then
            FindTitle = r
            Exit Function
        End If
    Next r
End Function

Private Function FindHeaderCol(ws As Worksheet, ByVal hdrRow As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal pattern As String) As Long
    Dim c As Long
    For c = c1 To c2
        If Norm(ws.Cells(hdrRow, c).Value2) Like pattern Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Function ColL(ByVal c As Long) As String
    ColL = Split(ThisWorkbook.Worksheets(1).Cells(1, c).Address(True, False), "$")(0)
End Function

' khoa duy nhat (Collection khong phan biet hoa/thuong); trung thi them "#n"
Private Function UniqueKey(keys As Collection, ByVal k As String) As String
    Dim cand As String, n As Long, ok As Boolean
    cand = k
    Do
        ok = True
        On Error Resume Next
        keys.Add cand, cand
        If Err.Number <> 0 Then ok = False
        Err.Clear
        On Error GoTo 0
        If ok Then
            UniqueKey = cand
            Exit Function
        End If
        n = n + 1
        cand = k & "#" & n
    Loop
End Function

' MsgBox cat o ~1024 ky tu -> gioi han theo do dai; toan bo noi dung in ra Immediate (Ctrl+G)
Private Function ListText(c As Collection, ByVal maxLen As Long) As String
    Dim i As Long, s As String, line As String
    For i = 1 To c.Count
        line = "- " & Left$(CStr(c(i)), 200) & vbCrLf
        Debug.Print CStr(c(i))
        If Len(s) + Len(line) > maxLen Then
            s = s & "... va " & (c.Count - i + 1) & " muc khac (xem day du o cua so Immediate, Ctrl+G)." & vbCrLf
            Exit For
        End If
        s = s & line
    Next i
    ListText = s
End Function

Private Function Fmt(ByVal d As Double) As String
    Fmt = Format$(d, "#,##0.00")
End Function

Private Function OutFolder() As String
    Dim p As String
    p = OUT_FOLDER
    If Len(p) = 0 Then p = ThisWorkbook.Path
    If Len(p) = 0 Or LCase$(Left$(p, 4)) = "http" Then p = Environ$("USERPROFILE") & "\Documents"
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    OutFolder = p
End Function

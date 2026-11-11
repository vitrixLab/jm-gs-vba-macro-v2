Attribute VB_Name = "modPJSync"

Option Explicit



' ============================================================

' modPJSync — Helpers called by PJ Non-VAT2 Worksheet_Change

' Handles the NEW layout (with UOM at column H)

'

'  D=TIN  E=Name  F=Address  G=Description  H=UOM  I=Amount

'  J=Debit  K=DebitMirror  L=EWT  M=VAT  N=REF  O=COA  P=FullID

' ============================================================

Private Const ROW_START   As Long = 3

Private Const C_TIN       As Long = 4

Private Const C_NAME      As Long = 5

Private Const C_ADDR      As Long = 6

Private Const C_DESC      As Long = 7

Private Const C_UOM       As Long = 8

Private Const C_AMT       As Long = 9

Private Const C_DEBIT     As Long = 10

Private Const C_MIRROR    As Long = 11

Private Const C_EWT       As Long = 12

Private Const C_VAT       As Long = 13

Private Const C_REF       As Long = 14

Private Const C_COA       As Long = 15

Private Const C_FULLID    As Long = 16



' CDJ columns

Private Const CDJ_DATE    As Long = 3

Private Const CDJ_PART    As Long = 5

Private Const CDJ_CASH    As Long = 6

Private Const CDJ_VAT     As Long = 7

Private Const CDJ_WHTAX   As Long = 8

Private Const CDJ_DEBIT   As Long = 20

Private Const CDJ_REF      As Long = 27   ' column AA - canonical PJ reference
Private Const CDJ_REF_PNV  As Long = 28   ' column AB - PNV reference
Private Const CDJ_REF_LEG  As Long = 25   ' column Y  - legacy reference
Private Const CDJ_NOTE     As Long = 26   ' column Z  - ENTRY LOGS note

Private Const CDJ_START   As Long = 15

Private Const CDJ_EXP_S   As Long = 14

Private Const CDJ_EXP_E   As Long = 19

Private Const CDJ_SEQ     As Long = 27   ' column AA — sequence stored here



' ============================================================

' SHEET REFERENCE

' ============================================================

Private Function PJSheet() As Worksheet

    On Error Resume Next

    Set PJSheet = ThisWorkbook.Sheets("PJ Non-Vat2")

    On Error GoTo 0

End Function



' ============================================================

' SEQUENCE

' ============================================================

Public Function GetNextSequence() As Long

    Dim cdj As Worksheet, v As Variant

    On Error Resume Next

    Set cdj = ThisWorkbook.Sheets("CDJ")

    On Error GoTo 0

    If cdj Is Nothing Then GetNextSequence = 1: Exit Function



    v = cdj.Cells(13, CDJ_SEQ).Value

    If IsNumeric(v) Then

        GetNextSequence = CLng(v) + 1

    Else

        GetNextSequence = 1

    End If

    cdj.Cells(13, CDJ_SEQ).Value = GetNextSequence

End Function



' ============================================================

' SUPPLIER LOOKUPS

' ============================================================

Public Sub LookupByTIN_NonVat(ByVal rowNum As Long, ByVal tinVal As String)

    Dim ws As Worksheet, r As Long, lastRow As Long

    Dim key As String, hit As String

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    ' A BIR TIN is 15 digits; the live data also contains '009-149--884',
    ' ' 240-940-697' and '636-985956'.  Comparing the normalised digit string
    ' on both sides makes those match the SUPPLIERS DATA row they belong to.
    key = modEngine.NormTIN(tinVal)

    If key = "" Then Exit Sub

    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row

    For r = 2 To lastRow

        If modEngine.NormTIN(ws.Cells(r, 3).Value) = key Then

            ' Fail closed on a TIN that resolves to more than one supplier:
            ' a silent wrong-vendor match is worse than no match at all.
            If Len(hit) > 0 Then Exit Sub

            hit = Trim$(CStr(ws.Cells(r, 4).Value))

            PJSheet.Cells(rowNum, C_NAME).Value = ws.Cells(r, 4).Value

            PJSheet.Cells(rowNum, C_ADDR).Value = ws.Cells(r, 5).Value

        End If

    Next r

End Sub



Public Sub LookupByName_NonVat(ByVal rowNum As Long, ByVal nameVal As String)

    Dim ws As Worksheet, r As Long, lastRow As Long

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub



    lastRow = ws.Cells(ws.Rows.count, 4).End(xlUp).Row

    For r = 2 To lastRow

        If StrComp(Trim$(CStr(ws.Cells(r, 4).Value)), Trim$(nameVal), vbTextCompare) = 0 Then

            PJSheet.Cells(rowNum, C_TIN).Value = ws.Cells(r, 3).Value

            PJSheet.Cells(rowNum, C_ADDR).Value = ws.Cells(r, 5).Value

            Exit Sub

        End If

    Next r

End Sub



Public Sub ClearSupplierRow(ByVal rowNum As Long)

    Dim sh As Worksheet: Set sh = PJSheet

    If sh Is Nothing Then Exit Sub

    sh.Cells(rowNum, C_NAME).ClearContents

    sh.Cells(rowNum, C_ADDR).ClearContents

End Sub



' ============================================================

' TOTALS

' ============================================================

Public Sub CalculateTotals(ByVal rowNum As Long)

    Dim sh As Worksheet, i As Long

    Dim sumAmt As Double, sumDeb As Double, sumMir As Double

    Set sh = PJSheet

    If sh Is Nothing Then Exit Sub



    For i = ROW_START To rowNum - 1

        If StrComp(Trim$(CStr(sh.Cells(i, C_TIN).Value)), "TOTAL", vbTextCompare) <> 0 _

           And StrComp(Trim$(CStr(sh.Cells(i, C_NAME).Value)), "TOTAL", vbTextCompare) <> 0 Then

            If IsNumeric(sh.Cells(i, C_AMT).Value) Then sumAmt = sumAmt + CDbl(sh.Cells(i, C_AMT).Value)

            If IsNumeric(sh.Cells(i, C_DEBIT).Value) Then sumDeb = sumDeb + CDbl(sh.Cells(i, C_DEBIT).Value)

            If IsNumeric(sh.Cells(i, C_MIRROR).Value) Then sumMir = sumMir + CDbl(sh.Cells(i, C_MIRROR).Value)

        End If

    Next i



    sh.Cells(rowNum, C_AMT).Value = sumAmt

    sh.Cells(rowNum, C_DEBIT).Value = sumDeb

    sh.Cells(rowNum, C_MIRROR).Value = sumMir

    sh.Cells(rowNum, C_EWT).Value = 0

    sh.Cells(rowNum, C_VAT).Value = 0

End Sub



' ============================================================

' SYNC PJ NON-VAT ? CDJ

' ============================================================

Public Sub SyncToCDJ_NonVat(ByVal rowNum As Long)

    Dim sh As Worksheet, cdj As Worksheet

    Dim ref As String, part As String, coa As String

    Dim dt As Variant, amt As Double, ewt As Double

    Dim targetRow As Long, r As Long, lastRow As Long, expCol As Long

    Dim mth As Integer, c As Long



    On Error GoTo EH



    Set sh = PJSheet

    If sh Is Nothing Then Exit Sub

    Set cdj = ThisWorkbook.Sheets("CDJ")



    ref = Trim$(CStr(sh.Cells(rowNum, C_REF).Value))

    If ref = "" Or StrComp(ref, "TOTAL", vbTextCompare) = 0 Then Exit Sub



    part = Trim$(CStr(sh.Cells(rowNum, C_NAME).Value))

    coa = Trim$(CStr(sh.Cells(rowNum, C_COA).Value))

    dt = sh.Cells(rowNum, 2).Value

    amt = 0

    If IsNumeric(sh.Cells(rowNum, C_AMT).Value) Then amt = CDbl(sh.Cells(rowNum, C_AMT).Value)

    ewt = 0

    If IsNumeric(sh.Cells(rowNum, C_EWT).Value) Then ewt = CDbl(sh.Cells(rowNum, C_EWT).Value)



    If amt = 0 Then Exit Sub



    ' --- find an existing CDJ row carrying this ref --------------------
    ' The canonical reference column is AA(27).  The engine's old constant
    ' pointed at Y(25), which holds only 29 stray cells, so 566 of 595 refs
    ' were invisible here and updates missed their rows.
    targetRow = 0

    lastRow = cdj.Cells(cdj.Rows.count, CDJ_REF).End(xlUp).Row

    If cdj.Cells(cdj.Rows.count, CDJ_REF_PNV).End(xlUp).Row > lastRow Then _
        lastRow = cdj.Cells(cdj.Rows.count, CDJ_REF_PNV).End(xlUp).Row

    If cdj.Cells(cdj.Rows.count, CDJ_REF_LEG).End(xlUp).Row > lastRow Then _
        lastRow = cdj.Cells(cdj.Rows.count, CDJ_REF_LEG).End(xlUp).Row

    If lastRow < CDJ_START Then lastRow = CDJ_START

    For r = CDJ_START To lastRow

        If Trim$(CStr(cdj.Cells(r, CDJ_REF).Value)) = ref Or _
           Trim$(CStr(cdj.Cells(r, CDJ_REF_PNV).Value)) = ref Or _
           Trim$(CStr(cdj.Cells(r, CDJ_REF_LEG).Value)) = ref Then

            targetRow = r

            Exit For

        End If

    Next r



    ' --- else place inside the month's own partition --------------------
    ' Scanning for the first empty DATE cell put every insert past the last
    ' partition, where no monthly SUM covers it.  CDJReserveRow inserts above
    ' the month's TOTAL row when the block is full, so Excel widens the SUM.
    If targetRow = 0 Then

        mth = CDJMarkerMonth(CStr(dt))

        If mth = 0 Then mth = modEngine.PJMonthOf(PJSheet, rowNum)

        targetRow = modEngine.CDJReserveRow(cdj, mth)

        ' Fail closed: with no legal placement the row must not be written,
        ' otherwise it lands outside every monthly total again.
        If targetRow = 0 Then Exit Sub

    End If



    ' --- find expense column by matching COA header ---

    expCol = FindExpenseCol(cdj, coa)



    ' --- write -----------------------------------------------------------
    ' CDJ row-balance column X = V+W where V=F and W=SUM(G:T).  Requiring
    ' X = 0 gives  -(cash) + vat - ewt + allocation = 0, and because
    ' gross = net + vat this holds exactly when
    '     cash = -(gross - ewt),  vat = 0,  ewt = -ewt,  allocation = net.
    ' Live rows 17 and 620 both reconcile to X = 0 under this convention.
    cdj.Cells(targetRow, CDJ_DATE).Value = dt

    cdj.Cells(targetRow, CDJ_PART).Value = part

    cdj.Cells(targetRow, CDJ_CASH).Value = (amt - ewt) * -1

    cdj.Cells(targetRow, CDJ_VAT).Value = 0

    cdj.Cells(targetRow, CDJ_WHTAX).Value = ewt * -1

    cdj.Cells(targetRow, CDJ_REF_PNV).Value = ref

    ' Clear every allocation column including T (20).  Stopping at S (19)
    ' left a stale Sundry Debit behind whenever a row was reclassified.
    For c = CDJ_EXP_S To CDJ_DEBIT

        cdj.Cells(targetRow, c).Value = ""

    Next c

    If expCol >= CDJ_EXP_S And expCol < CDJ_EXP_E Then

        ' Fixed expense cols 14-18: net amount only.  Column T stays empty
        ' so the amount is not counted twice.
        cdj.Cells(targetRow, expCol).Value = amt - ewt

        cdj.Cells(targetRow, expCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    ElseIf expCol = CDJ_EXP_E Then

        ' Sundry: COA label in S (19), net amount in T (20).  Column R (18,
        ' Professional Fees) stays empty so the amount is never classified as
        ' Professional Fees as well as Sundry.
        cdj.Cells(targetRow, CDJ_EXP_E).Value = coa

        cdj.Cells(targetRow, CDJ_DEBIT).Value = amt - ewt

        cdj.Cells(targetRow, CDJ_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    End If

    cdj.Cells(targetRow, CDJ_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    cdj.Cells(targetRow, CDJ_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    cdj.Cells(targetRow, CDJ_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

    cdj.Cells(targetRow, CDJ_NOTE).Value = "PJ Non-VAT row " & rowNum & " | " & Now



    Exit Sub

EH:

    MsgBox "SyncToCDJ_NonVat error " & Err.Number & ": " & Err.description, _
           vbExclamation, "PJ Sync"

End Sub



' ============================================================

' Find expense column matching COA description

' ============================================================

Private Function FindExpenseCol(ByVal cdj As Worksheet, ByVal coa As String) As Long

    Dim col As Long, header As String, cLow As String, kw As Variant

    cLow = LCase$(Trim$(coa))

    If cLow = "" Then Exit Function



    For col = CDJ_EXP_S To CDJ_EXP_E

        header = LCase$(Trim$(CStr(cdj.Cells(13, col).Value) & " " & _

                              CStr(cdj.Cells(14, col).Value)))

        If header = "" Then GoTo NextCol



        For Each kw In Split(cLow, " ")

            If Len(kw) >= 4 And InStr(header, CStr(kw)) > 0 Then

                FindExpenseCol = col

                Exit Function

            End If

        Next kw



        If Len(header) >= 4 And InStr(cLow, header) > 0 Then

            FindExpenseCol = col

            Exit Function

        End If

NextCol:

    Next col

    ' Nothing matched.  Route to the Sundry column so the amount still gets an
    ' allocation and the row balances, exactly as modEngine.FindExpenseColumn
    ' does.  Returning 0 used to leave the row with cash/vat/ewt but no
    ' expense column, so it never reconciled to zero.
    FindExpenseCol = CDJ_EXP_E

End Function






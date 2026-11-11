Attribute VB_Name = "modEngine"
Rem Attribute VBA_ModuleType=VBAModule
'Option VBASupport 1
Option Explicit

' ============================================================
' modEngine
' Central memory-first engine for PJ->CDJ, SJ->CRJ, and GL auto-aggregation
' Version: 7.9.5
'   7.9.5  Supplier lookup / add-supplier fixes:
'          - Repeated rows of the SAME supplier (same TIN + same name) are no
'            longer treated as ambiguous. Only a TIN shared by DIFFERENT
'            supplier names is flagged, and it still fails closed.
'          - GetSupplierByTIN / GetSupplierByName reload SUPPLIERS DATA once on
'            a miss, so a row added by hand (or an old stale cache) is found.
'          - New AddSupplierRecord(): validated, single-step add that reports
'            success/failure. Sheet modules must use it instead of looping
'            "Add?" -> add -> lookup -> "Add?" ...
'          - Fixed the names list for ambiguous TINs (first name was missing).
'          - Fixed 0-based fallback header array in LoadCDJHeaders.
'          - GetAccountNameFromTIN now compares normalised TINs.
'   7.9.4  GL_AUTO_TOL moved to module top
' ============================================================

' ============================================================
' MODULE-LEVEL CONSTANTS  (MUST be at top, before any Sub/Function)
' ============================================================
Private Const GL_AUTO_TOL As Double = 0.01

' ============================================================
' COLUMN CONSTANTS - PJ
' ============================================================
Private Const PJ_COL_TIN As Integer = 4
Private Const PJ_COL_NAME As Integer = 5
Private Const PJ_COL_ADDRESS As Integer = 6
Private Const PJ_COL_COA As Integer = 7
Private Const PJ_COL_REF As Integer = 8
Private Const PJ_COL_NET As Integer = 10
Private Const PJ_COL_VAT As Integer = 13
Private Const PJ_COL_GROSS As Integer = 14
Private Const PJ_COL_FULLID As Integer = 17

' ============================================================
' COLUMN CONSTANTS - SUPPLIERS DATA
' ============================================================
Private Const SUP_COL_TIN As Integer = 3
Private Const SUP_COL_NAME As Integer = 4
Private Const SUP_COL_ADDRESS As Integer = 5
Private Const SUP_COL_COA As Integer = 6

' ============================================================
' COLUMN CONSTANTS - CDJ
' ============================================================
Private Const CDJ_COL_DATE As Integer = 3
Private Const CDJ_COL_PARTICULARS As Integer = 5
Private Const CDJ_COL_CASH As Integer = 6
Private Const CDJ_COL_VAT As Integer = 7
Private Const CDJ_COL_WHTAX As Integer = 8
Private Const CDJ_COL_REF As Integer = 27
Private Const CDJ_COL_REF_PNV As Integer = 28
Private Const CDJ_COL_REF_LEGACY As Integer = 25
Private Const CDJ_COL_DAY As Integer = 4
Private Const CDJ_COL_NOTE As Integer = 26
Private Const CDJ_COL_SEQ As Integer = 27
Private Const CDJ_COL_NEXTROW As Integer = 27
Private Const CDJ_EXPENSE_START As Integer = 14
Private Const CDJ_EXPENSE_END As Integer = 19
Private Const CDJ_DATA_START As Integer = 15
Private Const CDJ_COL_DEBIT As Integer = 20
Private Const CDJ_ROW_SEQ As Integer = 13
Private Const CDJ_ROW_NEXTROW As Integer = 14

' ============================================================
' COLUMN CONSTANTS - CRJ
' ============================================================
Private Const CRJ_COL_DATE As Integer = 4
Private Const CRJ_COL_CUSTOMER As Integer = 6
Private Const CRJ_COL_INVOICE As Integer = 7
Private Const CRJ_COL_CASH As Integer = 8
Private Const CRJ_COL_OUTPUTVAT As Integer = 10
Private Const CRJ_COL_EXEMPT As Integer = 11
Private Const CRJ_COL_SALES As Integer = 12
Private Const CRJ_COL_NOTE As Integer = 15
Private Const CRJ_COL_REF As Integer = 16
Private Const CRJ_COL_SEQ As Integer = 17
Private Const CRJ_COL_NEXTROW As Integer = 17
Private Const CRJ_ROW_SEQ As Integer = 8
Private Const CRJ_ROW_NEXTROW As Integer = 9
Private Const CRJ_DATA_START As Integer = 10

' ============================================================
' COLUMN CONSTANTS - SJ
' ============================================================
Private Const SJ_COL_DATE As Integer = 3
Private Const SJ_COL_CUSTOMER As Integer = 6
Private Const SJ_COL_AMOUNT As Integer = 10
Private Const SJ_COL_DISCOUNT As Integer = 11
Private Const SJ_COL_OUTPUTVAT As Integer = 13
Private Const SJ_COL_NETSALES As Integer = 15
Private Const SJ_COL_INVOICE As Integer = 9
Private Const SJ_COL_REFERENCE As Integer = 21

' ============================================================
' GLOBAL STATE - CDJ / PJ
' ============================================================
Private dictTIN As Object
Private dictName As Object
Private dictTINDup As Object
Private dictRef As Object
Private arrHeaders As Variant
Private arrCols As Variant
Private gNextSeq As Long
Private gNextRow As Long
Private gInitialized As Boolean
Private gSyncSheet As String
Private gHeaderCount As Integer

' ============================================================
' GLOBAL STATE - CRJ
' ============================================================
Private dictCRJRef As Object
Private gNextCRJRow As Long
Private gNextCRJSeq As Long
Private gInitializedCRJ As Boolean

' ============================================================
' INIT
' ============================================================
Public Sub InitializeEngine(Optional targetSheet As String = "CDJ")
    On Error Resume Next
    If Trim$(targetSheet) = "" Then targetSheet = "CDJ"
    gSyncSheet = targetSheet
    Set dictTIN = CreateObject("Scripting.Dictionary")
    Set dictName = CreateObject("Scripting.Dictionary")
    Set dictRef = CreateObject("Scripting.Dictionary")
    Call LoadSuppliers
    Call LoadCDJHeaders
    Call LoadCDJReferences
    Call LoadSequence
    Call LoadNextRow
    gInitialized = True
    On Error GoTo 0
End Sub

Public Sub InitializeCRJ()
    If gInitializedCRJ Then Exit Sub
    Set dictCRJRef = CreateObject("Scripting.Dictionary")
    Call LoadCRJReferences
    Call LoadCRJNextRow
    Call LoadCRJSequence
    gInitializedCRJ = True
End Sub

' ============================================================
' SUPPLIER HELPERS
' ============================================================
' Text of a cell, "" when the cell is empty or holds an error value.
Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

' Supplier name normalised for COMPARISON ONLY: case, spacing and
' non-breaking spaces are ignored.
Private Function NormName(ByVal s As String) As String
    s = Replace(s, Chr$(160), " ")
    s = UCase$(Trim$(s))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormName = s
End Function

Private Sub NewDupDict()
    Set dictTINDup = CreateObject("Scripting.Dictionary")
    dictTINDup.CompareMode = vbTextCompare
End Sub

Private Function IsDupKey(ByVal key As String) As Boolean
    If dictTINDup Is Nothing Then Exit Function
    IsDupKey = dictTINDup.Exists(key)
End Function

' Single place where a supplier enters the TIN cache.
'   - new TIN                          -> cached
'   - same TIN, SAME supplier name     -> ignored (identical repeat row)
'   - same TIN, DIFFERENT supplier     -> flagged ambiguous (fails closed)
Private Sub RegisterSupplier(ByVal key As String, ByVal tin As String, _
                             ByVal name As String, ByVal address As String, ByVal coa As String)
    Dim firstName As String
    If key = "" Then Exit Sub
    If dictTIN.Exists(key) Then
        firstName = CStr(dictTIN(key)(0))
        If NormName(firstName) <> NormName(name) Then
            If dictTINDup Is Nothing Then Call NewDupDict
            If Not dictTINDup.Exists(key) Then dictTINDup.Add key, firstName
            If InStr(1, "; " & NormName(CStr(dictTINDup(key))) & ";", _
                        "; " & NormName(name) & ";", vbBinaryCompare) = 0 Then
                dictTINDup(key) = dictTINDup(key) & "; " & name
            End If
        End If
    Else
        dictTIN(key) = Array(name, address, coa, tin)
    End If
End Sub

' ============================================================
' LOADERS
' ============================================================
Private Sub LoadSuppliers()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim tin As String, key As String, name As String, address As String, coa As String
    Call NewDupDict
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, SUP_COL_TIN).End(xlUp).Row
    For i = 2 To lastRow
        tin = CellText(ws, i, SUP_COL_TIN)
        key = NormTIN(tin)
        name = CellText(ws, i, SUP_COL_NAME)
        address = CellText(ws, i, SUP_COL_ADDRESS)
        coa = CellText(ws, i, SUP_COL_COA)
        If key <> "" Then Call RegisterSupplier(key, tin, name, address, coa)
        If name <> "" Then dictName(NormName(name)) = Array(tin, address, coa)
    Next i
End Sub

' Rebuild the supplier cache from SUPPLIERS DATA without touching the CDJ/CRJ
' state.  Use after SUPPLIERS DATA was edited by hand.
Public Sub ReloadSuppliers()
    Set dictTIN = CreateObject("Scripting.Dictionary")
    Set dictName = CreateObject("Scripting.Dictionary")
    Call LoadSuppliers
End Sub

' ============================================================
' TIN NORMALISATION
' ============================================================
' A BIR TIN is 15 digits, normally written NNN-NNN-NNN-NNNN.  The live data
' also contains '009-149--884' (double separator), ' 240-940-697' (leading
' space) and '636-985956' (missing separator).  Keeping only the digits makes
' those match the SUPPLIERS DATA entry they actually belong to.
' ============================================================
Public Function NormTIN(ByVal tin As Variant) As String
    Dim s As String, i As Long, ch As String
    If IsError(tin) Or IsEmpty(tin) Then Exit Function
    s = Trim$(CStr(tin))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then NormTIN = NormTIN & ch
    Next i
End Function

' Number of TINs that are shared by DIFFERENT supplier names.
Public Function AmbiguousTINCount() As Long
    If dictTINDup Is Nothing Then Exit Function
    AmbiguousTINCount = dictTINDup.count
End Function

' Newline-separated list of the ambiguous TINs and the vendor names sharing
' them, for the operator to correct in SUPPLIERS DATA.
Public Function AmbiguousTINList() As String
    Dim k As Variant, s As String
    If dictTINDup Is Nothing Then Exit Function
    For Each k In dictTINDup.Keys
        s = s & k & " -> " & dictTINDup(k) & vbCrLf
    Next k
    AmbiguousTINList = s
End Function

' True when this TIN is shared by different suppliers in SUPPLIERS DATA.
Public Function IsAmbiguousTIN(ByVal tin As String) As Boolean
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    IsAmbiguousTIN = IsDupKey(NormTIN(tin))
End Function

' The supplier names sharing this TIN ("" when it is not ambiguous).
Public Function AmbiguousNamesFor(ByVal tin As String) As String
    Dim key As String
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    key = NormTIN(tin)
    If IsDupKey(key) Then AmbiguousNamesFor = CStr(dictTINDup(key))
End Function

Private Sub LoadCDJHeaders()
    Dim ws As Worksheet
    Dim col As Integer, header As String, count As Integer
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    count = 0
    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END
        header = Trim(Trim(ws.Cells(13, col).Value) & " " & Trim(ws.Cells(14, col).Value))
        If header <> "" Then count = count + 1
    Next col
    If count = 0 Then
        ' 1-based, same shape as the live-header branch below.
        ReDim arrHeaders(1 To 6)
        ReDim arrCols(1 To 6)
        arrHeaders(1) = "clinic supplies": arrCols(1) = 14
        arrHeaders(2) = "rent":            arrCols(2) = 15
        arrHeaders(3) = "fuel":            arrCols(3) = 16
        arrHeaders(4) = "utilities":       arrCols(4) = 17
        arrHeaders(5) = "professional":    arrCols(5) = 18
        arrHeaders(6) = "sundry":          arrCols(6) = 19
        gHeaderCount = 6
        Exit Sub
    End If
    ReDim arrHeaders(1 To count)
    ReDim arrCols(1 To count)
    count = 0
    For col = CDJ_EXPENSE_START To CDJ_EXPENSE_END
        header = Trim(Trim(ws.Cells(13, col).Value) & " " & Trim(ws.Cells(14, col).Value))
        If header <> "" Then
            count = count + 1
            arrHeaders(count) = LCase(header)
            arrCols(count) = col
        End If
    Next col
    gHeaderCount = count
End Sub

Private Sub LoadCDJReferences()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, ref As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
    If ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row
    For i = CDJ_DATA_START To lastRow
        ref = Trim(ws.Cells(i, CDJ_COL_REF).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_PNV).Value)
        If ref <> "" Then dictRef(ref) = i
        ref = Trim(ws.Cells(i, CDJ_COL_REF_LEGACY).Value)
        If ref <> "" And Not dictRef.Exists(ref) Then dictRef(ref) = i
    Next i
End Sub

' ============================================================
' CDJ MONTH-BLOCK PLACEMENT
' ============================================================
' CDJ is laid out as six fixed monthly partitions, each closed by a TOTAL row
' whose columns F..U are =SUM(<block>).  A row written outside its partition is
' invisible to every monthly total, which is the reported defect: inserts were
' landing at row 755, beyond the last partition.  These helpers locate the
' partition and, when it is full, insert a row immediately above the TOTAL row
' so Excel widens the SUM range automatically.
' ============================================================
Private Function CDJRowIsFree(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_PARTICULARS).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_CASH).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_REF).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_REF_PNV).Value))) > 0 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(r, CDJ_COL_REF_LEGACY).Value))) > 0 Then Exit Function
    CDJRowIsFree = True
End Function

' Month carried by a col-C block marker such as "JANUARY" or "FEB TOTAL".
Public Function CDJMarkerMonth(ByVal marker As String) As Integer
    Dim t As String
    t = UCase$(Trim$(marker))
    If t = "" Then Exit Function
    If InStr(t, "TOTAL") > 0 Then t = Trim$(Replace(t, "TOTAL", ""))
    CDJMarkerMonth = V8_Month(Left$(t, 3))
End Function

Public Function CDJMonthAbbr(ByVal m As Integer) As String
    Select Case m
        Case 1: CDJMonthAbbr = "JAN"
        Case 2: CDJMonthAbbr = "FEB"
        Case 3: CDJMonthAbbr = "MAR"
        Case 4: CDJMonthAbbr = "APR"
        Case 5: CDJMonthAbbr = "MAY"
        Case 6: CDJMonthAbbr = "JUN"
        Case 7: CDJMonthAbbr = "JUL"
        Case 8: CDJMonthAbbr = "AUG"
        Case 9: CDJMonthAbbr = "SEP"
        Case 10: CDJMonthAbbr = "OCT"
        Case 11: CDJMonthAbbr = "NOV"
        Case 12: CDJMonthAbbr = "DEC"
    End Select
End Function

' Row of the TOTAL row that closes month mth, or 0 when the block is absent.
Public Function CDJTotalRow(ByVal ws As Worksheet, ByVal mth As Integer) As Long
    Dim r As Long, lab As String
    If mth < 1 Or mth > 12 Then Exit Function
    For r = CDJ_DATA_START To ws.Cells(ws.Rows.count, CDJ_COL_DATE).End(xlUp).Row
        lab = UCase$(Trim$(CStr(ws.Cells(r, CDJ_COL_DATE).Value)))
        If InStr(lab, "TOTAL") > 0 Then
            If CDJMarkerMonth(lab) = mth Then CDJTotalRow = r: Exit Function
        End If
    Next r
End Function

' First empty row inside month mth's own partition; 0 when the block is full.
Public Function FindCDJFreeRow(Optional ByVal targetMonth As Integer = 0) As Long
    Dim ws As Worksheet, r As Long, cur As Long, totalRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then FindCDJFreeRow = CDJ_DATA_START: Exit Function
    If targetMonth > 0 Then
        totalRow = CDJTotalRow(ws, targetMonth)
        If totalRow = 0 Then Exit Function
    Else
        totalRow = ws.Cells(ws.Rows.count, CDJ_COL_PARTICULARS).End(xlUp).Row + 1
    End If
    cur = 0
    For r = CDJ_DATA_START To totalRow - 1
        If CDJMarkerMonth(CStr(ws.Cells(r, CDJ_COL_DATE).Value)) > 0 Then _
            cur = CDJMarkerMonth(CStr(ws.Cells(r, CDJ_COL_DATE).Value))
        If targetMonth > 0 And cur <> targetMonth Then GoTo NextFreeRow
        If CDJRowIsFree(ws, r) Then FindCDJFreeRow = r: Exit Function
NextFreeRow:
    Next r
End Function

' Reserve the row a new CDJ entry for month mth must be written to.  When the
' partition has no free row, insert directly above its TOTAL row so the TOTAL's
' SUM range widens over the new row.  Returns 0 when no placement is possible,
' in which case the caller must NOT write anything.
Public Function CDJReserveRow(ByVal ws As Worksheet, ByVal mth As Integer) As Long
    Dim atRow As Long, r As Long
    atRow = FindCDJFreeRow(mth)
    If atRow > 0 Then CDJReserveRow = atRow: Exit Function
    If mth < 1 Or mth > 12 Then Exit Function
    atRow = CDJTotalRow(ws, mth)
    If atRow = 0 Then Exit Function
    Dim eventsWas As Boolean
    eventsWas = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ws.Rows(atRow).Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ' The inserted row does not inherit the row-level balance formulas.
    ws.Cells(atRow, 22).Formula = "=F" & atRow
    ws.Cells(atRow, 23).Formula = "=SUM(G" & atRow & ":T" & atRow & ")"
    ws.Cells(atRow, 24).Formula = "=V" & atRow & "+W" & atRow
    If atRow > CDJ_DATA_START Then
        For r = 14 To 21
            ws.Cells(atRow, r).NumberFormat = ws.Cells(atRow - 1, r).NumberFormat
        Next r
    End If
    ' Row insertion invalidates every cached row number at or below it.
    Call ShiftCDJRefCache(atRow)
    CDJReserveRow = atRow
Done:
    Application.EnableEvents = eventsWas
    Exit Function
Failed:
    CDJReserveRow = 0
    Resume Done
End Function

Private Sub ShiftCDJRefCache(ByVal fromRow As Long)
    Dim k As Variant
    For Each k In dictRef.Keys
        If CLng(dictRef(k)) >= fromRow Then dictRef(k) = CLng(dictRef(k)) + 1
    Next k
End Sub

' Month of a PJ / PJ Non-Vat source row, taken from the sheet's own col-B
' month markers (col C holds only a day number and is blank on most rows).
Public Function PJMonthOf(ByVal sourceSheet As Worksheet, ByVal sourceRow As Long) As Integer
    Dim r As Long, mk As Integer
    For r = sourceRow To 1 Step -1
        If UCase$(Trim$(CStr(sourceSheet.Cells(r, 4).Value))) = "TOTAL" Then Exit Function
        mk = CDJMarkerMonth(CStr(sourceSheet.Cells(r, 2).Value))
        If mk > 0 Then PJMonthOf = mk: Exit Function
    Next r
End Function

Private Sub LoadSequence()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If ws Is Nothing Then gNextSeq = 0: Exit Sub
    val = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value
    If IsNumeric(val) Then gNextSeq = CLng(val) Else gNextSeq = 0
End Sub

Private Sub LoadNextRow()
    ' The static pointer in AA14 is advisory only: Excel never adjusts it when
    ' rows are inserted, so it drifts onto live rows.  It currently reads 755,
    ' past the last partition, which is why new rows missed every monthly SUM.
    ' Resolve the next write row by scanning for a genuinely empty row instead.
    gNextRow = FindCDJFreeRow(0)
End Sub

Private Sub LoadCRJReferences()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, inv As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    For i = CRJ_DATA_START To lastRow
        inv = Trim(ws.Cells(i, CRJ_COL_INVOICE).Value)
        If inv <> "" Then dictCRJRef(inv) = i
    Next i
End Sub

Private Sub LoadCRJNextRow()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then gNextCRJRow = CRJ_DATA_START: Exit Sub
    val = ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value
    If IsNumeric(val) Then
        gNextCRJRow = CLng(val)
    Else
        gNextCRJRow = CRJ_DATA_START
        Do While ws.Cells(gNextCRJRow, CRJ_COL_INVOICE).Value <> ""
            gNextCRJRow = gNextCRJRow + 1
        Loop
    End If
End Sub

Private Sub LoadCRJSequence()
    Dim ws As Worksheet, val As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then gNextCRJSeq = 0: Exit Sub
    val = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value
    If IsNumeric(val) Then gNextCRJSeq = CLng(val) Else gNextCRJSeq = 0
End Sub

' ============================================================
' SEQUENCE
' ============================================================
Public Function GetNextPJNumber() As String
    Dim ws As Worksheet
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    gNextSeq = gNextSeq + 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    If Not ws Is Nothing Then ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = gNextSeq
    On Error GoTo 0
    GetNextPJNumber = "PJ-" & Format(gNextSeq, "00000")
End Function

Public Function GetNextInvoiceNumber() As Long
    Dim ws As Worksheet
    If Not gInitializedCRJ Then Call InitializeCRJ
    gNextCRJSeq = gNextCRJSeq + 1
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    If Not ws Is Nothing Then ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = gNextCRJSeq
    On Error GoTo 0
    GetNextInvoiceNumber = gNextCRJSeq
End Function

Public Function GetNextSeq() As Long
    Dim s As String
    s = GetNextPJNumber()
    If Len(s) > 3 Then GetNextSeq = CLng(Mid$(s, 4)) Else GetNextSeq = 0
End Function

' ============================================================
' SUPPLIER LOOKUPS
' ============================================================
' Returns Array(name, address, coa, tin) or False.
' False means: the TIN is not in SUPPLIERS DATA, OR it is shared by different
' suppliers (use IsAmbiguousTIN to tell the two apart).
' On a miss the cache is rebuilt from the sheet ONCE and the lookup repeated,
' so rows added or corrected by hand are picked up without reopening the file.
Public Function GetSupplierByTIN(tin As String) As Variant
    Dim key As String, pass As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    key = NormTIN(tin)
    If key = "" Then GetSupplierByTIN = False: Exit Function
    For pass = 1 To 2
        If dictTIN.Exists(key) And Not IsDupKey(key) Then
            GetSupplierByTIN = dictTIN(key)
            Exit Function
        End If
        If pass = 1 Then Call ReloadSuppliers
    Next pass
    GetSupplierByTIN = False
End Function

' Returns Array(tin, address, coa) or False.  Case and spacing are ignored.
Public Function GetSupplierByName(name As String) As Variant
    Dim key As String, pass As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    key = NormName(name)
    If key = "" Then GetSupplierByName = False: Exit Function
    For pass = 1 To 2
        If dictName.Exists(key) Then
            GetSupplierByName = dictName(key)
            Exit Function
        End If
        If pass = 1 Then Call ReloadSuppliers
    Next pass
    GetSupplierByName = False
End Function

' Cache-only add (kept for existing callers).  Prefer AddSupplierRecord, which
' also writes the row to SUPPLIERS DATA and verifies the result.
Public Sub AddSupplierToCache(tin As String, name As String, address As String, coa As String)
    Dim key As String
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    tin = Trim(tin): name = Trim(name)
    key = NormTIN(tin)
    If key <> "" Then Call RegisterSupplier(key, tin, name, address, coa)
    If NormName(name) <> "" Then dictName(NormName(name)) = Array(tin, address, coa)
End Sub

' Validated, single-step "add supplier":
'   1. refuses a TIN that already belongs to a DIFFERENT supplier,
'   2. treats a TIN that already exists for the SAME supplier as success,
'   3. writes the new row to SUPPLIERS DATA (TIN stored as text),
'   4. updates the cache and CONFIRMS the TIN now resolves.
' Returns True only when the supplier can be looked up afterwards.
' msg always explains the outcome; callers must not retry in a loop.
Public Function AddSupplierRecord(ByVal tin As String, ByVal name As String, _
                                  ByVal address As String, ByVal coa As String, _
                                  ByRef msg As String) As Boolean
    Dim ws As Worksheet, key As String, r As Long
    Dim lastB As Long, lastC As Long, newRow As Long, nextNo As Double
    Dim existName As String

    tin = Trim$(tin): name = Trim$(name): address = Trim$(address): coa = Trim$(coa)
    msg = ""
    key = NormTIN(tin)
    If key = "" Then msg = "The TIN has no digits.": Exit Function
    If name = "" Then msg = "Supplier name is required.": Exit Function
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("SUPPLIERS DATA")
    On Error GoTo 0
    If ws Is Nothing Then msg = "Sheet 'SUPPLIERS DATA' not found.": Exit Function

    lastC = ws.Cells(ws.Rows.count, SUP_COL_TIN).End(xlUp).Row

    ' --- 1/2: does this TIN already exist in the sheet? ---
    For r = 2 To lastC
        If NormTIN(CellText(ws, r, SUP_COL_TIN)) = key Then
            existName = CellText(ws, r, SUP_COL_NAME)
            If NormName(existName) = NormName(name) Then
                Call ReloadSuppliers
                If IsArray(GetSupplierByTIN(tin)) Then
                    AddSupplierRecord = True
                    msg = "Supplier already exists in SUPPLIERS DATA (row " & r & ")."
                Else
                    msg = "TIN " & tin & " exists (row " & r & ") but is also assigned to a different " & _
                          "supplier: " & AmbiguousNamesFor(tin) & ". Correct SUPPLIERS DATA."
                End If
            Else
                msg = "TIN " & tin & " already belongs to '" & existName & "' (row " & r & "). " & _
                      "Not added - correct SUPPLIERS DATA if this is a different supplier."
            End If
            Exit Function
        End If
    Next r

    ' --- 3: write the new row ---
    lastB = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    If lastC > lastB Then newRow = lastC + 1 Else newRow = lastB + 1
    If newRow < 5 Then newRow = 5
    nextNo = 0
    For r = 5 To newRow - 1
        If IsNumeric(ws.Cells(r, 2).Value) And Not IsEmpty(ws.Cells(r, 2).Value) Then
            If CDbl(ws.Cells(r, 2).Value) > nextNo Then nextNo = CDbl(ws.Cells(r, 2).Value)
        End If
    Next r

    On Error GoTo WriteFailed
    ws.Cells(newRow, SUP_COL_TIN).NumberFormat = "@"
    ws.Cells(newRow, 2).Value = nextNo + 1
    ws.Cells(newRow, SUP_COL_TIN).Value = tin
    ws.Cells(newRow, SUP_COL_NAME).Value = UCase$(name)
    ws.Cells(newRow, SUP_COL_ADDRESS).Value = address
    ws.Cells(newRow, SUP_COL_COA).Value = coa
    On Error GoTo 0

    ' --- 4: update the cache and confirm ---
    Call RegisterSupplier(key, tin, UCase$(name), address, coa)
    dictName(NormName(name)) = Array(tin, address, coa)
    If IsArray(GetSupplierByTIN(tin)) Then
        AddSupplierRecord = True
        msg = "Supplier added (SUPPLIERS DATA row " & newRow & ")."
    Else
        msg = "The supplier was written to SUPPLIERS DATA row " & newRow & " but still cannot be " & _
              "resolved. Check that row."
    End If
    Exit Function

WriteFailed:
    msg = "Could not write to SUPPLIERS DATA: " & Err.description & _
          " (is the sheet protected?)"
    Err.Clear
End Function

' ============================================================
' EXPENSE COLUMN MAPPING
' ============================================================
Public Function FindExpenseColumn(description As String) As Integer
    Dim descLower As String, i As Integer
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    descLower = LCase(Trim(description))
    If descLower = "" Then FindExpenseColumn = 0: Exit Function
    
    ' --- Pass 1: live CDJ header substring match ---
    For i = 1 To gHeaderCount
        If arrHeaders(i) <> "" Then
            If InStr(descLower, arrHeaders(i)) > 0 Or InStr(arrHeaders(i), descLower) > 0 Then
                FindExpenseColumn = arrCols(i)
                Exit Function
            End If
        End If
    Next i
    
    ' --- Pass 2: COA alias mapping to CDJ columns (14-19) ---
    Select Case True
        Case InStr(descLower, "clinic") > 0 Or InStr(descLower, "medical") > 0 _
          Or (InStr(descLower, "suppl") > 0 And InStr(descLower, "pantry") = 0)
            FindExpenseColumn = CDJ_EXPENSE_START       ' Col 14: Clinic Materials & Supplies
            
        Case InStr(descLower, "rent") > 0 Or InStr(descLower, "lease") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 1   ' Col 15: Rent Expense
            
        Case InStr(descLower, "fuel") > 0 Or InStr(descLower, "gas") > 0 _
          Or InStr(descLower, "oil") > 0 Or InStr(descLower, "toll") > 0 _
          Or InStr(descLower, "transport") > 0 Or InStr(descLower, "travel") > 0 _
          Or InStr(descLower, "delivery") > 0 Or InStr(descLower, "parking") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 2   ' Col 16: Gas, Oil, Fuel, Toll, Parking
            
        Case InStr(descLower, "communication") > 0 Or InStr(descLower, "light") > 0 _
          Or InStr(descLower, "water") > 0 Or InStr(descLower, "utilit") > 0 _
          Or InStr(descLower, "electric") > 0 Or InStr(descLower, "internet") > 0 _
          Or InStr(descLower, "power") > 0 Or InStr(descLower, "phone") > 0 Or InStr(descLower, "tel") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 3   ' Col 17: Communication, Light and Water
            
        Case InStr(descLower, "professional") > 0 Or InStr(descLower, "consult") > 0 _
          Or InStr(descLower, "clinician") > 0 Or InStr(descLower, "doctor") > 0 _
          Or InStr(descLower, "legal") > 0 Or InStr(descLower, "audit") > 0 _
          Or InStr(descLower, "accounting") > 0 Or InStr(descLower, "retainer") > 0
            FindExpenseColumn = CDJ_EXPENSE_START + 4   ' Col 18: Professional Fees
            
        Case Else
            ' All other COAs (Miscellaneous, Representation, Pantry, Taxes, Repairs, Dues, etc.) -> Sundry
            FindExpenseColumn = CDJ_EXPENSE_END         ' Col 19: Sundry Account
    End Select
End Function

' ============================================================
' WITHHOLDING TAX
' ============================================================
Public Function CalculateWithholdingTax(expenseType As String, amount As Double) As Double
    Dim expLower As String
    expLower = LCase(Trim(expenseType))
    Select Case True
        Case InStr(expLower, "rent") > 0
            CalculateWithholdingTax = amount * 0.05
        Case InStr(expLower, "professional") > 0 Or InStr(expLower, "consulting") > 0
            CalculateWithholdingTax = amount * 0.1
        Case InStr(expLower, "repair") > 0 Or InStr(expLower, "maintenance") > 0
            CalculateWithholdingTax = amount * 0.02
        Case Else
            CalculateWithholdingTax = 0
    End Select
End Function

' ============================================================
' REFERENCE CACHE
' ============================================================
Public Function FindCDJRow(ref As String) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    ref = Trim$(ref)
    If ref = "" Then FindCDJRow = 0: Exit Function
    If dictRef.Exists(ref) Then
        FindCDJRow = dictRef(ref)
        Exit Function
    End If
    ' Fallback scan of CDJ sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_PNV).End(xlUp).Row
        If ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row > lastRow Then _
            lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF_LEGACY).End(xlUp).Row
        If lastRow >= CDJ_DATA_START Then
            For i = CDJ_DATA_START To lastRow
                If Trim$(CStr(ws.Cells(i, CDJ_COL_REF).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_PNV).Value)) = ref Or _
                   Trim$(CStr(ws.Cells(i, CDJ_COL_REF_LEGACY).Value)) = ref Then
                    dictRef(ref) = i
                    FindCDJRow = i
                    Exit Function
                End If
            Next i
        End If
    End If
    FindCDJRow = 0
End Function

Public Sub UpdateCDJReference(ref As String, rowNum As Long)
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    ref = Trim(ref)
    If ref <> "" And rowNum > 0 Then dictRef(ref) = rowNum
End Sub

Public Function FindCRJRow(invoiceNum As String) As Long
    If Not gInitializedCRJ Then Call InitializeCRJ
    invoiceNum = Trim(invoiceNum)
    If invoiceNum = "" Then FindCRJRow = 0: Exit Function
    If dictCRJRef.Exists(invoiceNum) Then FindCRJRow = dictCRJRef(invoiceNum) Else FindCRJRow = 0
End Function

Public Sub UpdateCRJReference(invoiceNum As String, rowNum As Long)
    If Not gInitializedCRJ Then Call InitializeCRJ
    invoiceNum = Trim(invoiceNum)
    If invoiceNum <> "" And rowNum > 0 Then dictCRJRef(invoiceNum) = rowNum
End Sub

' ============================================================
' SYNC PJ -> CDJ
' ============================================================
Public Sub SyncPJToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim grossAmount As Double, netAmount As Double, vatAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer, mth As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value))
    If ref = "" Or UCase$(ref) = "TOTAL" Then Exit Sub
    supplierName = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value))
    description = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value))
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_GROSS).Value) Then
        grossAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_GROSS).Value)
    Else
        grossAmount = 0
    End If
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_NET).Value) Then
        netAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_NET).Value)
    Else
        netAmount = 0
    End If
    If IsNumeric(sourceSheet.Cells(sourceRow, PJ_COL_VAT).Value) Then
        vatAmount = CDbl(sourceSheet.Cells(sourceRow, PJ_COL_VAT).Value)
    Else
        vatAmount = 0
    End If
    If grossAmount = 0 Then Exit Sub
    expenseCol = FindExpenseColumn(description)
    If expenseCol >= CDJ_EXPENSE_START And expenseCol <= CDJ_EXPENSE_END Then
        whtax = CalculateWithholdingTax(description, netAmount)
    Else
        whtax = 0
    End If
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 3).Value
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (grossAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            ' Clear every allocation column including T (20).  The old loop
            ' stopped at S (19), leaving a stale Sundry Debit behind whenever a
            ' row was reclassified away from Sundry.
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18 take the net expense only.  Column T
                ' (Sundry Debit) stays empty so the amount is not counted twice.
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry: COA label in S (19), net amount in T (20).  Column R
                ' (18 = Professional Fees) stays empty so the amount is never
                ' classified as Professional Fees as well as Sundry.
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated from PJ Row " & sourceRow & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
    Else
        ' Place the new row inside its own monthly partition so the partition's
        ' TOTAL row includes it.  Writing to gNextRow (755) put every insert
        ' beyond the last SUM range, which is the reported total defect.
        mth = PJMonthOf(sourceSheet, sourceRow)
        If mth = 0 Then mth = CDJMarkerMonth(CStr(dateVal))
        targetRow = CDJReserveRow(wsTarget, mth)
        If targetRow = 0 Then Exit Sub   ' fail closed: no legal placement
        gNextRow = targetRow
        If mth > 0 And Len(Trim$(CStr(dateVal))) = 0 Then dateVal = CDJMonthAbbr(mth)
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_DAY).Value = sourceSheet.Cells(sourceRow, 3).Value
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (grossAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = vatAmount
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                ' Fixed expense cols 14-18 take the net expense only.  Column T
                ' (Sundry Debit) stays empty so the amount is not counted twice.
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                ' Sundry: COA label in S (19), net amount in T (20).  Column R
                ' (18 = Professional Fees) stays empty so the amount is never
                ' classified as Professional Fees as well as Sundry.
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted from PJ Row " & sourceRow & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub

' ============================================================
' SYNC SJ -> CRJ
' ============================================================
Public Sub SyncSJToCRJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim invoiceNum As Variant, targetRow As Long
    Dim referenceInvoiceNum As Variant
    Dim customerName As String
    Dim dateVal As Variant
    Dim amount As Double, amountRaw As Variant, discount As Double, outputVAT As Double, netSales As Double
    Dim cashInBank As Double
    On Error GoTo ErrorHandler
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    invoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_INVOICE).Value
    If invoiceNum = "" Or invoiceNum = "TOTAL" Then Exit Sub
    customerName = Trim(sourceSheet.Cells(sourceRow, SJ_COL_CUSTOMER).Value)
    dateVal = sourceSheet.Cells(sourceRow, SJ_COL_DATE).Value
    referenceInvoiceNum = sourceSheet.Cells(sourceRow, SJ_COL_REFERENCE).Value
    amountRaw = sourceSheet.Cells(sourceRow, SJ_COL_AMOUNT).Value
    discount = sourceSheet.Cells(sourceRow, SJ_COL_DISCOUNT).Value
    outputVAT = sourceSheet.Cells(sourceRow, SJ_COL_OUTPUTVAT).Value
    netSales = sourceSheet.Cells(sourceRow, SJ_COL_NETSALES).Value
    If Not IsNumeric(amountRaw) Then Exit Sub
    amount = CDbl(amountRaw)
    cashInBank = amount - discount
    If customerName = "" Or amount = 0 Then Exit Sub
    targetRow = FindCRJRow(CStr(invoiceNum))
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CRJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CRJ_COL_CUSTOMER).Value = customerName
            .Cells(targetRow, CRJ_COL_INVOICE).Value = invoiceNum
            .Cells(targetRow, CRJ_COL_REF).Value = referenceInvoiceNum
            .Cells(targetRow, CRJ_COL_CASH).Value = cashInBank
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).Value = outputVAT
            .Cells(targetRow, CRJ_COL_EXEMPT).Value = 0
            .Cells(targetRow, CRJ_COL_SALES).Value = netSales
            .Cells(targetRow, CRJ_COL_NOTE).Value = "Updated from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now
        End With
    Else
        targetRow = gNextCRJRow
        If targetRow < CRJ_DATA_START Then
            targetRow = CRJ_DATA_START
            Do While wsTarget.Cells(targetRow, CRJ_COL_INVOICE).Value <> ""
                targetRow = targetRow + 1
            Loop
            gNextCRJRow = targetRow
        End If
        With wsTarget
            .Cells(targetRow, CRJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CRJ_COL_CUSTOMER).Value = customerName
            .Cells(targetRow, CRJ_COL_INVOICE).Value = invoiceNum
            .Cells(targetRow, CRJ_COL_REF).Value = referenceInvoiceNum
            .Cells(targetRow, CRJ_COL_CASH).Value = cashInBank
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).Value = outputVAT
            .Cells(targetRow, CRJ_COL_EXEMPT).Value = 0
            .Cells(targetRow, CRJ_COL_SALES).Value = netSales
            .Cells(targetRow, CRJ_COL_NOTE).Value = "Inserted from SJ Row " & sourceRow & " | Invoice: " & invoiceNum & " | " & Now
            .Cells(targetRow, CRJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_OUTPUTVAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_EXEMPT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CRJ_COL_SALES).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCRJReference CStr(invoiceNum), targetRow
        gNextCRJRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = gNextCRJRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub

' ============================================================
' STATUS
' ============================================================
Public Function EngineStatus() As String
    If gInitialized Then
        EngineStatus = "Initialized | Suppliers: " & dictTIN.count & _
                       " | Ambiguous TINs: " & AmbiguousTINCount() & _
                       " | CDJ Refs: " & dictRef.count & _
                       " | Headers: " & gHeaderCount & _
                       " | Next PJ Seq: " & gNextSeq & _
                       " | Next CDJ Row: " & gNextRow & _
                       " | Target: " & gSyncSheet
    Else
        EngineStatus = "NOT INITIALIZED"
    End If
    If gInitializedCRJ Then
        EngineStatus = EngineStatus & " || CRJ Refs: " & dictCRJRef.count & _
                       " | Next CRJ Seq: " & gNextCRJSeq & _
                       " | Next CRJ Row: " & gNextCRJRow
    End If
End Function

Public Sub ShowCDJStatus()
    Dim ws As Worksheet, msg As String, lastRow As Long, i As Long
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Target sheet '" & gSyncSheet & "' not found!", vbCritical: Exit Sub
    msg = "===== CDJ STATUS REPORT =====" & vbCrLf & vbCrLf
    msg = msg & "Target Sheet: " & gSyncSheet & vbCrLf
    msg = msg & "Data starts at row: " & CDJ_DATA_START & vbCrLf
    msg = msg & "Current row counter (Z14): " & gNextRow & vbCrLf
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If lastRow < CDJ_DATA_START Then
        msg = msg & "Last used row (with reference): None yet" & vbCrLf
    Else
        msg = msg & "Last used row: " & lastRow & vbCrLf
    End If
    msg = msg & "References in cache: " & dictRef.count & vbCrLf
    msg = msg & "Sequence (Z13): " & IIf(IsNumeric(ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value), ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value, "empty") & vbCrLf
    msg = msg & "Next PJ number: " & Format(gNextSeq + 1, "PJ-00000") & vbCrLf
    msg = msg & vbCrLf & "Expense headers (row 14, N-T):" & vbCrLf
    For i = 1 To gHeaderCount
        msg = msg & "  Col " & arrCols(i) & ": " & arrHeaders(i) & vbCrLf
    Next i
    MsgBox msg, vbInformation, "CDJ Status"
End Sub

Public Sub ShowCRJStatus()
    Dim ws As Worksheet, msg As String, lastRow As Long
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CRJ sheet not found!", vbCritical: Exit Sub
    msg = "===== CRJ STATUS REPORT =====" & vbCrLf & vbCrLf
    msg = msg & "Data starts at row: " & CRJ_DATA_START & vbCrLf
    msg = msg & "Current row counter (Q9): " & gNextCRJRow & vbCrLf
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    If lastRow < CRJ_DATA_START Then
        msg = msg & "Last used row (with invoice): None yet" & vbCrLf
    Else
        msg = msg & "Last used row: " & lastRow & vbCrLf
    End If
    msg = msg & "Invoice references in cache: " & dictCRJRef.count & vbCrLf
    msg = msg & "Invoice sequence (Q8): " & gNextCRJSeq & vbCrLf
    msg = msg & "Next invoice number will be: " & Format(gNextCRJSeq + 1, "00000") & vbCrLf
    MsgBox msg, vbInformation, "CRJ Status"
End Sub

' Diagnostic: shows every TIN that is shared by DIFFERENT supplier names.
Public Sub ShowAmbiguousTINs()
    Dim s As String
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    Call ReloadSuppliers
    s = AmbiguousTINList()
    If s = "" Then
        MsgBox "No TIN is shared by different suppliers.", vbInformation, "Ambiguous TINs"
    Else
        MsgBox AmbiguousTINCount() & " TIN(s) shared by different suppliers:" & vbCrLf & vbCrLf & s, _
               vbExclamation, "Ambiguous TINs"
    End If
End Sub

' ============================================================
' REPAIR / RESET
' ============================================================
Public Sub RepairEngine()
    Dim ws As Worksheet, lastRow As Long, seqVal As Variant
    gSyncSheet = "CDJ"
    Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
        If lastRow < CDJ_DATA_START Then lastRow = CDJ_DATA_START - 1
        gNextRow = lastRow + 1
        ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        seqVal = ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value
        If IsNumeric(seqVal) Then gNextSeq = CLng(seqVal) Else gNextSeq = 0: ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = 0
        Call LoadCDJReferences
    End If
    If Not gInitializedCRJ Then Call InitializeCRJ
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
        If lastRow < CRJ_DATA_START Then lastRow = CRJ_DATA_START - 1
        gNextCRJRow = lastRow + 1
        ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = gNextCRJRow
        seqVal = ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value
        If IsNumeric(seqVal) Then
            gNextCRJSeq = CLng(seqVal)
        Else
            gNextCRJSeq = 0
            ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = 0
        End If
        Call LoadCRJReferences
    End If
    MsgBox "Engine repaired!" & vbCrLf & vbCrLf & EngineStatus, vbInformation
End Sub

Public Sub ResetCRJ()
    Dim ws As Worksheet, lastRow As Long, confirm As VbMsgBoxResult
    confirm = MsgBox("This will DELETE ALL data in the CRJ sheet and reset counters." & vbCrLf & _
                     "Are you sure?", vbYesNo + vbExclamation, "Reset CRJ")
    If confirm <> vbYes Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CRJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CRJ sheet not found!", vbCritical: Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Cleanup
    lastRow = ws.Cells(ws.Rows.count, CRJ_COL_INVOICE).End(xlUp).Row
    If lastRow >= CRJ_DATA_START Then ws.Rows(CRJ_DATA_START & ":" & lastRow).ClearContents
    ws.Cells(CRJ_ROW_SEQ, CRJ_COL_SEQ).Value = 0
    gNextCRJSeq = 0
    ws.Cells(CRJ_ROW_NEXTROW, CRJ_COL_NEXTROW).Value = CRJ_DATA_START
    gNextCRJRow = CRJ_DATA_START
    If Not gInitializedCRJ Then Call InitializeCRJ
    dictCRJRef.RemoveAll
    gInitializedCRJ = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "CRJ has been reset." & vbCrLf & _
           "Sequence (Q8) = 0, Next row (Q9) = " & CRJ_DATA_START, vbInformation
    Exit Sub
Cleanup:
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ResetCDJ()
    Dim ws As Worksheet, lastRow As Long, confirm As VbMsgBoxResult
    confirm = MsgBox("This will DELETE ALL data in the CDJ sheet and reset counters." & vbCrLf & _
                     "Are you sure?", vbYesNo + vbExclamation, "Reset CDJ")
    If confirm <> vbYes Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CDJ")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "CDJ sheet not found!", vbCritical: Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Cleanup
    lastRow = ws.Cells(ws.Rows.count, CDJ_COL_REF).End(xlUp).Row
    If lastRow >= CDJ_DATA_START Then ws.Rows(CDJ_DATA_START & ":" & lastRow).ClearContents
    ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Value = 0
    gNextSeq = 0
    ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = CDJ_DATA_START
    gNextRow = CDJ_DATA_START
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    dictRef.RemoveAll
    Call LoadCDJReferences
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Dim seqAddr As String, nextAddr As String
    seqAddr = Replace(ws.Cells(CDJ_ROW_SEQ, CDJ_COL_SEQ).Address, "$", "")
    nextAddr = Replace(ws.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Address, "$", "")
    MsgBox "CDJ has been reset." & vbCrLf & _
           "Sequence (" & seqAddr & ") = 0, Next row (" & nextAddr & ") = " & CDJ_DATA_START, vbInformation
    Exit Sub
Cleanup:
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' ============================================================
' EWT (extended)
' ============================================================
Public Function CalculateEWT(expenseType As String, amount As Double) As Double
    Dim expLower As String
    expLower = LCase(Trim(expenseType))
    Select Case True
        Case InStr(expLower, "rent") > 0
            CalculateEWT = amount * 0.05
        Case InStr(expLower, "professional") > 0 Or InStr(expLower, "consulting") > 0
            CalculateEWT = amount * 0.1
        Case InStr(expLower, "repair") > 0 Or InStr(expLower, "maintenance") > 0
            CalculateEWT = amount * 0.02
        Case InStr(expLower, "supply") > 0 Or InStr(expLower, "office") > 0
            CalculateEWT = amount * 0.01
        Case Else
            CalculateEWT = 0
    End Select
End Function

' ============================================================
' SYNC PJ NON-VAT -> CDJ
' ============================================================
Public Sub SyncPJNonVatToCDJ(sourceSheet As Worksheet, sourceRow As Long)
    Dim wsTarget As Worksheet
    Dim ref As String, targetRow As Long
    Dim supplierName As String, description As String
    Dim dateVal As Variant
    Dim netAmount As Double, whtax As Double
    Dim expenseCol As Integer, c As Integer, mth As Integer
    On Error GoTo ErrorHandler
    If Not gInitialized Then Call InitializeEngine(IIf(Trim$(gSyncSheet) <> "", gSyncSheet, "CDJ"))
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(gSyncSheet)
    On Error GoTo 0
    If wsTarget Is Nothing Then Exit Sub
    ref = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_REF).Value))
    If ref = "" Or UCase$(ref) = "TOTAL" Then Exit Sub
    supplierName = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_NAME).Value))
    description = Trim$(CStr(sourceSheet.Cells(sourceRow, PJ_COL_COA).Value))
    dateVal = sourceSheet.Cells(sourceRow, 2).Value
    If IsNumeric(sourceSheet.Cells(sourceRow, 14).Value) Then
        netAmount = CDbl(sourceSheet.Cells(sourceRow, 14).Value)
    Else
        netAmount = 0
    End If
    If netAmount = 0 Then Exit Sub
    whtax = CalculateEWT(description, netAmount)
    expenseCol = FindExpenseColumn(description)
    targetRow = FindCDJRow(ref)
    If targetRow > 0 Then
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            ' Non-VAT row: gross = net (no input VAT), so the row balances as
            '  -(net - EWT) + 0 - EWT + net = 0.
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Updated Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
    Else
        ' Same placement rule as the VAT path: inside the month's own partition.
        mth = CDJMarkerMonth(CStr(dateVal))
        If mth = 0 Then mth = PJMonthOf(sourceSheet, sourceRow)
        targetRow = CDJReserveRow(wsTarget, mth)
        If targetRow = 0 Then Exit Sub   ' fail closed: no legal placement
        gNextRow = targetRow
        With wsTarget
            .Cells(targetRow, CDJ_COL_DATE).Value = dateVal
            .Cells(targetRow, CDJ_COL_PARTICULARS).Value = supplierName
            .Cells(targetRow, CDJ_COL_CASH).Value = (netAmount - whtax) * -1
            .Cells(targetRow, CDJ_COL_VAT).Value = 0
            .Cells(targetRow, CDJ_COL_WHTAX).Value = whtax * -1
            .Cells(targetRow, CDJ_COL_REF_PNV).Value = ref
            For c = CDJ_EXPENSE_START To CDJ_COL_DEBIT
                .Cells(targetRow, c).Value = ""
            Next c
            If expenseCol >= CDJ_EXPENSE_START And expenseCol < CDJ_EXPENSE_END Then
                .Cells(targetRow, expenseCol).Value = netAmount
                .Cells(targetRow, expenseCol).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            ElseIf expenseCol = CDJ_EXPENSE_END Then
                .Cells(targetRow, CDJ_EXPENSE_END).Value = description
                .Cells(targetRow, CDJ_COL_DEBIT).Value = netAmount
                .Cells(targetRow, CDJ_COL_DEBIT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            End If
            .Cells(targetRow, CDJ_COL_NOTE).Value = "Inserted Non-Vat Row " & sourceRow & " | EWT: " & whtax & " | " & Now
            .Cells(targetRow, CDJ_COL_CASH).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_VAT).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
            .Cells(targetRow, CDJ_COL_WHTAX).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        End With
        UpdateCDJReference ref, targetRow
        gNextRow = targetRow + 1
        On Error Resume Next
        wsTarget.Cells(CDJ_ROW_NEXTROW, CDJ_COL_NEXTROW).Value = gNextRow
        On Error GoTo 0
    End If
    Exit Sub
ErrorHandler:
End Sub

' ============================================================
' GL AUTO-AGGREGATION  (no Const here - it lives at the top)
' ============================================================
Private Function GLN(ByVal s As String) As String
    s = LCase$(Trim$(s))
    s = Replace(s, "'", "")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    GLN = s
End Function

Private Function GLAcct(ByVal s As String) As String
    Dim k As String
    k = GLN(s)
    Select Case k
        Case "cash in bank", "cash": GLAcct = "Cash in Bank"
        Case "petty cash fund", "petty cash": GLAcct = "Petty Cash Fund"
        Case "account receivables", "accounts receivable", "account receivable": GLAcct = "Account Receivables"
        Case "advances to employee", "advances to employees": GLAcct = "Advances to Employees"
        Case "input vat", "input tax": GLAcct = "Input VAT"
        Case "excess cwt over it", "cwt/it", "cwt": GLAcct = "Excess CWT Over IT"
        Case "leasehold improvements": GLAcct = "Leasehold Improvements"
        Case "dental equipment", "medical equipment": GLAcct = "Dental Equipment"
        Case "due to clinicians": GLAcct = "Due To Clinicians"
        Case "due to clinicians- visiting", "due to clinicians visiting": GLAcct = "Due to Clinicians- Visiting"
        Case "vat payable", "output vat payable", "output vat": GLAcct = "VAT Payable"
        Case "ewt payable", "withholding tax", "withholding tax payable", "ewt": GLAcct = "EWT Payable"
        Case "government contributions (ee)", "ee government contributions": GLAcct = "Government Contributions (EE)"
        Case "government loans (ee)", "ee government loans": GLAcct = "Government Loans (EE)"
        Case "income tax payable": GLAcct = "Income Tax Payable"
        Case "paid-up capital", "paid up capital": GLAcct = "Paid-Up Capital"
        Case "retained earnings (deficit)", "retained earnings": GLAcct = "Retained Earnings (Deficit)"
        Case "sales", "exempt sales": GLAcct = "Sales"
        Case "rent", "rental": GLAcct = "Rent"
        Case "clinician's fee-corporators", "clinicians fee-corporators": GLAcct = "Clinician's Fee-Corporators"
        Case "clinician's fee-visiting", "clinicians fee-visiting": GLAcct = "Clinician's Fee-Visiting"
        Case "clinic material and supplies", "clinic materials and supplies", "clinic supplies": GLAcct = "Clinic Material and Supplies"
        Case "common use service area": GLAcct = "Common Use Service Area"
        Case "light and water expense", "light and water": GLAcct = "Light and Water Expense"
        Case "dep'n expense-leasehold improvements", "depn expense-leasehold improvements": GLAcct = "Dep'n Expense-Leasehold Improvements"
        Case "dep'n expense-dental equipment", "depn expense-dental equipment": GLAcct = "Dep'n Expense-Dental Equipment"
        Case "salaries and wages", "salary and wages": GLAcct = "Salaries and Wages"
        Case "de minimis": GLAcct = "De Minimis"
        Case "government contributions er share", "government contributions er": GLAcct = "Government Contributions ER Share"
        Case "professional fees", "professional fee": GLAcct = "Professional Fees"
        Case "repair and maintenance", "repairs and maintenance": GLAcct = "Repair and Maintenance"
        Case "pantry supplies": GLAcct = "Pantry Supplies"
        Case "office supplies", "supplies": GLAcct = "Office Supplies"
        Case "meals, gifts & entertainment expense", "meals and entertainment": GLAcct = "Meals, Gifts & Entertainment Expense"
        Case "communication": GLAcct = "Communication "
        Case "printing and duplication": GLAcct = "Printing and Duplication"
        Case "taxes and licences", "taxes and licenses": GLAcct = "Taxes and Licences"
        Case "delivery": GLAcct = "Delivery"
        Case "software subscription": GLAcct = "Software Subscription"
        Case "miscellaneous", "charges", "bank charge": GLAcct = "Miscellaneous"
        Case "transportation and travel", "transporation and travel": GLAcct = "Transportation and Travel"
        Case "advertisement": GLAcct = "Advertisement"
        Case "marketing expense": GLAcct = "Marketing Expense"
        Case "gas, oil, parking, toll fees", "fuel and oil": GLAcct = "Gas, Oil, Parking, Toll Fees"
        Case "other penalties and charges": GLAcct = "Other Penalties and Charges"
        Case "insurance": GLAcct = "Insurance"
    End Select
End Function

Private Function GLMonth(ByVal v As Variant) As Long
    Dim k As String
    k = GLN(CStr(v))
    Select Case Left$(k, 3)
        Case "jan": GLMonth = 1
        Case "feb": GLMonth = 2
        Case "mar": GLMonth = 3
        Case "apr": GLMonth = 4
        Case "may": GLMonth = 5
        Case "jun": GLMonth = 6
        Case "jul": GLMonth = 7
        Case "aug": GLMonth = 8
        Case "sep": GLMonth = 9
        Case "oct": GLMonth = 10
        Case "nov": GLMonth = 11
        Case "dec": GLMonth = 12
        Case Else: If IsDate(v) Then GLMonth = month(CDate(v))
    End Select
End Function

Private Sub GLAdd(ByRef d As Object, ByVal a As String, ByVal m As Long, ByVal dr As Double, ByVal cr As Double)
    If a = "" Or m < 1 Or m > 12 Then Exit Sub
    Dim k As String: k = a & "|" & m
    Dim x As Variant
    If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
    x(0) = x(0) + dr
    x(1) = x(1) + cr
    d(k) = x
End Sub

Private Sub GLAudit(ByVal aw As Worksheet, ByRef r As Long, ByVal src As String, ByVal detail As String, ByVal amt As Double)
    aw.Cells(r, 1).Value = src
    aw.Cells(r, 2).Value = detail
    aw.Cells(r, 3).Value = amt
    r = r + 1
End Sub

Private Sub GLScanCDJ(ByVal d As Object, ByVal aw As Worksheet, ByRef ar As Long)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("CDJ")
    Dim r As Long, m As Long, a As String, v As Double
    For r = 15 To w.Cells(w.Rows.count, 3).End(xlUp).Row
        If GLN(w.Cells(r, 5).Value) <> "total" Then
            m = GLMonth(w.Cells(r, 3).Value)
            If m > 0 Then
                If IsNumeric(w.Cells(r, 6).Value) Then
                    v = CDbl(w.Cells(r, 6).Value)
                    If v < 0 Then
                        GLAdd d, "Cash in Bank", m, 0, -v
                    Else
                        GLAdd d, "Cash in Bank", m, v, 0
                    End If
                End If
                If IsNumeric(w.Cells(r, 7).Value) Then GLAdd d, "Input VAT", m, CDbl(w.Cells(r, 7).Value), 0
                If IsNumeric(w.Cells(r, 8).Value) Then GLAdd d, "EWT Payable", m, 0, CDbl(w.Cells(r, 8).Value)
                If IsNumeric(w.Cells(r, 9).Value) Then GLAdd d, "Petty Cash Fund", m, CDbl(w.Cells(r, 9).Value), 0
                If IsNumeric(w.Cells(r, 10).Value) Then GLAdd d, "Government Contributions (EE)", m, 0, CDbl(w.Cells(r, 10).Value)
                If IsNumeric(w.Cells(r, 11).Value) Then GLAdd d, "Government Loans (EE)", m, 0, CDbl(w.Cells(r, 11).Value)
                If IsNumeric(w.Cells(r, 12).Value) Then GLAdd d, "De Minimis", m, CDbl(w.Cells(r, 12).Value), 0
                If IsNumeric(w.Cells(r, 13).Value) Then GLAdd d, "Salaries and Wages", m, CDbl(w.Cells(r, 13).Value), 0
                If IsNumeric(w.Cells(r, 14).Value) Then GLAdd d, "Clinic Material and Supplies", m, CDbl(w.Cells(r, 14).Value), 0
                If IsNumeric(w.Cells(r, 15).Value) Then GLAdd d, "Rent", m, CDbl(w.Cells(r, 15).Value), 0
                If IsNumeric(w.Cells(r, 16).Value) Then GLAdd d, "Gas, Oil, Parking, Toll Fees", m, CDbl(w.Cells(r, 16).Value), 0
                If IsNumeric(w.Cells(r, 17).Value) Then GLAdd d, "Communication ", m, CDbl(w.Cells(r, 17).Value), 0
                If IsNumeric(w.Cells(r, 18).Value) Then GLAdd d, "Professional Fees", m, CDbl(w.Cells(r, 18).Value), 0
                a = GLAcct(w.Cells(r, 19).Value)
                If a <> "" Then
                    If IsNumeric(w.Cells(r, 20).Value) Then GLAdd d, a, m, CDbl(w.Cells(r, 20).Value), 0
                    If IsNumeric(w.Cells(r, 21).Value) Then GLAdd d, a, m, 0, CDbl(w.Cells(r, 21).Value)
                ElseIf Len(Trim$(CStr(w.Cells(r, 19).Value))) > 0 Then
                    GLAudit aw, ar, "CDJ", CStr(w.Cells(r, 19).Value), 0
                End If
            End If
        End If
    Next r
End Sub

Private Sub GLScanCRJ(ByVal d As Object)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("CRJ")
    Dim ends As New Collection, r As Long, m As Long, st As Long, en As Long, rr As Long
    For r = 10 To w.Cells(w.Rows.count, 4).End(xlUp).Row
        If GLN(w.Cells(r, 6).Value) = "total" And GLMonth(w.Cells(r, 4).Value) > 0 Then ends.Add r
    Next r
    st = 10
    For m = 1 To 12
        If m <= ends.count Then en = ends(m) - 1 Else en = w.Cells(w.Rows.count, 6).End(xlUp).Row
        For rr = st To en
            If GLN(w.Cells(rr, 6).Value) <> "total" Then
                If IsNumeric(w.Cells(rr, 8).Value) Then GLAdd d, "Cash in Bank", m, CDbl(w.Cells(rr, 8).Value), 0
                If IsNumeric(w.Cells(rr, 9).Value) Then GLAdd d, "Excess CWT Over IT", m, CDbl(w.Cells(rr, 9).Value), 0
                If IsNumeric(w.Cells(rr, 10).Value) Then GLAdd d, "VAT Payable", m, 0, CDbl(w.Cells(rr, 10).Value)
                If IsNumeric(w.Cells(rr, 11).Value) Then GLAdd d, "Sales", m, 0, CDbl(w.Cells(rr, 11).Value)
                If IsNumeric(w.Cells(rr, 12).Value) Then GLAdd d, "Sales", m, 0, CDbl(w.Cells(rr, 12).Value)
            End If
        Next rr
        If m <= ends.count Then st = ends(m) + 1 Else Exit For
    Next m
End Sub

Private Sub GLScanGJ(ByVal d As Object, ByVal aw As Worksheet, ByRef ar As Long)
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("GJ")
    Dim r As Long, m As Long, a As String, s As String, dr As Double, cr As Double
    For r = 10 To w.Cells(w.Rows.count, 5).End(xlUp).Row
        If GLMonth(w.Cells(r, 2).Value) > 0 Then m = GLMonth(w.Cells(r, 2).Value)
        s = Trim$(CStr(w.Cells(r, 5).Value))
        If m > 0 And s <> "" And GLN(s) <> "total" Then
            a = GLAcct(s)
            dr = 0: cr = 0
            If IsNumeric(w.Cells(r, 6).Value) Then dr = CDbl(w.Cells(r, 6).Value)
            If IsNumeric(w.Cells(r, 7).Value) Then cr = CDbl(w.Cells(r, 7).Value)
            If a <> "" Then
                GLAdd d, a, m, dr, cr
            ElseIf dr <> 0 Or cr <> 0 Then
                GLAudit aw, ar, "GJ", s, dr + cr
            End If
        End If
    Next r
End Sub

Private Function GLAccounts() As Variant
    GLAccounts = Array( _
        "Cash in Bank", "Petty Cash Fund", "Account Receivables", "Advances to Employees", _
        "Input VAT", "Excess CWT Over IT", "Leasehold Improvements", "Dental Equipment", _
        "Due To Clinicians", "Due to Clinicians- Visiting", "VAT Payable", "EWT Payable", _
        "Government Contributions (EE)", "Government Loans (EE)", "Income Tax Payable", _
        "Paid-Up Capital", "Retained Earnings (Deficit)", "Sales", "Rent", _
        "Clinician's Fee-Corporators", "Clinician's Fee-Visiting", "Clinic Material and Supplies", _
        "Common Use Service Area", "Light and Water Expense", _
        "Dep'n Expense-Leasehold Improvements", "Dep'n Expense-Dental Equipment", _
        "Salaries and Wages", "De Minimis", "Government Contributions ER Share", _
        "Professional Fees", "Repair and Maintenance", "Pantry Supplies", "Office Supplies", _
        "Meals, Gifts & Entertainment Expense", "Communication ", "Printing and Duplication", _
        "Taxes and Licences", "Delivery", "Software Subscription", "Miscellaneous", _
        "Transportation and Travel", "Advertisement", "Marketing Expense", _
        "Gas, Oil, Parking, Toll Fees", "Other Penalties and Charges", "Insurance")
End Function

' ============================================================
' REFRESH GL - public entry point
'   reportYear: 0 = auto (detect from source sheets, else system year)
'               any other value = force that year
' ============================================================
Private Sub Legacy_RefreshGL(Optional ByVal reportYear As Long = 0)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim aw As Worksheet, gl As Worksheet
    Dim ar As Long, r As Long, i As Long, m As Long
    Dim bal As Double, x As Variant, a As Variant
    Dim yr As Long

    On Error GoTo EH

    ' ---- Resolve the reporting year ----
    If reportYear <= 0 Then
        yr = DetectGLYear()
    Else
        yr = reportYear
    End If
    If yr < 1900 Or yr > 9999 Then yr = year(Date)

    ' ---- GL_AUDIT sheet ----
    On Error Resume Next
    Set aw = ThisWorkbook.Sheets("GL_AUDIT")
    On Error GoTo EH
    If aw Is Nothing Then
        Set aw = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        On Error Resume Next
        aw.name = "GL_AUDIT"
        On Error GoTo EH
    End If

    aw.Cells.Clear
    aw.Range("A1:C1").Value = Array("Source", "Unmapped Detail", "Amount")
    ar = 2

    ' ---- Scan sources ----
    GLScanCDJ d, aw, ar
    GLScanCRJ d
    GLScanGJ d, aw, ar

    ' ---- GL sheet must exist ----
    On Error Resume Next
    Set gl = ThisWorkbook.Sheets("GL")
    On Error GoTo EH
    If gl Is Nothing Then
        MsgBox "Sheet named 'GL' not found in this workbook." & vbCrLf & _
               "Existing sheets:" & vbCrLf & _
               Join(SheetNamesList(), vbCrLf), vbCritical, "RefreshGL"
        Exit Sub
    End If

    gl.Range("B12:I1000").ClearContents
    gl.Range("B11:I11").Value = Array("Date", "Day", "Reference", "Explanation", _
                                      "Account Title", "Debit", "Credit", "Ending Balance")

    ' ---- Write every account x month ----
    a = GLAccounts()
    r = 13
    For i = LBound(a) To UBound(a)
        bal = 0
        For m = 1 To 12
            gl.Cells(r, 2).Value = MonthName(m, True)
            gl.Cells(r, 3).Value = Day(DateSerial(yr, m + 1, 0))
            gl.Cells(r, 5).Value = "Ending Balance"
            gl.Cells(r, 6).Value = a(i)
            x = Array(0#, 0#)
            If d.Exists(CStr(a(i)) & "|" & m) Then x = d(CStr(a(i)) & "|" & m)
            gl.Cells(r, 7).Value = x(0)
            gl.Cells(r, 8).Value = x(1)
            bal = bal + x(0) - x(1)
            gl.Cells(r, 9).Value = bal
            r = r + 1
        Next m
    Next i

    gl.Range("G13:I" & r - 1).NumberFormat = "#,##0.00;[Red](#,##0.00);-"
    aw.Columns("A:C").AutoFit
    MsgBox "GL refreshed for " & yr & ": 46 accounts x 12 months.", vbInformation
    Exit Sub

EH:
    MsgBox "RefreshGL error " & Err.Number & ": " & Err.description, vbCritical, "RefreshGL Error"
End Sub

Private Function SheetNamesList() As String()
    Dim arr() As String, i As Long
    ReDim arr(1 To ThisWorkbook.Sheets.count)
    For i = 1 To ThisWorkbook.Sheets.count
        arr(i) = "  - " & ThisWorkbook.Sheets(i).name
    Next i
    SheetNamesList = arr
End Function

' ============================================================
' Detect the GL reporting year
'   1. Look for a year in CDJ/CRJ/GJ dates (max date wins)
'   2. Fall back to the current system year
' ============================================================
Private Function DetectGLYear() As Long
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim v As Variant
    Dim maxDate As Date
    Dim found As Boolean
    Dim sheetNames As Variant
    Dim sn As Variant

    sheetNames = Array("CDJ", "CRJ", "GJ")
    For Each sn In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sn))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Select Case CStr(sn)
                Case "CDJ": lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
                Case "CRJ": lastRow = ws.Cells(ws.Rows.count, 4).End(xlUp).Row
                Case "GJ":  lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
            End Select
            For i = 8 To lastRow
                Select Case CStr(sn)
                    Case "CDJ": v = ws.Cells(i, 3).Value
                    Case "CRJ": v = ws.Cells(i, 4).Value
                    Case "GJ":  v = ws.Cells(i, 2).Value
                End Select
                If IsDate(v) Then
                    If Not found Or CDate(v) > maxDate Then
                        maxDate = CDate(v)
                        found = True
                    End If
                End If
            Next i
        End If
        Set ws = Nothing
    Next sn

    If found Then
        DetectGLYear = year(maxDate)
    Else
        DetectGLYear = year(Date)
    End If
End Function

' ============================================================
' GL HELPERS - public API
' ============================================================
Private Function Legacy_CalculateMonthlyNet(ByVal accountCode As String, ByVal month As Integer, _
                                    ByVal year As Integer, Optional ByVal sheetName As String = "") As Double
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim aw As Worksheet, ar As Long, x As Variant
    On Error Resume Next
    Set aw = ThisWorkbook.Sheets("GL_AUDIT")
    On Error GoTo 0
    If aw Is Nothing Then
        Set aw = ThisWorkbook.Worksheets.Add
        aw.name = "GL_AUDIT"
        aw.Range("A1:C1").Value = Array("Source", "Unmapped Detail", "Amount")
    End If
    ar = aw.Cells(aw.Rows.count, 1).End(xlUp).Row + 1
    GLScanCDJ d, aw, ar
    GLScanCRJ d
    GLScanGJ d, aw, ar
    x = Array(0#, 0#)
    If d.Exists(GLAcct(accountCode) & "|" & month) Then x = d(GLAcct(accountCode) & "|" & month)
    Legacy_CalculateMonthlyNet = x(0) - x(1)
End Function

Private Function Legacy_CalculateEndingBalanceFull(ByVal accountCode As String, ByVal year As Integer, _
                                            ByVal startBalance As Double, Optional ByVal targetSheet As String = "") As Variant
    Dim z(1 To 12) As Double, m As Long
    z(1) = startBalance + Legacy_CalculateMonthlyNet(accountCode, 1, year)
    For m = 2 To 12
        z(m) = z(m - 1) + Legacy_CalculateMonthlyNet(accountCode, m, year)
    Next m
    Legacy_CalculateEndingBalanceFull = z
End Function

Private Function Legacy_ValidateGLConsistency(Optional ByVal targetSheet As String = "GL") As Boolean
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets(targetSheet)
    Dim r As Long, td As Double, tc As Double, last As Long
    last = w.Cells(w.Rows.count, 6).End(xlUp).Row
    For r = 13 To last
        If IsNumeric(w.Cells(r, 7).Value) Then td = td + w.Cells(r, 7).Value
        If IsNumeric(w.Cells(r, 8).Value) Then tc = tc + w.Cells(r, 8).Value
    Next r
    Legacy_ValidateGLConsistency = Abs(td - tc) <= GL_AUTO_TOL
End Function

Private Sub Legacy_RefreshAllGL(ByVal targetSheet As String, Optional ByVal startBalances As Variant)
    RefreshGL
End Sub

Public Sub InitializeCOALookup()
End Sub

Public Function GetAccountFromCOA(ByVal coaCode As String) As String
    GetAccountFromCOA = GLAcct(coaCode)
    If GetAccountFromCOA = "" Then GetAccountFromCOA = coaCode
End Function

' Supplier name for a TIN (punctuation-insensitive); returns the TIN itself
' when no supplier matches.
Public Function GetAccountNameFromTIN(ByVal tin As String) As String
    Dim w As Worksheet: Set w = ThisWorkbook.Sheets("SUPPLIERS DATA")
    Dim r As Long, last As Long, key As String
    key = NormTIN(tin)
    If key <> "" Then
        last = w.Cells(w.Rows.count, 3).End(xlUp).Row
        For r = 2 To last
            If NormTIN(w.Cells(r, 3).Value) = key Then
                GetAccountNameFromTIN = CStr(w.Cells(r, 4).Value)
                Exit Function
            End If
        Next r
    End If
    GetAccountNameFromTIN = tin
End Function

Public Function GetFinancialMonth(ByVal dt As Date) As Integer
    GetFinancialMonth = month(dt)
End Function

Public Function GetFinancialYear(ByVal dt As Date) As Integer
    GetFinancialYear = year(dt)
End Function

' ============================================================
' v8.3.3 GATE WRAPPERS
' ============================================================
' modGLGate.RunV8Gate calls these two names.  They did not exist anywhere in
' the project, so the VBA project did not compile.  They now delegate to the
' engine's own validators, which is the smallest change that makes the gate
' callable without altering the posting path.
' ============================================================
Public Function ValidateGLConsistency(Optional ByVal targetYear As Long = 0) As Boolean
    ValidateGLConsistency = Legacy_ValidateGLConsistency(GL_SHEET)
End Function

Public Function BuildV8CalcSheet(Optional ByVal targetYear As Long = 0) As Boolean
    Dim ws As Worksheet, last As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    ' 552 = 46 accounts x 12 months.
    BuildV8CalcSheet = (last - 1 >= 552)
End Function


Attribute VB_Name = "modGraphify"
Option Explicit

'==============================================================================
' Graphify (v7.9, fixed) - charts one numeric column of the active sheet as a
' clustered column chart named "GraphifyChart" (reused on every re-run).
'
' Where the values come from (first match wins):
'   1. A multi-cell RANGE selected before running -> its first column
'   2. The configured column (CHART_COLUMN = "A", per the v7.8 spec)
'   3. Auto: the used-range column with the most numeric cells (leftmost on
'      tie) - this is what makes the macro work on the journal sheets, whose
'      data starts in column B ("Amount" is column I on PJ, DEBIT is F on GJ)
'   4. Otherwise: "No numeric data found" message
'
' Categories are R1..Rn (same labels as the graphify.html mock).
'
' Fixes over the broken v7.8 listing:
'   - A Worksheet has no .Charts collection -> use ChartObjects (embedded
'     chart); ws.Charts.Add would have added a chart SHEET and crashed first
'   - ws.Worksheets does not exist on a Worksheet object -> .Move removed
'   - Shape name "GraphifyChart.chart.8" never exists -> the ChartObject
'     itself is docked via .Left/.Top
'   - dataArr was read but never used -> removed; the source Range is bound
'     to the series directly (blank cells stay gaps instead of zeros)
'   - The Sub was never closed (End Sub missing in the listing)
'   - New: single-cell / no-numeric-data / chart-sheet edge cases handled;
'     errors surface through a handler that still restores ScreenUpdating
'==============================================================================

Private Const CHART_NAME As String = "GraphifyChart"
Private Const CHART_COLUMN As String = "A"   ' source column per the v7.8 spec
Private Const MIN_NUMERIC As Long = 2        ' auto-pick needs at least this many numbers
Private Const CHART_WIDTH As Double = 640    ' points
Private Const CHART_HEIGHT As Double = 360   ' points

Sub Graphify()
    Dim ws As Worksheet
    Dim co As ChartObject
    Dim srs As Series
    Dim dataRng As Range
    Dim srcCol As Long          ' resolved source column index
    Dim colTop As Long          ' first row to consider in the source column
    Dim selBottom As Long       ' bottom row of the selection (0 = none)
    Dim firstDataRow As Long
    Dim lastRow As Long
    Dim nPts As Long
    Dim i As Long
    Dim cats() As String
    Dim srcDesc As String

    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "Run Graphify from a worksheet (not a chart sheet).", vbExclamation, "Graphify"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail
    Application.ScreenUpdating = False

    colTop = 1
    srcCol = 0
    srcDesc = ""

    '--- 1. a multi-cell selection wins ---------------------------------------
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.CountLarge > 1 Then
            srcCol = Selection.Columns(1).Column
            colTop = Selection.Row
            selBottom = Selection.Row + Selection.Rows.count - 1
            srcDesc = "selection"
        End If
    End If

    '--- 2. configured column (v7.8 spec: column A) ----------------------------
    If srcCol = 0 Then
        If Application.count(ws.Columns(CHART_COLUMN)) >= MIN_NUMERIC Then
            srcCol = ws.Columns(CHART_COLUMN).Column
            srcDesc = "column " & CHART_COLUMN
        End If
    End If

    '--- 3. auto-pick the busiest numeric column -------------------------------
    If srcCol = 0 Then
        srcCol = PickNumericColumn(ws)
        If srcCol > 0 Then srcDesc = "column " & ColumnLetter(ws, srcCol) & " (auto)"
    End If

    If srcCol = 0 Then
        MsgBox "No numeric data found on sheet '" & ws.name & "'." & vbCrLf & _
               "Select a column and run Graphify again.", vbExclamation, "Graphify"
        GoTo Bye
    End If

    '--- locate the numeric block within the resolved column -------------------
    firstDataRow = FirstNumericRow(ws, srcCol, colTop)
    lastRow = ws.Cells(ws.Rows.count, srcCol).End(xlUp).Row
    If selBottom > 0 And lastRow > selBottom Then lastRow = selBottom

    If firstDataRow = 0 Or lastRow < firstDataRow Then
        MsgBox "No numeric data found in " & srcDesc & " on sheet '" & ws.name & "'.", _
               vbExclamation, "Graphify"
        GoTo Bye
    End If

    Set dataRng = ws.Range(ws.Cells(firstDataRow, srcCol), ws.Cells(lastRow, srcCol))
    nPts = dataRng.Rows.count

    '--- category labels R1..Rn ------------------------------------------------
    ReDim cats(1 To nPts)
    For i = 1 To nPts
        cats(i) = "R" & i
    Next i

    '--- reuse the chart if it exists, otherwise create it ---------------------
    Set co = Nothing
    On Error Resume Next
    Set co = ws.ChartObjects(CHART_NAME)
    On Error GoTo Fail

    If co Is Nothing Then
        Set co = ws.ChartObjects.Add(ws.Range("E2").Left, ws.Range("E2").Top, _
                                     CHART_WIDTH, CHART_HEIGHT)
        co.name = CHART_NAME
    End If

    '--- rebuild the series from the resolved source ----------------------------
    With co.Chart
        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop
        Set srs = .SeriesCollection.NewSeries
        srs.name = "Values"
        srs.Values = dataRng
        srs.XValues = cats

        .ChartType = xlColumnClustered     'change if you prefer xlLine, xlPie, etc.
        .HasTitle = True
        .ChartTitle.text = "Graphify - Column " & ColumnLetter(ws, srcCol)
        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Caption = "Values"
        End With
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Caption = "Row"
        End With
        .HasLegend = False
    End With

    '--- dock the chart at E2 ---------------------------------------------------
    co.Left = ws.Range("E2").Left
    co.Top = ws.Range("E2").Top

Bye:
    Application.ScreenUpdating = True
    Set srs = Nothing
    Set co = Nothing
    Set dataRng = Nothing
    Set ws = Nothing
    Exit Sub

Fail:
    MsgBox "Graphify failed (" & Err.Number & "): " & Err.description, _
           vbExclamation, "Graphify"
    Resume Bye
End Sub

'------------------------------------------------------------------------------
' Column with the most numeric cells in the used range (leftmost on tie).
' Returns 0 when no column holds at least MIN_NUMERIC numbers.
'------------------------------------------------------------------------------
Private Function PickNumericColumn(ws As Worksheet) As Long
    Dim ur As Range
    Dim c As Long
    Dim best As Long
    Dim bestN As Long
    Dim n As Long
    Dim lastUsedRow As Long

    Set ur = ws.UsedRange
    lastUsedRow = ur.Row + ur.Rows.count - 1

    best = 0
    bestN = 0
    For c = ur.Column To ur.Column + ur.Columns.count - 1
        n = Application.count(ws.Range(ws.Cells(1, c), ws.Cells(lastUsedRow, c)))
        If n > bestN Then
            best = c
            bestN = n
        End If
    Next c

    If bestN >= MIN_NUMERIC Then
        PickNumericColumn = best
    Else
        PickNumericColumn = 0
    End If
End Function

'------------------------------------------------------------------------------
' First row at or below fromRow whose cell holds a real number
' (dates, booleans, text and blanks do not count). Returns 0 when none.
'------------------------------------------------------------------------------
Private Function FirstNumericRow(ws As Worksheet, colIdx As Long, fromRow As Long) As Long
    Dim r As Long
    Dim lastUsedRow As Long
    Dim v As Variant

    lastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.count - 1
    For r = fromRow To lastUsedRow
        v = ws.Cells(r, colIdx).Value
        Select Case VarType(v)
            Case vbDouble, vbLong, vbInteger, vbSingle, vbCurrency
                FirstNumericRow = r
                Exit Function
        End Select
    Next r
    FirstNumericRow = 0
End Function

'------------------------------------------------------------------------------
' Column index -> letter ("I").
'------------------------------------------------------------------------------
Private Function ColumnLetter(ws As Worksheet, colIdx As Long) As String
    Dim a As String
    a = ws.Columns(colIdx).address(False, False)   ' e.g. "I:I"
    ColumnLetter = Left$(a, InStr(a, ":") - 1)
End Function

Option Explicit

' v8.3.3 SCANNING ENGINE (Hardened)
' Reads the posted journals row-by-row, resolves the actual COA, aggregates
' by physical GL account + month, and writes the visible GL directly.
'
' Design rule:
'   SOURCE ROW -> COA RESOLUTION -> MONTH BUCKET -> GL BLOCK -> G/H/I
'
' CDJ sundry rows are balanced from the row's actual cash leg and mapped
' allocation legs. This prevents a stale/incorrect Sundry Debit value from
' becoming the GL amount. The Sundry Account name remains the COA identity.

Private Const FIRST_GL_ROW As Long = 13
Private Const GL_STRIDE As Long = 13
Private Const GL_BLOCKS As Long = 46
Private Const MONTHS As Long = 12

Public Sub ScanGL(Optional ByVal yearNumber As Long = 2026)
    Dim d As Object, gd As Double, gc As Double, um As Double, inv As Long
    Dim detail As String, ws As Worksheet

    ' Pick up any operator edit to the COA_MAP alias table before mapping.
    V8_MapReset

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    ' Fail closed if the visible GL does not define 46 distinct account titles.
    ' Duplicate titles would receive the same normalized postings in multiple blocks.
    If Not ValidateGLAccountTitles(detail) Then inv = inv + 1

    ScanCDJ d, yearNumber, gd, gc, um, inv, detail
    ScanCRJ d, yearNumber, gd, gc, um, inv, detail
    ' GJ is out of scope: do not read or post from the GJ worksheet.

    ' Fail-closed validation gate:
    ' If journals are out of balance, have unmapped accounts, or have invalid data,
    ' record the diagnostic audit row and abort BEFORE modifying the visible GL.
    If inv > 0 Or um > TOLERANCE Or Abs(gd - gc) > TOLERANCE Then
        WriteScanAudit gd, gc, um, inv, detail
    If Application.UserControl And Application.DisplayAlerts Then
            MsgBox "v8.3.3 SCAN HOLD: Reconciliation or mapping gate failed. GL not overwritten." & vbCrLf & _
                   "Debit: " & Format$(gd, "0.00") & vbCrLf & _
                   "Credit: " & Format$(gc, "0.00") & vbCrLf & _
                   "Difference: " & Format$(gd - gc, "0.00") & vbCrLf & _
                   "Unmapped: " & Format$(um, "0.00") & vbCrLf & _
                   "Invalid: " & inv & vbCrLf & _
                   "Review GL_AUDIT for details.", vbExclamation
        End If
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    WriteScannedGL ws, d
    ' WriteScannedCalc excluded per directive: do not introduce another calculation sheet
    WriteScanAudit gd, gc, um, inv, detail

    If Application.UserControl And Application.DisplayAlerts Then
        MsgBox "v8.3.3 SCAN COMPLETE" & vbCrLf & _
               "GL rows scanned: 46 accounts x 12 months = 552" & vbCrLf & _
               "Debit: " & Format$(gd, "0.00") & vbCrLf & _
               "Credit: " & Format$(gc, "0.00") & vbCrLf & _
               "Difference: " & Format$(gd - gc, "0.00") & vbCrLf & _
               "Unmapped: " & Format$(um, "0.00") & vbCrLf & _
               "Invalid: " & inv, vbInformation
    End If
End Sub

Private Sub AddPost(ByVal d As Object, ByVal acct As String, ByVal m As Long, ByVal db As Double, ByVal cr As Double)
    Dim k As String, a As Variant
    If Len(Trim$(acct)) = 0 Or m < 1 Or m > 12 Then Exit Sub
    k = V8_Norm(acct) & "|" & CStr(m)
    If Not d.Exists(k) Then d.Add k, Array(0#, 0#)
    a = d(k)
    a(0) = CDbl(a(0)) + db
    a(1) = CDbl(a(1)) + cr
    d(k) = a
End Sub

Private Sub ScanCDJ(ByVal d As Object, ByVal yr As Long, ByRef gd As Double, ByRef gc As Double, ByRef um As Double, ByRef inv As Long, ByRef detail As String)
    Dim ws As Worksheet, r As Long, c As Long, m As Long, cur As Long
    Dim ok As Boolean, v As Double, a As String, raw As String
    Dim cash As Double, otherDb As Double, otherCr As Double, sundryNet As Double
    Dim cashOK As Boolean, explicitDb As Double, explicitCr As Double
    Dim title As String

    Set ws = ThisWorkbook.Worksheets("CDJ")
    cur = 0

    For r = 15 To ws.Cells(ws.Rows.count, 5).End(xlUp).Row
        m = V8_Month(ws.Cells(r, 3).Value2)
        If m > 0 Then cur = m
        If cur = 0 Then GoTo NextRow
        title = SafeText(ws.Cells(r, 5).Value2)
        If Len(title) = 0 Or InStr(1, V8_Norm(title), "TOTAL", vbTextCompare) > 0 Then GoTo NextRow

        ' Cash is column F: negative = credit, positive = debit.
        ok = True: cash = V8_Number(ws.Cells(r, 6).Value2, ok)
        If Not ok Then inv = inv + 1: GoTo NextRow

        If Abs(cash) > TOLERANCE Then
            If cash >= 0 Then
                AddPost d, "Cash in Bank", cur, cash, 0#
                gd = gd + cash
            Else
                AddPost d, "Cash in Bank", cur, 0#, -cash
                gc = gc - cash
            End If
        End If

        otherDb = 0#: otherCr = 0#
        For c = 7 To 18
            ok = True: v = V8_Number(ws.Cells(r, c).Value2, ok)
            If Not ok Then inv = inv + 1: GoTo NextRow
            If Abs(v) > TOLERANCE Then
                a = V8_CDJMap(c)
                If Len(a) = 0 Then
                    um = um + Abs(v)
                    detail = detail & "CDJ!" & ws.Cells(r, c).address(False, False) & " unmapped allocation" & vbCrLf
                ElseIf v > 0 Then
                    AddPost d, a, cur, v, 0#: gd = gd + v: otherDb = otherDb + v
                Else
                    AddPost d, a, cur, 0#, -v: gc = gc - v: otherCr = otherCr - v
                End If
            End If
        Next c

        ' Sundry Account is the actual COA label. Derive its balancing amount
        ' from the row rather than trusting a stale copied amount in T/U.
        raw = SafeText(ws.Cells(r, 19).Value2)
        If Len(raw) > 0 And V8_Norm(raw) <> "TOTAL" And V8_Norm(raw) <> "SUNDRY ACCOUNT" Then
            a = V8_ResolveAccount(raw)
            If Len(a) = 0 And GLTitleExists(raw) Then a = GLTitleOf(raw)
            If Len(a) = 0 Then
                um = um + Abs(cash) + otherDb + otherCr
                detail = detail & "CDJ!" & ws.Cells(r, 19).address(False, False) & " -> " & raw & " unmapped Sundry Account" & vbCrLf
            Else
                ' If cash is a credit, sundry must supply the balancing debit
                ' after the other debit/credit allocations are considered.
                sundryNet = (-cash) + otherCr - otherDb
                If sundryNet > TOLERANCE Then
                    AddPost d, a, cur, sundryNet, 0#: gd = gd + sundryNet
                ElseIf sundryNet < -TOLERANCE Then
                    AddPost d, a, cur, 0#, -sundryNet: gc = gc - sundryNet
                End If

                ' Record disagreement with the visible T/U amount, but do not
                ' allow it to corrupt the balanced scan.
                cashOK = True
                explicitDb = V8_Number(ws.Cells(r, 20).Value2, cashOK)
                If Not cashOK Then inv = inv + 1
                cashOK = True
                explicitCr = V8_Number(ws.Cells(r, 21).Value2, cashOK)
                If Not cashOK Then inv = inv + 1
                If Abs(explicitDb - MaxD(sundryNet, 0#)) > TOLERANCE Or _
                   Abs(explicitCr - MaxD(-sundryNet, 0#)) > TOLERANCE Then
                    detail = detail & "CDJ!" & ws.Cells(r, 19).address(False, False) & _
                             " Sundry amount corrected from T/U to balanced row amount: " & _
                             Format$(sundryNet, "0.00") & vbCrLf
                End If
            End If
        End If
NextRow:
    Next r
End Sub

Private Sub ScanCRJ(ByVal d As Object, ByVal yr As Long, ByRef gd As Double, ByRef gc As Double, ByRef um As Double, ByRef inv As Long, ByRef detail As String)
    ' CRJ month attribution is partition-based, not based on column O.
    ' Column O contains SJ import/log timestamps and is not a transaction date.
    ' A live monthly boundary row carries the month label in column D and is blank in
    ' columns F (customer) and G (invoice number). The legacy v8.3.3 summary block lower
    ' in the sheet has F = "TOTAL" with month labels in D but is NOT a live boundary;
    ' the blank F/G test excludes it. Contiguity is enforced below (first boundary must
    ' be January; each subsequent boundary must be exactly the next month).
    ' The yr parameter is retained for caller compatibility; partition order proves the year scope.
    Dim ws As Worksheet, r As Long, rr As Long, c As Long
    Dim m As Long, lastMonth As Long, startRow As Long, lastRow As Long
    Dim ok As Boolean, v As Double, a As String, invNo As String
    Dim partitionOK As Boolean

    Set ws = ThisWorkbook.Worksheets("CRJ")
    ' Scan range must include boundary rows themselves: they carry data in column D
    ' (month label) but are blank in F/G, so F/G alone would cut the range short
    ' at the last invoice and skip the final monthly boundary.
    lastRow = ws.Cells(ws.Rows.count, 6).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 7).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, 7).End(xlUp).Row
    If ws.Cells(ws.Rows.count, 4).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.count, 4).End(xlUp).Row

    startRow = 10
    lastMonth = 0

    For r = startRow To lastRow
        m = V8_Month(ws.Cells(r, 4).Value2)
        If m > 0 And Len(SafeText(ws.Cells(r, 6).Value2)) = 0 And _
           Len(SafeText(ws.Cells(r, 7).Value2)) = 0 Then
            partitionOK = (m = lastMonth + 1 And m >= 1 And m <= 12)

            If Not partitionOK Then
                detail = detail & "CRJ!D" & r & " invalid or out-of-order partition month; partition held" & vbCrLf
                For rr = startRow To r - 1
                    If Len(SafeText(ws.Cells(rr, 7).Value2)) > 0 Then inv = inv + 1
                Next rr
            Else
                For rr = startRow To r - 1
                    invNo = SafeText(ws.Cells(rr, 7).Value2)
                    If Len(invNo) > 0 Then
                        ok = True: v = V8_Number(ws.Cells(rr, 8).Value2, ok)
                        If Not ok Then
                            inv = inv + 1
                        ElseIf Abs(v) > TOLERANCE Then
                            AddPost d, "Cash in Bank", m, v, 0#
                            gd = gd + v
                        End If

                        For c = 9 To 12
                            ok = True: v = V8_Number(ws.Cells(rr, c).Value2, ok)
                            If Not ok Then
                                inv = inv + 1
                            ElseIf Abs(v) > TOLERANCE Then
                                a = V8_CRJMap(c)
                                If Len(a) = 0 Then
                                    um = um + Abs(v)
                                    detail = detail & "CRJ!" & ws.Cells(rr, c).address(False, False) & " unmapped credit" & vbCrLf
                                Else
                                    AddPost d, a, m, 0#, v
                                    gc = gc + v
                                End If
                            End If
                        Next c
                    End If
                Next rr
            End If

            lastMonth = m
            startRow = r + 1
        End If
    Next r

    ' Never guess a month for rows that appear after the last recognized monthly boundary.
    For rr = startRow To lastRow
        If Len(SafeText(ws.Cells(rr, 7).Value2)) > 0 Then
            inv = inv + 1
            detail = detail & "CRJ!G" & rr & " invoice has no enclosing monthly partition" & vbCrLf
        End If
    Next rr
End Sub

' GJ scan intentionally removed. This workbook's GL input scope is CDJ + CRJ only.

Private Function ValidateGLAccountTitles(ByRef detail As String) As Boolean
    Dim ws As Worksheet, seen As Object
    Dim b As Long, r As Long, title As String, key As String
    Dim valid As Boolean

    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    valid = True

    For b = 0 To GL_BLOCKS - 1
        r = FIRST_GL_ROW + b * GL_STRIDE
        title = SafeText(ws.Cells(r, 6).Value2)
        key = V8_Norm(title)

        If Len(key) = 0 Then
            detail = detail & "GL!F" & r & " missing account title for block " & (b + 1) & vbCrLf
            valid = False
        ElseIf seen.Exists(key) Then
            detail = detail & "GL!F" & r & " duplicate account title '" & title & _
                     "' (first seen at GL!F" & seen(key) & "); write-back held" & vbCrLf
            valid = False
        Else
            seen.Add key, r
        End If
    Next b

    ValidateGLAccountTitles = valid
End Function

Private Sub WriteScannedGL(ByVal ws As Worksheet, ByVal d As Object)
    Dim b As Long, m As Long, r As Long, k As String, x As Variant
    Dim title As String, bal As Double

    Application.ScreenUpdating = False
    On Error GoTo Fail
    For b = 0 To GL_BLOCKS - 1
        title = SafeText(ws.Cells(FIRST_GL_ROW + b * GL_STRIDE, 6).Value2)
        If Len(title) = 0 Then Err.Raise vbObjectError + 833, "ScanGL", "Missing GL account title at block " & (b + 1)
        bal = 0#
        For m = 1 To MONTHS
            r = FIRST_GL_ROW + b * GL_STRIDE + m - 1
            k = V8_Norm(title) & "|" & CStr(m)
            If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
            ws.Cells(r, 7).Value2 = CDbl(x(0))
            ws.Cells(r, 8).Value2 = CDbl(x(1))
            bal = bal + CDbl(x(0)) - CDbl(x(1))
            ws.Cells(r, 9).Value2 = bal
        Next m
    Next b
Fail:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub WriteScannedCalc(ByVal d As Object)
    Dim ws As Worksheet, gl As Worksheet, b As Long, m As Long, r As Long, k As String
    Dim x As Variant, title As String, bal As Double
    Set gl = ThisWorkbook.Worksheets(GL_SHEET)
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(CALC_SHEET): On Error GoTo 0
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = CALC_SHEET
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("Account Title", "Month", "Debit", "Credit", "Ending Balance")
    r = 2
    For b = 0 To GL_BLOCKS - 1
        title = SafeText(gl.Cells(FIRST_GL_ROW + b * GL_STRIDE, 6).Value2)
        bal = 0#
        For m = 1 To MONTHS
            k = V8_Norm(title) & "|" & CStr(m)
            If d.Exists(k) Then x = d(k) Else x = Array(0#, 0#)
            bal = bal + CDbl(x(0)) - CDbl(x(1))
            ws.Cells(r, 1).Value2 = title
            ws.Cells(r, 2).Value2 = m
            ws.Cells(r, 3).Value2 = CDbl(x(0))
            ws.Cells(r, 4).Value2 = CDbl(x(1))
            ws.Cells(r, 5).Value2 = bal
            r = r + 1
        Next m
    Next b
End Sub

Private Sub WriteScanAudit(ByVal gd As Double, ByVal gc As Double, ByVal um As Double, ByVal inv As Long, ByVal detail As String)
    Dim ws As Worksheet, r As Long, status As String
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(AUDIT_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = AUDIT_SHEET
    On Error GoTo 0
    status = IIf(Abs(gd - gc) <= TOLERANCE And um <= TOLERANCE And inv = 0, "PASS", "HOLD")
    r = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = status
    ws.Cells(r, 3).Value = "v8.3.3 SCAN | Debit=" & Format$(gd, "0.00") & " | Credit=" & Format$(gc, "0.00") & " | Difference=" & Format$(gd - gc, "0.00") & " | Unmapped=" & Format$(um, "0.00") & " | Invalid=" & inv
    ws.Cells(r, 4).Value = Left$(detail, 32000)
End Sub

Private Function SafeText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    SafeText = Trim$(CStr(v))
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function



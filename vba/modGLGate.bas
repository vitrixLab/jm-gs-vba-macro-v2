Option Explicit

' v8.2 fail-closed gate. Exact workbook structure is validated before calculation.

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(name): On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function ValidateV8Structure() As Boolean
    Dim names As Variant, n As Variant, ws As Worksheet
    names = Array("CDJ", "CRJ", "GJ", "GL", "SUPPLIERS DATA")
    For Each n In names
        If Not SheetExists(CStr(n)) Then
            ValidateV8Structure = False
            Exit Function
        End If
    Next n

    Set ws = ThisWorkbook.Worksheets("CDJ")
    If V8_Norm(ws.Cells(13, 3).Value2) <> "DATE" Or V8_Norm(ws.Cells(14, 20).Value2) <> "DEBIT" Or V8_Norm(ws.Cells(14, 21).Value2) <> "CREDIT" Then Exit Function
    Set ws = ThisWorkbook.Worksheets("CRJ")
    If V8_Norm(ws.Cells(8, 4).Value2) <> "DATE" Or V8_Norm(ws.Cells(8, 8).Value2) <> "DEBIT" Then Exit Function
    Set ws = ThisWorkbook.Worksheets("GJ")
    If V8_Norm(ws.Cells(9, 2).Value2) <> "DATE" Or V8_Norm(ws.Cells(9, 6).Value2) <> "DEBIT" Or V8_Norm(ws.Cells(9, 7).Value2) <> "CREDIT" Then Exit Function
    Set ws = ThisWorkbook.Worksheets("GL")
    If V8_Norm(ws.Cells(11, 6).Value2) <> "ACCOUNT TITLE" Then Exit Function
    Dim b As Long, r As Long, n As Long
    For b = 0 To 45
        r = 13 + b * 13
        If Len(Trim$(CStr(ws.Cells(r, 6).Value2))) = 0 Then Exit Function
        n = n + 1
    Next b
    If n <> 46 Then Exit Function
    ValidateV8Structure = True
End Function

Public Sub RunV8Gate()
    If Not ValidateV8Structure() Then
        MsgBox "v8.2 gate HOLD: workbook structure does not match the reviewed mapping.", vbExclamation
        Exit Sub
    End If
    If ValidateGLConsistency(2026) Then
        If BuildV8CalcSheet(2026) Then
            MsgBox "v8.2 gate PASS: posting population reconciles and 46x12 matrix exists.", vbInformation
        Else
            MsgBox "v8.0 gate HOLD: matrix size is not 552.", vbExclamation
        End If
    Else
        MsgBox "v8.0 gate HOLD: review GL_AUDIT for unmapped or invalid activity.", vbExclamation
    End If
End Sub

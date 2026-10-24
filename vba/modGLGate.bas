Option Explicit

' v8.2 fail-closed gate. Exact workbook structure is validated before calculation.

Private Function SheetExists(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(name): On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Public Function ValidateV8Structure() As Boolean
    Dim names As Variant, n As Variant, ws As Worksheet
    names = Array("CDJ", "CRJ", "GL", "COA_MAP")
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
    ' GJ structure is deliberately not inspected; it is outside GL scan scope.
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
        If Application.UserControl And Application.DisplayAlerts Then _
            MsgBox "GL gate HOLD: CDJ/CRJ/GL/COA_MAP structure does not match the expected layout.", vbExclamation
        Exit Sub
    End If

    ' Use the current fail-closed CDJ + CRJ scan. Do not invoke the legacy
    ' GL_V8_CALC builder or inspect the excluded GJ worksheet.
    ScanGL 2026
End Sub

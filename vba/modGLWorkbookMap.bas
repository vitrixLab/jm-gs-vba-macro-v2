Option Explicit

' Account resolution for the Global-Smile 2026 workbook layout.
' Posting sources in this GL scan are CDJ and CRJ only. GJ is excluded.

Public Const AUDIT_SHEET As String = "GL_AUDIT"
Public Const CALC_SHEET As String = "GL_V8_CALC"
Public Const GL_SHEET As String = "GL"
Public Const COAMAP_SHEET As String = "COA_MAP"
Public Const TOLERANCE As Double = 0.01

' Module-level caches. The account resolver builds dictionaries once per scan
' and clears them through V8_MapReset.
Private mGLByNorm As Object      ' V8_Norm(title) -> title
Private mGLByKey As Object       ' V8_Key(title)  -> title
Private mAlias As Object         ' V8_Key(source label) -> GL title
Private mMapsBuilt As Boolean

' Drop every cached map so an edit to COA_MAP or GL takes effect on the next
' scan without reopening the workbook.
Public Sub V8_MapReset()
    Set mGLByNorm = Nothing
    Set mGLByKey = Nothing
    Set mAlias = Nothing
    mMapsBuilt = False
End Sub

' Punctuation- and spacing-insensitive key.  "Clinician's Fee-Corporators" and
' "Clinicians Fee- Corporators" both reduce to CLINICIANSFEECORPORATORS, which
' is how the nine form-only sundry labels find their real GL account.
Public Function V8_Key(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    s = UCase$(Replace(CStr(v), ChrW(160), " "))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            V8_Key = V8_Key & ch
        End If
    Next i
End Function

Private Sub EnsureMaps()
    Dim d As Object, ka As Object, al As Object
    Dim ws As Worksheet, r As Long, a As String, k As String, acct As String
    Dim st As String
    If mMapsBuilt Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ka = CreateObject("Scripting.Dictionary"): ka.CompareMode = vbTextCompare
    Set al = CreateObject("Scripting.Dictionary"): al.CompareMode = vbTextCompare

    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    For r = 13 To ws.Cells(ws.Rows.count, 6).End(xlUp).Row
        a = Trim$(CStr(ws.Cells(r, 6).Value2))
        If Len(a) > 0 Then
            If Not d.Exists(V8_Norm(a)) Then d.Add V8_Norm(a), a
            If Not ka.Exists(V8_Key(a)) Then ka.Add V8_Key(a), a
        End If
    Next r

    ' Aliases come from the COA_MAP sheet, which is the operator-editable
    ' single source of truth.  A row is only honoured when it names an
    ' account AND its status says MAPPED, so the seven labels held for
    ' accounting decision (blank account, status NEEDS ACCOUNTING DECISION)
    ' are deliberately not aliased and the gate keeps holding on them.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(COAMAP_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        r = 5
        For r = 5 To ws.Cells(ws.Rows.count, 1).End(xlUp).Row
            acct = Trim$(CStr(ws.Cells(r, 4).Value2))
            st = UCase$(Trim$(CStr(ws.Cells(r, 6).Value2)))
            If Len(acct) > 0 And InStr(st, "MAPPED") > 0 Then
                k = V8_Key(ws.Cells(r, 1).Value2)
                If Len(k) > 0 And Not al.Exists(k) Then al.Add k, acct
            End If
            ' (increment handled by For loop)
        Next r
    End If

    Set mGLByNorm = d
    Set mGLByKey = ka
    Set mAlias = al
    mMapsBuilt = True
End Sub

Public Function V8_Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    V8_Norm = UCase$(Application.WorksheetFunction.Trim(Replace(CStr(v), ChrW(160), " ")))
End Function

Public Function V8_Month(ByVal v As Variant) As Long
    Select Case V8_Norm(v)
        Case "JAN", "JANUARY": V8_Month = 1
        Case "FEB", "FEBRUARY": V8_Month = 2
        Case "MAR", "MARCH": V8_Month = 3
        Case "APR", "APRIL": V8_Month = 4
        Case "MAY": V8_Month = 5
        Case "JUN", "JUNE": V8_Month = 6
        Case "JUL", "JULY": V8_Month = 7
        Case "AUG", "AUGUST": V8_Month = 8
        Case "SEP", "SEPTEMBER": V8_Month = 9
        Case "OCT", "OCTOBER": V8_Month = 10
        Case "NOV", "NOVEMBER": V8_Month = 11
        Case "DEC", "DECEMBER": V8_Month = 12
    End Select
End Function

Public Function V8_GLAccounts() As Object
    Dim d As Object, ws As Worksheet, r As Long, a As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ws = ThisWorkbook.Worksheets(GL_SHEET)
    For r = 13 To ws.Cells(ws.Rows.count, 6).End(xlUp).Row
        a = V8_Norm(ws.Cells(r, 6).Value2)
        If Len(a) > 0 Then If Not d.Exists(a) Then d.Add a, CStr(ws.Cells(r, 6).Value2)
    Next r
    Set V8_GLAccounts = d
End Function

Public Function V8_ResolveAccount(ByVal raw As String) As String
    Dim k As String, d As Object, kd As String
    k = V8_Norm(raw): kd = V8_Key(raw)
    EnsureMaps
    Set d = mGLByNorm
    ' Operator-editable aliases on COA_MAP win over the built-in table, so a
    ' held label can be resolved without shipping a new build.
    If Len(kd) > 0 Then
        If mAlias.Exists(kd) Then V8_ResolveAccount = mAlias(kd): Exit Function
    End If
    Select Case k
        Case "TRANSPORATION AND TRAVEL": V8_ResolveAccount = "Transportation and Travel"
        Case "ADVANCES TO EMPLOYEE": V8_ResolveAccount = "Advances to Employees"
        Case "REPAIRS AND MAINTENANCE": V8_ResolveAccount = "Repair and Maintenance"
        Case "OUTPUT VAT PAYABLE": V8_ResolveAccount = "VAT Payable"
        Case "CLINIC SUPPLIES": V8_ResolveAccount = "Clinic Material and Supplies"
        ' v8.1 HOLD mappings are now executable runtime mappings.
        Case "MEDICAL EQUIPMENT": V8_ResolveAccount = "Dental Equipment"
        Case "COST OF REVENUE": V8_ResolveAccount = "Clinic Material and Supplies"
        Case "SUPPLIES": V8_ResolveAccount = "Clinic Material and Supplies"
        Case "BANK CHARGE": V8_ResolveAccount = "Miscellaneous"
        Case "CHARGES": V8_ResolveAccount = "Miscellaneous"
        Case Else
            If d.Exists(k) Then
                V8_ResolveAccount = d(k)
            ElseIf Len(kd) > 0 And mGLByKey.Exists(kd) Then
                ' Same account, different punctuation or spacing in the journal
                ' label.  Safe because the exact match above was already tried.
                V8_ResolveAccount = mGLByKey(kd)
            End If
    End Select
End Function

Public Function V8_CDJMap(ByVal col As Long) As String
    Select Case col
        Case 6: V8_CDJMap = "Cash in Bank"
        Case 7: V8_CDJMap = "Input VAT"
        Case 8: V8_CDJMap = "EWT Payable"
        Case 9: V8_CDJMap = "Petty Cash Fund"
        Case 10: V8_CDJMap = "Government Contributions (EE)"
        Case 11: V8_CDJMap = "Government Loans (EE)"
        Case 12: V8_CDJMap = "De Minimis"
        Case 13: V8_CDJMap = "Salaries and Wages"
        Case 14: V8_CDJMap = "Clinic Material and Supplies"
        Case 15: V8_CDJMap = "Rent"
        Case 16: V8_CDJMap = "Gas, Oil, Parking, Toll Fees"
        Case 17: V8_CDJMap = "Light and Water Expense"
        Case 18: V8_CDJMap = "Professional Fees"
        Case Else: V8_CDJMap = ""
    End Select
End Function

Public Function GLTitleExists(ByVal raw As String) As Boolean
    Dim kd As String
    EnsureMaps
    kd = V8_Key(raw)
    GLTitleExists = mGLByNorm.Exists(V8_Norm(raw)) Or _
                    (Len(kd) > 0 And mGLByKey.Exists(kd))
End Function

Public Function GLTitleOf(ByVal raw As String) As String
    Dim k As String, kd As String
    EnsureMaps
    k = V8_Norm(raw): kd = V8_Key(raw)
    If mGLByNorm.Exists(k) Then
        GLTitleOf = mGLByNorm(k)
    ElseIf Len(kd) > 0 And mGLByKey.Exists(kd) Then
        GLTitleOf = mGLByKey(kd)
    End If
End Function

Public Function V8_CRJMap(ByVal col As Long) As String
    Select Case col
        Case 8: V8_CRJMap = "Cash in Bank"
        Case 9: V8_CRJMap = "Excess CWT Over IT"
        Case 10: V8_CRJMap = "VAT Payable"
        Case 11, 12: V8_CRJMap = "Sales"
        Case Else: V8_CRJMap = ""
    End Select
End Function

Public Function V8_Number(ByVal v As Variant, ByRef ok As Boolean) As Double
    ok = True
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then Exit Function
    If Not IsNumeric(v) Then ok = False Else V8_Number = CDbl(v)
End Function



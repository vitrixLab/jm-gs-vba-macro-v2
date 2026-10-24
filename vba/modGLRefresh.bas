Option Explicit

' v8.3.3 compatibility wrapper.
' The old RefreshGL matrix writer is retired from the execution path.
' RefreshGL delegates to the row-scanning engine so the visible GL,
' GL_V8_CALC, and GL_AUDIT all come from the same scanned postings.
'
' IMPORTANT: Macros intended for the Excel Macro Dialog (Alt+F8) MUST NOT
' take any parameters (even Optional ones). Subroutines below are zero-argument
' wrappers to ensure complete visibility in Alt+F8.

Public Sub RefreshGL()
    ScanGL 2026
End Sub

Public Sub RefreshAllGL()
    ScanGL 2026
End Sub

Public Sub RefreshGLIntoSheet()
    ScanGL 2026
End Sub

Public Sub RefreshGLForYear(ByVal yearNumber As Long)
    ScanGL yearNumber
End Sub

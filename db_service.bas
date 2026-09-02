Option Compare Database
Option Explicit

' =========================================================
' ماژول: db_service.bas
' =========================================================
' (existing content kept) + added helper for print attempt increment
' =========================================================

' ... (rest of existing functions) ...

' =========================================================
' تابع: db_SetPrintAttemptStarted
' =========================================================
'
' وظیفه:
' ثبت شروع تلاش چاپ: افزایش PrintAttempts و ثبت LastPrintAttempt' 
Public Function db_SetPrintAttemptStarted(ByVal recordID As Long) As Boolean
    On Error GoTo ErrHandlern        Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim pa As Longn        If recordID <= 0 Then Exit Function
        Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT PrintAttempts, LastPrintAttempt FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recordID), dbOpenDynaset)
    If rs.EOF Then GoTo CleanExit
        pa = Nz(rs!PrintAttempts, 0) + 1
    If pa < 0 Then pa = 0
    If pa > MAX_PRINT_ATTEMPTS Then pa = MAX_PRINT_ATTEMPTS
        rs.Edit
    rs!LastPrintAttempt = Now()
    rs!PrintAttempts = pa
    rs.Update
        db_SetPrintAttemptStarted = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    Call LogError("db_SetPrintAttemptStarted", Err.Number, Err.Description, CStr(recordID))
    db_SetPrintAttemptStarted = False
    Resume CleanExit
End Function

Option Compare Database
Option Explicit

' =========================================================
' ماژول: db_service.bas
' =========================================================
' سرویس مرکزی دسترسی به دیتابیس: توابع کمکی برای به‌روزرسانی وضعیت‌ها و مهاجرت schema
' هدف: ارائه API اتمیک برای بقیهٔ ماژول‌ها تا منطق DB پراکنده نشود.
' =========================================================

Public Function db_migration_CreateOrUpdateSchema() As Boolean
    On Error GoTo ErrHandler
    ' پیاده‌سازی ساده: اطمینان از وجود جداول موردنیاز        Dim db As DAO.Database
    Set db = CurrentDb()
    ' این تابع فرض می‌کند جداول از قبل وجود دارند یا از طریق ابزار مدیر ساخته می‌شوند.
    db_migration_CreateOrUpdateSchema = True
    Exit Function
ErrHandler:
    Call LogError("db_migration_CreateOrUpdateSchema", Err.Number, Err.Description, "")
    db_migration_CreateOrUpdateSchema = False
End Function

' =========================================================
' تابع: db_UpdateRecordState
' =========================================================
' وظیفه: به‌صورت اتمیک وضعیت پردازش یک رکورد تردد را به‌روز می‌کند.
Public Function db_UpdateRecordState(ByVal recordID As Long, ByVal newState As String, _
                                      ByVal processedOK As Boolean, ByVal printAttempts As Long, _
                                      ByVal processedDate As Variant, ByVal note As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
        If recordID <= 0 Then Exit Function
    Set db = CurrentDb()
    sql = "SELECT ProcessingResult, PrintAttempts, ProcessedDate, Note FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recordID)
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    If rs.EOF Then GoTo CleanExit
    rs.Edit
    rs!ProcessingResult = newState
    rs!PrintAttempts = printAttempts
    If Not IsMissing(processedDate) Then
        If Not IsNull(processedDate) Then rs!ProcessedDate = processedDate
    End If
    If Len(Trim$(note)) > 0 Then rs!Note = note
    rs.Update
    db_UpdateRecordState = True
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    Call LogError("db_UpdateRecordState", Err.Number, Err.Description, CStr(recordID))
    db_UpdateRecordState = False
    Resume CleanExit
End Function

' =========================================================
' تابع: db_SetPrintAttemptStarted' 
' ضبط شروع تلاش چاپ: افزایش PrintAttempts و ثبت LastPrintAttempt به صورت اتمیک
Public Function db_SetPrintAttemptStarted(ByVal recordID As Long) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim pa As Long
        If recordID <= 0 Then Exit Function
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT PrintAttempts, LastPrintAttempt FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recordID), dbOpenDynaset)
    If rs.EOF Then GoTo CleanExit
    pa = Nz(rs!PrintAttempts, 0) + 1    If pa < 0 Then pa = 0
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

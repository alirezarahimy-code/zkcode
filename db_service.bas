Option Compare Database
Option Explicit

' =========================================================
' ماژول: db_service.bas
' =========================================================
'
' توضیح ماژول:
' این ماژول عملیات Atomic و Thread-Safe بر روی پایگاه داده را انجام می‌دهد.
' تمام تردد‌ها به صورت Atomic و یکتا پردازش می‌شوند.
'
' کاربرد:
' - Atomic Claim (ادعا) رکوردها برای پردازش
' - جلوگیری از پردازش تکراری
' - بروزرسانی وضعیت و State Machine
' - معامله (Transaction) امن
'
' ویژگی‌های مهم:
' - Optimistic Locking (بدون Lock فیزیکی)
' - State Machine Pattern (حالات تعریف‌شده)
' - Retry Logic (تلاش دوباره)
' - Recovery Mechanism (بازیابی از خرابی)
'
' معماری:
' - db_TryClaimRecordAtomic: ادعای یک رکورد برای پردازش
' - db_UpdateRecordState: بروزرسانی وضعیت رکورد
' - db_GetRecordState: دریافت وضعیت رکورد
' - db_RollbackRecord: بازگرداندن رکورد به حالت قبل
'
' =========================================================

' =========================================================
' تابع: db_TryClaimRecordAtomic
' =========================================================
'
' وظیفه:
' یک رکورد تردد را برای پردازش Claim می‌کند
' اگر رکورد قبلاً Claim شده باشد، بازمی‌گردد False
' (Optimistic Locking)
'
' پارامترها:
'   recordID (Long): شناسه رکورد تردد
'   newState (String): حالت جدید (معمولاً "PROCESSING")
'   allowedStates (String): حالات مجاز برای Claim
'                          فرمت: "'STATE1','STATE2','STATE3'"
'
' خروجی:
'   Boolean: True اگر Claim موفق، False اگر ناموفق
'
' نمونه استفاده:
'   If db_TryClaimRecordAtomic(123, "PROCESSING", _
'       "'NEW','PRINT_FAILED','NO_EMPLOYEE'") Then
'       ' رکورد Claim شد، می‌توان پردازش کرد
'       Call proc_ProcessRecordAtomic(123)
'   End If
'
' فرآیند:
'   1. بررسی وجود رکورد
'   2. دریافت وضعیت فعلی
'   3. بررسی: وضعیت فعلی در allowedStates است؟
'   4. اگر بله: تغییر به newState
'   5. اگر خیر: بازگردان False
'
' نکات مهم:
' - تنها یک بار در Batch پردازش
' - اگر موفق: وضعیت به PROCESSING تغییر می‌یابد
' - اگر ناموفق: وضعیت بدون تغییر باقی می‌ماند
' - Single-threaded (Access)، بنابراین Atomic است
'
' =========================================================

Public Function db_TryClaimRecordAtomic(ByVal recordID As Long, ByVal newState As String, _
                                       ByVal allowedStates As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim currentState As String
    
    If recordID <= 0 Then Exit Function
    If Len(Trim$(newState)) = 0 Then Exit Function
    If Len(Trim$(allowedStates)) = 0 Then Exit Function
    
    Set db = CurrentDb()
    
    ' دریافت رکورد برای بروزرسانی
    sql = "SELECT ProcessingResult FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE RecordID=" & CStr(recordID)
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    
    If rs.EOF Then
        ' رکورد موجود نیست
        Exit Function
    End If
    
    ' دریافت وضعیت فعلی
    currentState = Nz(rs!ProcessingResult, "NEW")
    
    ' بررسی: وضعیت فعلی در allowedStates است؟
    If Not IsStateAllowed(currentState, allowedStates) Then
        ' وضعیت فعلی مجاز نیست
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
    ' بروزرسانی وضعیت
    rs.Edit
    rs!ProcessingResult = newState
    rs!IsProcessed = False
    rs.Update
    
    db_TryClaimRecordAtomic = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("db_TryClaimRecordAtomic", Err.Number, Err.Description, CStr(recordID))
    db_TryClaimRecordAtomic = False
    Resume CleanExit
End Function

' =========================================================
' تابع: db_UpdateRecordState
' =========================================================
'
' وظیفه:
' وضعیت یک رکورد را بروزرسانی می‌کند
' همراه با اطلاعات اضافی (تلاش‌ها، تاریخ آخرین محاولهٌ و غیره)
'
' پارامترها:
'   recordID (Long): شناسه رکورد
'   newState (String): وضعیت جدید
'   isProcessed (Boolean): آیا پردازش تکمیل شده؟
'   printAttempts (Long): تعداد تلاش‌های چاپ
'   lastPrintAttempt (Date): تاریخ آخرین تلاش
'   errorMessage (String): پیام خطا (اختیاری)
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' =========================================================

Public Function db_UpdateRecordState(ByVal recordID As Long, ByVal newState As String, _
                                     Optional ByVal isProcessed As Boolean = False, _
                                     Optional ByVal printAttempts As Long = -1, _
                                     Optional ByVal lastPrintAttempt As Date, _
                                     Optional ByVal errorMessage As String = "") As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    If recordID <= 0 Then Exit Function
    
    Set db = CurrentDb()
    
    sql = "SELECT ProcessingResult, IsProcessed, PrintAttempts, LastPrintAttempt " & _
          "FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recordID)
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    
    If rs.EOF Then
        Exit Function
    End If
    
    rs.Edit
    rs!ProcessingResult = newState
    rs!IsProcessed = isProcessed
    
    If printAttempts >= 0 Then
        rs!PrintAttempts = printAttempts
    End If
    
    If lastPrintAttempt <> 0 Then
        rs!LastPrintAttempt = lastPrintAttempt
    End If
    
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
' تابع: db_GetRecordState
' =========================================================
'
' وظیفه:
' وضعیت فعلی یک رکورد را برمی‌گرداند
'
' پارامتر:
'   recordID (Long): شناسه رکورد
'
' خروجی:
'   String: وضعیت رکورد یا خالی اگر موجود نباشد
'
' =========================================================

Public Function db_GetRecordState(ByVal recordID As Long) As String
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    If recordID <= 0 Then Exit Function
    
    Set db = CurrentDb()
    
    sql = "SELECT ProcessingResult FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE RecordID=" & CStr(recordID)
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        db_GetRecordState = Nz(rs!ProcessingResult, "NEW")
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    db_GetRecordState = ""
End Function

' =========================================================
' تابع: db_RollbackRecord
' =========================================================
'
' وظیفه:
' یک رکورد را به حالت قبل بازمی‌گرداند
' در صورت خرابی یا نیاز به تلاش دوباره
'
' پارامترها:
'   recordID (Long): شناسه رکورد
'   previousState (String): حالت قبلی
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' =========================================================

Public Function db_RollbackRecord(ByVal recordID As Long, ByVal previousState As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    If recordID <= 0 Then Exit Function
    
    Set db = CurrentDb()
    
    sql = "SELECT ProcessingResult FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE RecordID=" & CStr(recordID)
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    
    If rs.EOF Then
        Exit Function
    End If
    
    rs.Edit
    rs!ProcessingResult = previousState
    rs!IsProcessed = False
    rs.Update
    
    db_RollbackRecord = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("db_RollbackRecord", Err.Number, Err.Description, CStr(recordID))
    db_RollbackRecord = False
    Resume CleanExit
End Function

' =========================================================
' تابع: db_GetPendingCount
' =========================================================
'
' وظیفه:
' تعداد رکوردهای درانتظار پردازش را برمی‌گرداند
' برای نمایش در UI
'
' خروجی:
'   Long: تعداد رکوردهای درانتظار
'
' =========================================================

Public Function db_GetPendingCount() As Long
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim count As Long
    
    Set db = CurrentDb()
    
    sql = "SELECT COUNT(*) as cnt FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE Nz(ProcessingResult,'NEW') IN " & _
          "('NEW','PRINT_FAILED','NO_EMPLOYEE','WAITING_FOR_MEAL','RECEIPT_FAILED','MEAL_FINALIZE_FAILED') " & _
          "AND Nz(PrintAttempts,0) < " & CStr(MAX_PRINT_ATTEMPTS)
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        count = Nz(rs!cnt, 0)
    End If
    
    db_GetPendingCount = count
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    db_GetPendingCount = 0
End Function

' =========================================================
' تابع: db_GetProcessingCount
' =========================================================
'
' وظیفه:
' تعداد رکوردهای در حال پردازش را برمی‌گرداند
'
' خروجی:
'   Long: تعداد رکوردهای درحال پردازش
'
' =========================================================

Public Function db_GetProcessingCount() As Long
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim count As Long
    
    Set db = CurrentDb()
    
    sql = "SELECT COUNT(*) as cnt FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE Nz(ProcessingResult,'')='PROCESSING'"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        count = Nz(rs!cnt, 0)
    End If
    
    db_GetProcessingCount = count
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    db_GetProcessingCount = 0
End Function

' =========================================================
' تابع کمکی: IsStateAllowed
' =========================================================
'
' وظیفه:
' بررسی می‌کند که آیا وضعیت کنونی در لیست مجاز است یا نه
'
' پارامترها:
'   currentState (String): وضعیت فعلی
'   allowedStates (String): لیست وضعیت‌های مجاز
'                          فرمت: "'STATE1','STATE2'"
'
' خروجی:
'   Boolean: True اگر مجاز، False اگر ممنوع
'
' =========================================================

Private Function IsStateAllowed(ByVal currentState As String, ByVal allowedStates As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim pattern As String
    
    ' فرمت: 'STATE1','STATE2'
    ' جستجو برای: 'CURRENTSTATE'
    
    currentState = Trim$(currentState)
    pattern = "'" & Replace(currentState, "'", "''") & "'"
    
    IsStateAllowed = (InStr(1, allowedStates, pattern, vbTextCompare) > 0)
    
    Exit Function
    
ErrHandler:
    IsStateAllowed = False
End Function

' =========================================================
' تابع: db_GetRecordDetails
' =========================================================
'
' وظیفه:
' تمام جزئیات یک رکورد تردد را برمی‌گرداند
'
' پارامتر:
'   recordID (Long): شناسه رکورد
'
' خروجی:
'   Variant: آرایه‌ای با جزئیات یا Null اگر موجود نباشد
'
' =========================================================

Public Function db_GetRecordDetails(ByVal recordID As Long) As Variant
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim details As Variant
    
    If recordID <= 0 Then Exit Function
    
    Set db = CurrentDb()
    
    sql = "SELECT RecordID, DeviceEnrollID, AttendanceDateTime, DeviceIP, " & _
          "DevicePort, DeviceMachineNumber, ProcessingResult, IsProcessed, " & _
          "PrintAttempts, LastPrintAttempt, CreatedDate " & _
          "FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recordID)
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ReDim details(0 To 10)
        details(0) = rs!RecordID
        details(1) = rs!DeviceEnrollID
        details(2) = rs!AttendanceDateTime
        details(3) = rs!DeviceIP
        details(4) = rs!DevicePort
        details(5) = rs!DeviceMachineNumber
        details(6) = rs!ProcessingResult
        details(7) = rs!IsProcessed
        details(8) = rs!PrintAttempts
        details(9) = rs!LastPrintAttempt
        details(10) = rs!CreatedDate
        db_GetRecordDetails = details
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("db_GetRecordDetails", Err.Number, Err.Description, CStr(recordID))
    db_GetRecordDetails = Null
    Resume CleanExit
End Function

' =========================================================
' تابع: db_GetSuccessfulRecordsCount
' =========================================================
'
' وظیفه:
' تعداد تردد‌های موفق امروز را برمی‌گرداند
'
' خروجی:
'   Long: تعداد تردد‌های موفق
'
' =========================================================

Public Function db_GetSuccessfulRecordsCount() As Long
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim count As Long
    Dim todayStart As Date
    
    Set db = CurrentDb()
    
    todayStart = DateValue(Now())
    
    sql = "SELECT COUNT(*) as cnt FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE ProcessingResult='" & RESULT_PRINTED & "' " & _
          "AND AttendanceDateTime >= " & SqlDateTime(todayStart)
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        count = Nz(rs!cnt, 0)
    End If
    
    db_GetSuccessfulRecordsCount = count
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    db_GetSuccessfulRecordsCount = 0
End Function

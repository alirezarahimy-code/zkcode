Option Compare Database
Option Explicit

' =========================================================
' ماژول: receipt_service.bas
' =========================================================
' 
' توضیح ماژول:
' مدیریت فیش‌های چاپ‌شده
' ایجاد، بروزرسانی، پرس‌و‌جو فیش‌ها
'
' =========================================================

' =========================================================
' تابع: receipt_CreatePending
' =========================================================
' 
' وظیفه:
' فیش درانتظار برای تردد و غذا ایجاد می‌کند
'
' =========================================================

Public Function receipt_CreatePending(ByVal empID As String, ByVal mealListID As Long, _
                                       ByVal deviceIP As String, ByVal attendanceRecordID As Long) As Long
    On Error GoTo ErrHandler
    
    Dim existing As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim transactionStarted As Boolean
    
    empID = Trim$(empID)
    
    If Len(empID) = 0 Or mealListID <= 0 Or attendanceRecordID <= 0 Then Exit Function
    
    ' بررسی فیش موجود
    existing = receipt_GetByAttendance(attendanceRecordID)
    If existing > 0 Then
        receipt_CreatePending = existing
        Exit Function
    End If
    
    Set db = CurrentDb()
    
    DBEngine(0).BeginTrans
    transactionStarted = True
    
    Set rs = db.OpenRecordset(TABLE_RECEIPTS, dbOpenDynaset)
    
    rs.AddNew
    rs!EmployeeID = empID
    rs!MealListID = mealListID
    rs!AttendanceRecordID = attendanceRecordID
    rs!PendingDateTime = Now()
    rs!PrinterName = GetDefaultPrinterName()
    rs!PrinterSharePath = GetDefaultPrinterSharePath()
    rs!PrinterIP = GetDefaultPrinterIP()
    rs!PrinterPort = GetDefaultPrinterPort()
    rs!PrinterType = GetDefaultPrinterType()
    rs!IsSuccessful = False
    rs!PrintStatus = PRINT_STATUS_PENDING
    rs!ErrorMessage = Null
    rs!IPAddress = deviceIP
    rs!CreatedDate = Now()
    
    rs.Update
    rs.Bookmark = rs.LastModified
    receipt_CreatePending = Nz(rs!ReceiptID, 0)
    
    rs.Close
    Set rs = Nothing
    
    DBEngine(0).CommitTrans
    transactionStarted = False
    
    Set db = Nothing
    Exit Function
    
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If transactionStarted Then DBEngine(0).Rollback
    Set db = Nothing
    
    existing = receipt_GetByAttendance(attendanceRecordID)
    If existing > 0 Then
        receipt_CreatePending = existing
        Exit Function
    End If
    
    Call LogError("receipt_CreatePending", en, ed, _
                  "EmployeeID=" & empID & " | MealListID=" & CStr(mealListID))
    receipt_CreatePending = 0
End Function

' =========================================================
' تابع: receipt_GetByAttendance
' =========================================================
' 
' وظیفه:
' فیش برای یک تردد را برمی‌گرداند
'
' =========================================================

Public Function receipt_GetByAttendance(ByVal attendanceRecordID As Long) As Long
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    
    If attendanceRecordID <= 0 Then Exit Function
    
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT TOP 1 ReceiptID FROM " & TABLE_RECEIPTS & _
        " WHERE AttendanceRecordID=" & CStr(attendanceRecordID), _
        dbOpenSnapshot)
    
    If Not rs.EOF Then
        receipt_GetByAttendance = Nz(rs!ReceiptID, 0)
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("receipt_GetByAttendance", Err.Number, Err.Description, _
                  CStr(attendanceRecordID))
    receipt_GetByAttendance = 0
    Resume CleanExit
End Function

' =========================================================
' تابع: receipt_GetPrintStatus
' =========================================================
' 
' وظیفه:
' وضعیت چاپ فیش را برمی‌گرداند
'
' =========================================================

Public Function receipt_GetPrintStatus(ByVal receiptID As Long) As String
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    
    If receiptID <= 0 Then Exit Function
    
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT PrintStatus FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), _
        dbOpenSnapshot)
    
    If Not rs.EOF Then
        receipt_GetPrintStatus = UCase$(Nz(rs!PrintStatus, ""))
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    receipt_GetPrintStatus = ""
    Resume CleanExit
End Function

' =========================================================
' تابع: receipt_GetMealListID
' =========================================================
' 
' وظیفه:
' شماره غذا برای فیش را برمی‌گرداند
'
' =========================================================

Public Function receipt_GetMealListID(ByVal receiptID As Long) As Long
    On Error GoTo ErrHandler
    
    If receiptID <= 0 Then Exit Function
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT MealListID FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), _
        dbOpenSnapshot)
    
    If Not rs.EOF Then receipt_GetMealListID = Nz(rs!MealListID, 0)
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    receipt_GetMealListID = 0
    Resume CleanExit
End Function

' =========================================================
' تابع: receipt_AssignMeal
' =========================================================
' 
' وظیفه:
' غذا را به فیش اختصاص می‌دهد
'
' =========================================================

Public Function receipt_AssignMeal(ByVal receiptID As Long, ByVal mealListID As Long) As Boolean
    On Error GoTo ErrHandler
    
    If receiptID <= 0 Or mealListID <= 0 Then Exit Function
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT MealListID, PrintStatus FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), _
        dbOpenDynaset)
    
    If rs.EOF Then GoTo CleanExit
    
    ' اگر قبلاً چاپ شده، نتوانیم تغییر دهیم
    If UCase$(Nz(rs!PrintStatus, "")) = PRINT_STATUS_SUCCESS Then
        receipt_AssignMeal = (Nz(rs!MealListID, 0) = mealListID)
        GoTo CleanExit
    End If
    
    rs.Edit
    rs!MealListID = mealListID
    rs.Update
    receipt_AssignMeal = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    receipt_AssignMeal = False
    Resume CleanExit
End Function

' =========================================================
' تابع: receipt_SetPrintStatus
' =========================================================
' 
' وظیفه:
' وضعیت چاپ فیش را بروزرسانی می‌کند
'
' =========================================================

Public Function receipt_SetPrintStatus(ByVal receiptID As Long, ByVal status As String, _
                                        Optional ByVal errorText As String = "") As Boolean
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim normalized As String
    
    If receiptID <= 0 Then Exit Function
    
    normalized = UCase$(Trim$(status))
    
    ' اعتبار‌سنجی وضعیت
    Select Case normalized
        Case PRINT_STATUS_PENDING, PRINT_STATUS_SUBMITTED, _
             PRINT_STATUS_SUCCESS, PRINT_STATUS_FAILED, PRINT_STATUS_UNKNOWN
        Case Else
            Call LogError("receipt_SetPrintStatus", -1, "Invalid status", normalized)
            Exit Function
    End Select
    
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT IsSuccessful, PrintStatus, ErrorMessage, PrintDateTime " & _
        "FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), _
        dbOpenDynaset)
    
    If rs.EOF Then GoTo CleanExit
    
    rs.Edit
    rs!PrintStatus = normalized
    rs!IsSuccessful = (normalized = PRINT_STATUS_SUCCESS)
    
    If normalized = PRINT_STATUS_SUCCESS Then
        rs!PrintDateTime = Now()
        rs!ErrorMessage = Null
    ElseIf Len(errorText) > 0 Then
        rs!ErrorMessage = errorText
    ElseIf normalized = PRINT_STATUS_SUBMITTED Then
        rs!ErrorMessage = "ارسال شد؛ نتیجه معلوم نیست"
    ElseIf normalized = PRINT_STATUS_UNKNOWN Then
        rs!ErrorMessage = "نتیجه نامعلوم"
    ElseIf normalized = PRINT_STATUS_FAILED Then
        rs!PrintDateTime = Null
        rs!ErrorMessage = "چاپ ناموفق"
    End If
    
    rs.Update
    receipt_SetPrintStatus = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("receipt_SetPrintStatus", Err.Number, Err.Description, CStr(receiptID))
    receipt_SetPrintStatus = False
    Resume CleanExit
End Function
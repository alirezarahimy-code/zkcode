Option Compare Database
Option Explicit

' =========================================================
' ماژول: meal_service.bas
' =========================================================
' 
' توضیح ماژول:
' مدیریت سفارشات غذا
' بررسی سفارش برای کارمند
' Claim غذا برای فیش
'
' =========================================================

' =========================================================
' تابع: meal_FindForEmployee
' =========================================================
' 
' وظیفه:
' غذای سفارش‌شده برای کارمند در تاریخ مشخص را پیدا می‌کند
'
' پارامترها:
'   empID (String): کد پرسنلی
'   attDT (Date): تاریخ تردد
'
' خروجی: Variant (Array)
'   Array(0) = MealListID
'   Array(1) = MealType
'   یا Null اگر یافت نشد
'
' =========================================================

Public Function meal_FindForEmployee(ByVal empID As String, ByVal attDT As Date) As Variant
    On Error GoTo ErrHandler
    
    empID = Trim$(empID)
    If Len(empID) = 0 Then
        meal_FindForEmployee = Null
        Exit Function
    End If
    
    Dim rs As DAO.Recordset
    Dim startDT As Date, endDT As Date
    Dim a(0 To 1) As Variant
    
    startDT = DateSerial(Year(attDT), Month(attDT), Day(attDT))
    endDT = DateAdd("d", 1, startDT)
    
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT TOP 1 MealListID, MealType " & _
        "FROM " & TABLE_DAILY_MEALS & " " & _
        "WHERE EmployeeID='" & Replace(empID, "'", "''") & "' " & _
        "AND MealDate >= " & SqlDateTime(startDT) & " " & _
        "AND MealDate < " & SqlDateTime(endDT) & " " & _
        "AND Nz(MealDelivered,False)=False " & _
        "AND Nz(DeliveryReceiptID,0)=0 " & _
        "ORDER BY MealListID", _
        dbOpenSnapshot)
    
    If rs.EOF Then
        meal_FindForEmployee = Null
    Else
        a(0) = Nz(rs!MealListID, 0)
        a(1) = Nz(rs!MealType, "-")
        meal_FindForEmployee = a
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("meal_FindForEmployee", Err.Number, Err.Description, "EmployeeID=" & empID)
    meal_FindForEmployee = Null
    Resume CleanExit
End Function

' =========================================================
' تابع: meal_GetTypeByID
' =========================================================
' 
' وظیفه:
' نوع غذا را از شماره غذا برمی‌گرداند
'
' =========================================================

Public Function meal_GetTypeByID(ByVal mealListID As Long) As String
    On Error GoTo ErrHandler
    
    If mealListID <= 0 Then Exit Function
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT MealType FROM " & TABLE_DAILY_MEALS & " WHERE MealListID=" & CStr(mealListID), _
        dbOpenSnapshot)
    
    If Not rs.EOF Then meal_GetTypeByID = Nz(rs!MealType, "")
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandler:
    meal_GetTypeByID = ""
    Resume CleanExit
End Function

' =========================================================
' تابع: meal_ClaimForReceipt
' =========================================================
' 
' وظیفه:
' Atomic Claim غذا برای فیش
' فقط یک فیش می‌تواند این غذا را بگیرد
'
' =========================================================

Public Function meal_ClaimForReceipt(ByVal mealListID As Long, ByVal receiptID As Long, _
                                      ByVal attendanceRecordID As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim sql As String
    
    If mealListID <= 0 Or receiptID <= 0 Or attendanceRecordID <= 0 Then Exit Function
    
    Set db = CurrentDb()
    
    ' اعتبارسنجی رابطه Receipt و Attendance
    Dim rsCheck As DAO.Recordset
    Set rsCheck = db.OpenRecordset( _
        "SELECT MealListID,AttendanceRecordID FROM " & TABLE_RECEIPTS & _
        " WHERE ReceiptID=" & CStr(receiptID), dbOpenSnapshot)
    
    If rsCheck.EOF Then GoTo CleanExit
    If Nz(rsCheck!MealListID, 0) <> mealListID Then GoTo CleanExit
    If Nz(rsCheck!AttendanceRecordID, 0) <> attendanceRecordID Then GoTo CleanExit
    rsCheck.Close
    Set rsCheck = Nothing
    
    ' UPDATE Atomic
    sql = "UPDATE " & TABLE_DAILY_MEALS & " " & _
          "SET DeliveryReceiptID=" & CStr(receiptID) & " " & _
          "WHERE MealListID=" & CStr(mealListID) & " " & _
          "AND Nz(MealDelivered,False)=False " & _
          "AND Nz(DeliveryReceiptID,0)=0"
    
    db.Execute sql, dbFailOnError
    
    If db.RecordsAffected > 0 Then
        meal_ClaimForReceipt = True
        Set db = Nothing
        Exit Function
    End If
    
    ' اگر قبلاً claim شده
    sql = "SELECT DeliveryReceiptID,MealDelivered FROM " & TABLE_DAILY_MEALS & _
          " WHERE MealListID=" & CStr(mealListID)
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        If Nz(rs!DeliveryReceiptID, 0) = receiptID And Nz(rs!MealDelivered, False) = False Then
            meal_ClaimForReceipt = True
        End If
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not rsCheck Is Nothing Then rsCheck.Close
    Set rs = Nothing
    Set rsCheck = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("meal_ClaimForReceipt", Err.Number, Err.Description, _
                  "MealListID=" & CStr(mealListID) & " | ReceiptID=" & CStr(receiptID))
    meal_ClaimForReceipt = False
    Resume CleanExit
End Function

' =========================================================
' تابع: meal_MarkDelivered
' =========================================================
' 
' وظیفه:
' غذا را به عنوان تحویل داده شده علامت می‌زند
' در Transaction برای اطمینان از یکپارچگی
'
' =========================================================

Public Function meal_MarkDelivered(ByVal mealListID As Long, ByVal receiptID As Long, _
                                    ByVal attendanceRecordID As Long) As Boolean
    On Error GoTo ErrHandler
    
    If mealListID <= 0 Or receiptID <= 0 Or attendanceRecordID <= 0 Then Exit Function
    
    Dim db As DAO.Database
    Dim rsMeal As DAO.Recordset
    Dim rsReceipt As DAO.Recordset
    Dim transactionStarted As Boolean
    Dim nowDT As Date
    
    Set db = CurrentDb()
    nowDT = Now()
    
    DBEngine(0).BeginTrans
    transactionStarted = True
    
    ' بررسی غذا
    Set rsMeal = db.OpenRecordset( _
        "SELECT MealListID,EmployeeID,MealDelivered,DeliveryReceiptID " & _
        "FROM " & TABLE_DAILY_MEALS & " WHERE MealListID=" & CStr(mealListID), _
        dbOpenDynaset)
    
    If rsMeal.EOF Then GoTo TransactionFailed
    
    ' بررسی فیش
    Set rsReceipt = db.OpenRecordset( _
        "SELECT ReceiptID,EmployeeID,MealListID,AttendanceRecordID,IsSuccessful,PrintStatus " & _
        "FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), _
        dbOpenDynaset)
    
    If rsReceipt.EOF Then GoTo TransactionFailed
    
    ' اعتبارسنجی‌ها
    If Nz(rsReceipt!MealListID, 0) <> mealListID Then GoTo TransactionFailed
    If Nz(rsReceipt!AttendanceRecordID, 0) <> attendanceRecordID Then GoTo TransactionFailed
    If UCase$(Nz(rsReceipt!PrintStatus, "")) <> PRINT_STATUS_SUCCESS Then GoTo TransactionFailed
    If Nz(rsReceipt!IsSuccessful, False) <> True Then GoTo TransactionFailed
    
    ' اگر قبلاً تحویل داده شده
    If Nz(rsMeal!MealDelivered, False) = True Then
        DBEngine(0).CommitTrans
        transactionStarted = False
        meal_MarkDelivered = True
        GoTo CleanExit
    End If
    
    ' بروزرسانی غذا
    rsMeal.Edit
    rsMeal!DeliveryReceiptID = receiptID
    rsMeal!MealDelivered = True
    rsMeal!DeliveredDateTime = nowDT
    rsMeal.Update
    
    ' بروزرسانی فیش
    rsReceipt.Edit
    rsReceipt!IsSuccessful = True
    rsReceipt!PrintStatus = PRINT_STATUS_SUCCESS
    rsReceipt!PrintDateTime = nowDT
    rsReceipt!ErrorMessage = Null
    rsReceipt.Update
    
    DBEngine(0).CommitTrans
    transactionStarted = False
    meal_MarkDelivered = True
    
CleanExit:
    On Error Resume Next
    If Not rsMeal Is Nothing Then rsMeal.Close
    If Not rsReceipt Is Nothing Then rsReceipt.Close
    Set rsMeal = Nothing
    Set rsReceipt = Nothing
    Set db = Nothing
    Exit Function
    
TransactionFailed:
    On Error Resume Next
    If transactionStarted Then DBEngine(0).Rollback
    transactionStarted = False
    If Not rsMeal Is Nothing Then rsMeal.Close
    If Not rsReceipt Is Nothing Then rsReceipt.Close
    Set rsMeal = Nothing
    Set rsReceipt = Nothing
    Set db = Nothing
    meal_MarkDelivered = False
    Exit Function
    
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    
    On Error Resume Next
    If transactionStarted Then DBEngine(0).Rollback
    transactionStarted = False
    If Not rsMeal Is Nothing Then rsMeal.Close
    If Not rsReceipt Is Nothing Then rsReceipt.Close
    Set rsMeal = Nothing
    Set rsReceipt = Nothing
    Set db = Nothing
    
    Call LogError("meal_MarkDelivered", en, ed, _
                  "MealListID=" & CStr(mealListID) & " | ReceiptID=" & CStr(receiptID))
    meal_MarkDelivered = False
End Function

' =========================================================
' تابع: meal_GetAll
' =========================================================
' 
' وظیفه:
' تمام غذاهای امروز را برمی‌گرداند
'
' =========================================================

Public Function meal_GetAll() As DAO.Recordset
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    Set meal_GetAll = db.OpenRecordset( _
        "SELECT MealListID, EmployeeID, FullName, MealType, MealDelivered " & _
        "FROM " & TABLE_DAILY_MEALS & " " & _
        "WHERE Int(MealDate)=Int(Now()) " & _
        "ORDER BY FullName", dbOpenSnapshot)
    
    Exit Function
    
ErrHandler:
    Call LogError("meal_GetAll", Err.Number, Err.Description, "")
    Set meal_GetAll = Nothing
End Function
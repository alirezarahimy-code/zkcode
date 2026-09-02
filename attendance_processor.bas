Option Compare Database
Option Explicit

' =========================================================
' ماژول: attendance_processor.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول تردد‌های ثبت‌شده را پردازش می‌کند.
' فرآیند پردازش:
'   1. بررسی معتبر بودن کارمند
'   2. بررسی سفارش غذا برای امروز
'   3. ایجاد یا بازیابی فیش
'   4. Atomic Claim غذا برای فیش
'   5. چاپ فیش
'   6. علامت‌گذاری غذا به عنوان تحویل داده شده
'
' کاربرد:
' - پردازش تردد‌های جدید (Real-Time)
' - پردازش تردد‌های باقی‌مانده (Batch)
' - بازیابی تردد‌های ناموفق
'
' ویژگی‌های مهم:
' - Atomic Transaction برای جلوگیری از تاثیرات دوگانه
' - Idempotent (چند بار اجرا شود، نتیجه یکی است)
' - معالجه تمام خطاهای احتمالی
' - لاگ کامل برای دیباگ
'
' معماری:
' - proc_ProcessPendingBatch: پردازش دسته‌ای
' - proc_ProcessRecordAtomic: پردازش یک تردد
' - proc_RecoverStuckPrinting: بازیابی چاپ‌های گیر‌کرده
'
' =========================================================

' =========================================================
' تابع: proc_ProcessPendingBatch
' =========================================================
' 
' وظیفه:
' تردد‌های درانتظار و ناموفق را یکی یکی پردازش می‌کند
' معمولاً هر 5 ثانیه یک بار (در monitor_Tick) فراخوانی می‌شود
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' فرآیند:
'   1. انتخاب تا 100 رکورد درانتظار
'   2. بر اساس تعداد تلاش (کم‌تر اول)
'   3. برای هر رکورد: تلاش Atomic Claim
'   4. اگر موفق: فراخوانی proc_ProcessRecordAtomic
'
' نمونه استفاده:
'   Call proc_ProcessPendingBatch()
'
' =========================================================

Public Sub proc_ProcessPendingBatch()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim recID As Long
    
    Set db = CurrentDb()
    
    ' انتخاب تردد‌های درانتظار (تا 100 رکورد)
    Dim sql As String
    sql = "SELECT TOP " & CStr(PROCESS_BATCH_SIZE) & " RecordID " & _
          "FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE Nz(ProcessingResult,'NEW') IN " & _
          "('NEW','PRINT_FAILED','NO_EMPLOYEE','WAITING_FOR_MEAL','RECEIPT_FAILED','MEAL_FINALIZE_FAILED') " & _
          "AND Nz(PrintAttempts,0) < " & CStr(MAX_PRINT_ATTEMPTS) & " " & _
          "ORDER BY Nz(PrintAttempts,0), RecordID"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    Do While Not rs.EOF
        recID = Nz(rs!RecordID, 0)
        
        If recID > 0 Then
            ' تلاش برای Atomic Claim
            If db_TryClaimRecordAtomic(recID, "PROCESSING", _
                                      "'NEW','PRINT_FAILED','NO_EMPLOYEE','WAITING_FOR_MEAL','RECEIPT_FAILED','MEAL_FINALIZE_FAILED'") Then
                Call proc_ProcessRecordAtomic(recID)
            End If
        End If
        
        rs.MoveNext
    Loop
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Call LogError("proc_ProcessPendingBatch", en, ed, "")
End Sub

' =========================================================
' تابع: proc_ProcessRecordAtomic
' =========================================================
' 
' وظیفه:
' یک تردد واحد را به صورت Atomic پردازش می‌کند
' هر مرحله دارای معالجه خطا است
'
' پارامتر:
'   recID (Long): شماره رکورد تردد
'
' فرآیند:
'   1. دریافت اطلاعات تردد
'   2. بررسی: کارمند موجود است؟
'   3. بررسی: امروز غذا سفارش داده؟
'   4. بررسی: قبلاً فیش چاپ شده؟
'   5. اگر فیش قدیمی موجود: استفاده از آن
'   6. اگر نه: ایجاد فیش جدید
'   7. Atomic Claim غذا برای فیش
'   8. چاپ فیش
'   9. علامت‌گذاری غذا به عنوان تحویل داده شده
'  10. بروزرسانی وضعیت رکورد
'
' خطاهای احتمالی:
'   - NO_EMPLOYEE: کارمند یافت نشد
'   - WAITING_FOR_MEAL: غذا سفارش نشده
'   - ALREADY_PRINTED: فیش قبلاً چاپ شده
'   - RECEIPT_FAILED: خطا در ایجاد فیش
'   - MEAL_FINALIZE_FAILED: خطا در تکمیل غذا
'   - PRINT_FAILED: چاپ ناموفق
'   - PRINT_UNKNOWN: نتیجه چاپ نامعلوم
'
' =========================================================

Public Sub proc_ProcessRecordAtomic(ByVal recID As Long)
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim deviceEnroll As String, deviceIP As String
    Dim attDT As Date
    Dim devicePort As Long, machineNumber As Long
    Dim printAttempts As Long
    Dim emp As Variant, meal As Variant
    Dim empID As String, fullName As String
    Dim mealListID As Long, mealType As String
    Dim receiptID As Long, receiptMealID As Long
    Dim receiptStatus As String
    Dim printerCalled As Boolean
    Dim printedOK As Boolean
    Dim mealOK As Boolean
    Dim newAttempts As Long
    
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT DeviceEnrollID,AttendanceDateTime,DeviceIP,DevicePort,DeviceMachineNumber,PrintAttempts,ProcessingResult " & _
        "FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recID), dbOpenSnapshot)
    
    If rs.EOF Then GoTo CleanExit
    
    ' بررسی: رکورد در حالت PROCESSING است؟
    If UCase$(Nz(rs!ProcessingResult, "")) <> "PROCESSING" Then GoTo CleanExit
    
    deviceEnroll = Nz(rs!DeviceEnrollID, "")
    attDT = Nz(rs!AttendanceDateTime, Now())
    deviceIP = Nz(rs!DeviceIP, "")
    devicePort = Nz(rs!DevicePort, DEFAULT_ZK_PORT)
    machineNumber = Nz(rs!DeviceMachineNumber, 1)
    printAttempts = Nz(rs!PrintAttempts, 0)
    
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If machineNumber <= 0 Then machineNumber = 1
    
    rs.Close
    Set rs = Nothing
    
    ' مرحله 1: بررسی کارمند
    emp = emp_ResolveByDeviceEnroll(deviceEnroll)
    If IsNull(emp) Then
        Call UpdateProcessingState(recID, RESULT_NO_EMPLOYEE, False, printAttempts, False)
        Call LogAttendanceEvent(deviceEnroll, RESULT_NO_EMPLOYEE, "کارمند یافت نشد")
        Exit Sub
    End If
    
    empID = Nz(emp(0), "")
    fullName = Nz(emp(1), "")
    If Len(Trim$(empID)) = 0 Then
        Call UpdateProcessingState(recID, RESULT_NO_EMPLOYEE, False, printAttempts, False)
        Call LogAttendanceEvent(deviceEnroll, RESULT_NO_EMPLOYEE, "کارمند یافت نشد")
        Exit Sub
    End If
    
    ' مرحله 2: بررسی فیش قدیمی
    receiptID = receipt_GetByAttendance(recID)
    If receiptID > 0 Then
        receiptStatus = receipt_GetPrintStatus(receiptID)
        receiptMealID = receipt_GetMealListID(receiptID)
        
        ' اگر فیش قبلاً با موفقیت چاپ شد
        If receiptStatus = PRINT_STATUS_SUCCESS Then
            If receiptMealID > 0 Then
                printedOK = True
                printerCalled = False
                mealOK = meal_MarkDelivered(receiptMealID, receiptID, recID)
                If mealOK Then
                    Call UpdateProcessingState(recID, RESULT_PRINTED, True, printAttempts, False)
                    Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش چاپ شد")
                Else
                    Call UpdateProcessingState(recID, RESULT_MEAL_FINALIZE_FAILED, False, printAttempts, False)
                    Call LogAttendanceEvent(deviceEnroll, RESULT_MEAL_FINALIZE_FAILED, "خطا در تکمیل غذا")
                End If
            Else
                Call UpdateProcessingState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, False)
                Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "فیش بدون غذا")
            End If
            Exit Sub
        End If
        
        ' اگر چاپ نتیجه نامعلوم است
        If receiptStatus = PRINT_STATUS_SUBMITTED Or receiptStatus = PRINT_STATUS_UNKNOWN Then
            Call UpdateProcessingState(recID, RESULT_PRINT_UNKNOWN, False, printAttempts, False)
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "نتیجه چاپ نامعلوم")
            Exit Sub
        End If
        
        ' اگر فیش خالی است
        If receiptStatus = "" Then
            Call UpdateProcessingState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, False)
            Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "فیش خالی")
            Exit Sub
        End If
    End If
    
    ' مرحله 3: بررسی سفارش غذا
    receiptMealID = 0
    If receiptID > 0 Then
        receiptMealID = receipt_GetMealListID(receiptID)
    End If
    
    If receiptMealID > 0 Then
        mealListID = receiptMealID
    Else
        meal = meal_FindForEmployee(empID, attDT)
        If IsNull(meal) Then
            Call UpdateProcessingState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, False)
            Call LogAttendanceEvent(deviceEnroll, RESULT_NO_MEAL_ORDER, "غذا سفارش نشده")
            Exit Sub
        End If
        
        mealListID = Nz(meal(0), 0)
        mealType = Nz(meal(1), "-")
        
        If mealListID <= 0 Then
            Call UpdateProcessingState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, False)
            Call LogAttendanceEvent(deviceEnroll, RESULT_NO_MEAL_ORDER, "غذا سفارش نشده")
            Exit Sub
        End If
    End If
    
    ' مرحله 4: ایجاد یا بازیابی فیش
    If receiptID <= 0 Then
        receiptID = receipt_CreatePending(empID, mealListID, deviceIP, recID)
        receiptMealID = receipt_GetMealListID(receiptID)
    End If
    
    If receiptID > 0 And receiptMealID <= 0 Then
        If Not receipt_AssignMeal(receiptID, mealListID) Then
            Call UpdateProcessingState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, False)
            Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "اختصاص غذا ناموفق")
            Exit Sub
        End If
        receiptMealID = mealListID
    End If
    
    ' دریافت نوع غذا برای چاپ
    If receiptMealID > 0 Then
        mealListID = receiptMealID
        mealType = meal_GetTypeByID(mealListID)
        If Len(Trim$(mealType)) = 0 Then mealType = "-"
    End If
    
    If receiptID <= 0 Then
        Call UpdateProcessingState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, False)
        Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "فیش ایجاد نشد")
        Exit Sub
    End If
    
    ' مرحله 5: Atomic Claim غذا
    If Not meal_ClaimForReceipt(mealListID, receiptID, recID) Then
        Call receipt_SetPrintStatus(receiptID, PRINT_STATUS_FAILED, "غذا قبلاً گرفته شده")
        Call UpdateProcessingState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, False)
        Call LogAttendanceEvent(deviceEnroll, RESULT_NO_MEAL_ORDER, "غذا قبلاً گرفته شده")
        Exit Sub
    End If
    
    receiptStatus = receipt_GetPrintStatus(receiptID)
    If receiptStatus = "" Then
        Call UpdateProcessingState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, False)
        Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "وضعیت فیش خالی")
        Exit Sub
    End If
    
    ' اگر فیش قبلاً چاپ شده
    If receiptStatus = PRINT_STATUS_SUCCESS Then
        mealOK = meal_MarkDelivered(mealListID, receiptID, recID)
        If mealOK Then
            Call UpdateProcessingState(recID, RESULT_PRINTED, True, printAttempts, False)
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش قبلاً چاپ شد")
        Else
            Call UpdateProcessingState(recID, RESULT_MEAL_FINALIZE_FAILED, False, printAttempts, False)
            Call LogAttendanceEvent(deviceEnroll, RESULT_MEAL_FINALIZE_FAILED, "خطا در تکمیل غذا")
        End If
        Exit Sub
    End If
    
    ' مرحله 6: چاپ فیش
    If Not SetPrintAttemptStarted(recID) Then
        Call UpdateProcessingState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, False)
        Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "خطا در ثبت تلاش چاپ")
        Exit Sub
    End If
    
    printerCalled = True
    
    printedOK = printer_PrintReceipt(empID, fullName, mealType, mealListID, receiptID, deviceIP)
    
    If printedOK Then
        mealOK = meal_MarkDelivered(mealListID, receiptID, recID)
        If mealOK Then
            newAttempts = MinAttempts(printAttempts + 1)
            Call UpdateProcessingState(recID, RESULT_PRINTED, True, newAttempts, True)
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش چاپ شد ✓")
        Else
            newAttempts = MinAttempts(printAttempts + 1)
            Call UpdateProcessingState(recID, RESULT_PRINT_UNKNOWN, False, newAttempts, True)
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "خطا در تکمیل غذا")
        End If
    Else
        receiptStatus = receipt_GetPrintStatus(receiptID)
        newAttempts = MinAttempts(printAttempts + 1)
        
        If receiptStatus = PRINT_STATUS_SUBMITTED Or receiptStatus = PRINT_STATUS_UNKNOWN Then
            Call UpdateProcessingState(recID, RESULT_PRINT_UNKNOWN, False, newAttempts, True)
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "نتیجه چاپ نامعلوم")
        Else
            Call UpdateProcessingState(recID, RESULT_PRINT_FAILED, False, newAttempts, True)
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_FAILED, "چاپ ناموفق")
        End If
    End If
    
    Exit Sub
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    Call LogError("proc_ProcessRecordAtomic", en, ed, CStr(recID))
    
    If printerCalled Then
        Call UpdateProcessingState(recID, RESULT_PRINT_UNKNOWN, False, _
                                  MinAttempts(printAttempts + 1), True)
    Else
        Call UpdateProcessingState(recID, RESULT_MEAL_FINALIZE_FAILED, False, printAttempts, False)
    End If
End Sub

' =========================================================
' تابع کمکی: SetPrintAttemptStarted
' =========================================================

Private Function SetPrintAttemptStarted(ByVal recID As Long) As Boolean
    On Error GoTo EH
    
    Dim rs As DAO.Recordset
    
    If recID <= 0 Then Exit Function
    
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT LastPrintAttempt FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recID), _
        dbOpenDynaset)
    
    If rs.EOF Then GoTo CleanExit
    
    rs.Edit
    rs!LastPrintAttempt = Now()
    rs!PrintAttempts = MinAttempts(Nz(rs!PrintAttempts, 0) + 1)
    rs.Update
    
    SetPrintAttemptStarted = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
    
EH:
    SetPrintAttemptStarted = False
    Resume CleanExit
End Function

' =========================================================
' تابع کمکی: MinAttempts
' =========================================================

Private Function MinAttempts(ByVal n As Long) As Long
    If n < 0 Then n = 0
    If n > MAX_PRINT_ATTEMPTS Then n = MAX_PRINT_ATTEMPTS
    MinAttempts = n
End Function

' =========================================================
' تابع کمکی: UpdateProcessingState
' =========================================================

Private Sub UpdateProcessingState(ByVal recID As Long, ByVal newState As String, _
                                  Optional ByVal markProcessed As Boolean = False, _
                                  Optional ByVal newPrintAttempts As Long = -1, _
                                  Optional ByVal updatePrintAttemptDate As Boolean = False)
    On Error GoTo ErrHandler
    
    If recID <= 0 Then Exit Sub
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb().OpenRecordset( _
        "SELECT ProcessingResult, IsProcessed, PrintAttempts, LastPrintAttempt " & _
        "FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recID), _
        dbOpenDynaset)
    
    If rs.EOF Then GoTo CleanExit
    
    rs.Edit
    rs!ProcessingResult = newState
    rs!IsProcessed = markProcessed
    
    If newPrintAttempts >= 0 Then
        rs!PrintAttempts = newPrintAttempts
    End If
    
    If updatePrintAttemptDate Then
        rs!LastPrintAttempt = Now()
    End If
    
    rs.Update
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    Call LogError("UpdateProcessingState", Err.Number, Err.Description, CStr(recID))
    Resume CleanExit
End Sub

' =========================================================
' تابع: proc_RecoverStuckPrinting
' =========================================================
' 
' وظیفه:
' چاپ‌های گیر‌کرده را بازیابی می‌کند
' اگر یک چاپ بیش از 10 دقیقه "PROCESSING" بود، بازگرد آن
'
' =========================================================

Public Sub proc_RecoverStuckPrinting()
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT RecordID, LastPrintAttempt FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE ProcessingResult='PROCESSING' " & _
          "AND Nz(LastPrintAttempt,CreatedDate)<" & _
          SqlDateTime(DateAdd("n", -PRINTING_STUCK_TIMEOUT_MINUTES, Now()))
    
    Set rs = CurrentDb().OpenRecordset(sql, dbOpenDynaset)
    
    Do While Not rs.EOF
        rs.Edit
        
        ' اگر هیچ تلاش چاپی نبود، ناموفق
        If IsNull(rs!LastPrintAttempt) Then
            rs!ProcessingResult = RESULT_PRINT_FAILED
        Else
            ' اگر تلاش بود، نامعلوم
            rs!ProcessingResult = RESULT_PRINT_UNKNOWN
        End If
        
        rs!IsProcessed = False
        rs.Update
        
        Call LogSystemEvent("proc_RecoverStuckPrinting", _
                           "چاپ گیر‌کرده بازیابی شد: RecordID=" & CStr(rs!RecordID))
        
        rs.MoveNext
    Loop
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    Call LogError("proc_RecoverStuckPrinting", Err.Number, Err.Description, "")
    Resume CleanExit
End Sub
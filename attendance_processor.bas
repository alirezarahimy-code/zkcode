Option Compare Database
Option Explicit

' =========================================================
' ماژول: attendance_processor.bas (cleaned)
' =========================================================
' این نسخه از proc_ProcessRecordAtomic از سرویس‌های متمرکز db_service و receipt_service استفاده می‌کند' تا جلوگیری از پراکندگی منطق بر روی ماژول‌ها شود.
' =========================================================

Public Sub proc_ProcessRecordAtomic(ByVal recID As Long)
    On Error GoTo ErrHandler
    If recID <= 0 Then Exit Sub
    Dim db As DAO.Database
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
    Dim printedOK As Boolean
    Dim mealOK As Boolean
    Set db = CurrentDb()        ' خواندن رکورد پایه
    Set rs = db.OpenRecordset("SELECT DeviceEnrollID,AttendanceDateTime,DeviceIP,DevicePort,DeviceMachineNumber,PrintAttempts,ProcessingResult FROM " & TABLE_ATTENDANCE & " WHERE RecordID=" & CStr(recID), dbOpenSnapshot)
    If rs.EOF Then GoTo CleanExit
    deviceEnroll = Nz(rs!DeviceEnrollID, "")
    attDT = Nz(rs!AttendanceDateTime, Now())
    deviceIP = Nz(rs!DeviceIP, "")
    devicePort = Nz(rs!DevicePort, DEFAULT_ZK_PORT)
    machineNumber = Nz(rs!DeviceMachineNumber, 1)
    printAttempts = Nz(rs!PrintAttempts, 0)
    rs.Close: Set rs = Nothing
    ' 1) Resolve employee
    emp = emp_ResolveByDeviceEnroll(deviceEnroll)
    If IsNull(emp) Then
        Call db_UpdateRecordState(recID, RESULT_NO_EMPLOYEE, False, printAttempts, Null, "کارمند یافت نشد")
        Call LogAttendanceEvent(deviceEnroll, RESULT_NO_EMPLOYEE, "کارمند یافت نشد")
        Exit Sub
    End If
    empID = Nz(emp(0), "")
    fullName = Nz(emp(1), "")
    If Len(Trim$(empID)) = 0 Then
        Call db_UpdateRecordState(recID, RESULT_NO_EMPLOYEE, False, printAttempts, Null, "کارمند نامعتبر")
        Call LogAttendanceEvent(deviceEnroll, RESULT_NO_EMPLOYEE, "کارمند نامعتبر")
        Exit Sub
    End If
    ' 2) Check for existing receipt
    receiptID = receipt_GetByAttendance(recID)
    If receiptID > 0 Then
        receiptStatus = receipt_GetPrintStatus(receiptID)
        receiptMealID = receipt_GetMealListID(receiptID)
        If receiptStatus = PRINT_STATUS_SUCCESS Then
            If receiptMealID > 0 Then
                mealOK = meal_MarkDelivered(receiptMealID, receiptID, recID)
                If mealOK Then
                    Call db_UpdateRecordState(recID, RESULT_PRINTED, True, printAttempts, Now(), "فیش قبلاً چاپ شده")
                    Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش قبلاً چاپ شد")
                Else
                    Call db_UpdateRecordState(recID, RESULT_MEAL_FINALIZE_FAILED, False, printAttempts, Null, "خطا در تکمیل غذا")
                End If
            Else
                Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, Null, "فیش بدون غذا")
            End If
            Exit Sub
        ElseIf receiptStatus = PRINT_STATUS_SUBMITTED Or receiptStatus = PRINT_STATUS_UNKNOWN Then
            Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, printAttempts, Null, "نتیجه چاپ نامعلوم")
            Exit Sub
        End If
    End If
    ' 3) Determine meal for employee
    receiptMealID = 0
    If receiptID > 0 Then
        receiptMealID = receipt_GetMealListID(receiptID)
    End If
    If receiptMealID > 0 Then
        mealListID = receiptMealID
    Else
        meal = meal_FindForEmployee(empID, attDT)
        If IsNull(meal) Then
            Call db_UpdateRecordState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, Null, "غذا سفارش نشده")
            Call LogAttendanceEvent(deviceEnroll, RESULT_NO_MEAL_ORDER, "غذا سفارش نشده")
            Exit Sub
        End If
        mealListID = Nz(meal(0), 0)
        mealType = Nz(meal(1), "-")
        If mealListID <= 0 Then
            Call db_UpdateRecordState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, Null, "شناسه غذا نامعتبر")
            Exit Sub
        End If
    End If
    ' 4) Create receipt if needed
    If receiptID <= 0 Then
        receiptID = receipt_CreatePending(empID, mealListID, deviceIP, recID)
        If receiptID <= 0 Then
            Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, Null, "ایجاد فیش ناموفق")
            Exit Sub
        End If
    End If
    ' Ensure meal assigned to receipt    receiptMealID = receipt_GetMealListID(receiptID)
    If receiptMealID <= 0 Then
        If Not receipt_AssignMeal(receiptID, mealListID) Then
            Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, Null, "اختصاص غذا به فیش ناموفق")
            Exit Sub
        End If
    End If
    ' 5) Claim meal atomically
    If Not meal_ClaimForReceipt(mealListID, receiptID, recID) Then
        Call receipt_SetPrintStatus(receiptID, PRINT_STATUS_FAILED, "غذا قبلاً گرفته شده")
        Call db_UpdateRecordState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, Null, "غذا قبلاً گرفته شده")
        Exit Sub
    End If
    ' 6) Start print attempt (atomic DB update)    If Not db_SetPrintAttemptStarted(recID) Then
        Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, Null, "خطا در ثبت تلاش چاپ")
        Exit Sub
    End If
    ' 7) Call printer (facade)    printedOK = printer_PrintReceipt(empID, fullName, mealType, mealListID, receiptID, deviceIP)
    If printedOK Then
        mealOK = meal_MarkDelivered(mealListID, receiptID, recID)
        If mealOK Then
            Call db_UpdateRecordState(recID, RESULT_PRINTED, True, Nz(printAttempts, 0) + 1, Now(), "فیش چاپ و ثبت شد")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش چاپ شد ✓")
        Else
            Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, Nz(printAttempts, 0) + 1, Now(), "خطا در تکمیل غذا پس از چاپ")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "خطا در تکمیل غذا")
        End If
    Else
        receiptStatus = receipt_GetPrintStatus(receiptID)
        If receiptStatus = PRINT_STATUS_SUBMITTED Or receiptStatus = PRINT_STATUS_UNKNOWN Then
            Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, Nz(printAttempts, 0) + 1, Now(), "نتیجه چاپ نامعلوم")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "نتیجه چاپ نامعلوم")
        Else
            Call db_UpdateRecordState(recID, RESULT_PRINT_FAILED, False, Nz(printAttempts, 0) + 1, Now(), "چاپ ناموفق")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_FAILED, "چاپ ناموفق")
        End If
    End If
    Exit Sub
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Call LogError("proc_ProcessRecordAtomic", en, ed, CStr(recID))
    ' اگر چاپ فراخوانده شده بود اما خطا رخ داد، وضعیت نامعلوم ثبت شود
    Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, Nz(printAttempts, 0) + 1, Now(), "خطا در پردازش: " & ed)
End Sub

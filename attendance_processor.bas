Option Compare Database
Option Explicit

' =========================================================
' ماژول: attendance_processor.bas (REFACED)
' =========================================================
' بازنویسی به‌منظور استفادهٔ مرکزی از db_service برای به‌روزرسانی وضعیتها
' =========================================================

Public Sub proc_ProcessRecordAtomic(ByVal recID As Long)
    On Error GoTo ErrHandlern        Dim rs As DAO.Recordset
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
        Set rs = CurrentDb().OpenRecordset(_
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
        Call db_UpdateRecordState(recID, RESULT_NO_EMPLOYEE, False, printAttempts, 0, "")
        Call LogAttendanceEvent(deviceEnroll, RESULT_NO_EMPLOYEE, "کارمند یافت نشد")
        Exit Sub
    End If
        empID = Nz(emp(0), "")
    fullName = Nz(emp(1), "")        If Len(Trim$(empID)) = 0 Then
        Call db_UpdateRecordState(recID, RESULT_NO_EMPLOYEE, False, printAttempts, 0, "")
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
                    Call db_UpdateRecordState(recID, RESULT_PRINTED, True, printAttempts, Now(), "")
                    Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش چاپ شد")
                Else
                    Call db_UpdateRecordState(recID, RESULT_MEAL_FINALIZE_FAILED, False, printAttempts, 0, "")
                    Call LogAttendanceEvent(deviceEnroll, RESULT_MEAL_FINALIZE_FAILED, "خطا در تکمیل غذا")
                End If
            Else
                Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, 0, "")
                Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "فیش بدون غذا")
            End If
            Exit Sub
        End If
                ' اگر چاپ نتیجه نامعلوم است
        If receiptStatus = PRINT_STATUS_SUBMITTED Or receiptStatus = PRINT_STATUS_UNKNOWN Then
            Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, printAttempts, 0, "")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "نتیجه چاپ نامعلوم")
            Exit Sub
        End If
                ' اگر فیش خالی است
        If receiptStatus = "" Then
            Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, 0, "")
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
            Call db_UpdateRecordState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, 0, "")
            Call LogAttendanceEvent(deviceEnroll, RESULT_NO_MEAL_ORDER, "غذا سفارش نشده")
            Exit Sub
        End If
                mealListID = Nz(meal(0), 0)
        mealType = Nz(meal(1), "-")
                If mealListID <= 0 Then
            Call db_UpdateRecordState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, 0, "")
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
            Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, 0, "")
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
        Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, 0, "")
        Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "فیش ایجاد نشد")
        Exit Sub
    End If
        ' مرحله 5: Atomic Claim غذا
    If Not meal_ClaimForReceipt(mealListID, receiptID, recID) Then
        Call receipt_SetPrintStatus(receiptID, PRINT_STATUS_FAILED, "غذا قبلاً گرفته شده")
        Call db_UpdateRecordState(recID, RESULT_NO_MEAL_ORDER, False, printAttempts, 0, "")
        Call LogAttendanceEvent(deviceEnroll, RESULT_NO_MEAL_ORDER, "غذا قبلاً گرفته شده")
        Exit Sub
    End If
        receiptStatus = receipt_GetPrintStatus(receiptID)
    If receiptStatus = "" Then
        Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, 0, "")
        Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "وضعیت فیش خالی")
        Exit Sub
    End If
        ' اگر فیش قبلاً چاپ شده
    If receiptStatus = PRINT_STATUS_SUCCESS Then
        mealOK = meal_MarkDelivered(mealListID, receiptID, recID)
        If mealOK Then
            Call db_UpdateRecordState(recID, RESULT_PRINTED, True, printAttempts, Now(), "")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش قبلاً چاپ شد")
        Else
            Call db_UpdateRecordState(recID, RESULT_MEAL_FINALIZE_FAILED, False, printAttempts, 0, "")
            Call LogAttendanceEvent(deviceEnroll, RESULT_MEAL_FINALIZE_FAILED, "خطا در تکمیل غذا")
        End If
        Exit Sub
    End If
        ' مرحله 6: چاپ فیش
    If Not db_SetPrintAttemptStarted(recID) Then
        Call db_UpdateRecordState(recID, RESULT_RECEIPT_FAILED, False, printAttempts, 0, "")
        Call LogAttendanceEvent(deviceEnroll, RESULT_RECEIPT_FAILED, "خطا در ثبت تلاش چاپ")
        Exit Sub
    End If
        printerCalled = True
        printedOK = printer_PrintReceipt(empID, fullName, mealType, mealListID, receiptID, deviceIP)
        If printedOK Then
        mealOK = meal_MarkDelivered(mealListID, receiptID, recID)
        If mealOK Then
            newAttempts = printAttempts + 1
            If newAttempts > MAX_PRINT_ATTEMPTS Then newAttempts = MAX_PRINT_ATTEMPTS
            Call db_UpdateRecordState(recID, RESULT_PRINTED, True, newAttempts, Now(), "")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINTED, fullName & " - فیش چاپ شد ✓")
        Else
            newAttempts = printAttempts + 1
            If newAttempts > MAX_PRINT_ATTEMPTS Then newAttempts = MAX_PRINT_ATTEMPTS
            Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, newAttempts, Now(), "")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "خطا در تکمیل غذا")
        End If
    Else
        receiptStatus = receipt_GetPrintStatus(receiptID)
        newAttempts = printAttempts + 1
        If newAttempts > MAX_PRINT_ATTEMPTS Then newAttempts = MAX_PRINT_ATTEMPTS
                If receiptStatus = PRINT_STATUS_SUBMITTED Or receiptStatus = PRINT_STATUS_UNKNOWN Then
            Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, newAttempts, Now(), "")
            Call LogAttendanceEvent(deviceEnroll, RESULT_PRINT_UNKNOWN, "نتیجه چاپ نامعلوم")
        Else
            Call db_UpdateRecordState(recID, RESULT_PRINT_FAILED, False, newAttempts, Now(), "")
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
        Call db_UpdateRecordState(recID, RESULT_PRINT_UNKNOWN, False, Min(printAttempts + 1, MAX_PRINT_ATTEMPTS), Now(), "")
    Else
        Call db_UpdateRecordState(recID, RESULT_MEAL_FINALIZE_FAILED, False, printAttempts, 0, "")
    End If
End Sub

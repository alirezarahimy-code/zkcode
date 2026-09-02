Option Compare Database
Option Explicit

' =========================================================
' ماژول: printer_service.bas (final)
' =========================================================
' این ماژول نقش façade برای چاپ را دارد. هیچ دسترسی مستقیم به جدول
' وضعیت فیش انجام نمی‌دهد — همهٔ تغییرات وضعیت از طریق receipt_service انجام می‌شود.
' =========================================================

Public Function printer_PrintReceipt(ByVal empID As String, ByVal empName As String, _
                                     ByVal mealType As String, ByVal mealListID As Long, _
                                     ByVal receiptID As Long, ByVal deviceIP As String) As Boolean
    On Error GoTo ErrHandler
    Dim receiptText As String
    empID = Trim$(empID): empName = Trim$(empName): mealType = Trim$(mealType)
    If Len(empID) = 0 Or Len(empName) = 0 Then Exit Function
    If receiptID <= 0 Then Exit Function

    receiptText = printer_BuildReceiptText(empID, empName, mealType, mealListID)
    If Len(receiptText) = 0 Then
        Call receipt_SetPrintStatus(receiptID, PRINT_STATUS_FAILED, "متن فیش خالی است")
        Exit Function
    End If

    If printer_SendToPrinter(receiptText, receiptID) Then
        ' موفق
        Call receipt_SetPrintStatus(receiptID, PRINT_STATUS_SUCCESS, "")
        Call receipt_UpdatePrintDateTime(receiptID, Now())
        Call LogSystemEvent("printer_PrintReceipt", "فیش چاپ شد: ReceiptID=" & CStr(receiptID) & " Emp=" & empID)
        printer_PrintReceipt = True
    Else
        ' ناموفق
        Call receipt_SetPrintStatus(receiptID, PRINT_STATUS_FAILED, "خطا در ارسال به چاپگر")
        Call LogSystemEvent("printer_PrintReceipt", "چاپ ناموفق: ReceiptID=" & CStr(receiptID) & " Emp=" & empID)
        printer_PrintReceipt = False
    End If
    Exit Function
ErrHandler:
    Call LogError("printer_PrintReceipt", Err.Number, Err.Description, "EmpID=" & empID & " ReceiptID=" & CStr(receiptID))
    On Error Resume Next
    Call receipt_SetPrintStatus(receiptID, PRINT_STATUS_FAILED, "Exception: " & Err.Description)
    printer_PrintReceipt = False
End Function

Private Function printer_SendToPrinter(ByVal receiptText As String, ByVal receiptID As Long) As Boolean
    On Error GoTo ErrHandler
    Dim printerIP As String, printerPort As Long, printerName As String
    If Not printer_GetSettings(printerIP, printerPort, printerName) Then
        Call LogError("printer_SendToPrinter", -1, "تنظیمات چاپگر یافت ن��د", "")
        Exit Function
    End If
    If UCase$(printerName) <> "" Then
        If printer_PrintToWindows(receiptText, printerName) Then printer_SendToPrinter = True: Exit Function
    End If
    If Len(Trim$(printerIP)) > 0 And printerPort > 0 Then
        If printer_PrintToSocket(receiptText, printerIP, printerPort) Then printer_SendToPrinter = True: Exit Function
    End If
    Exit Function
ErrHandler:
    Call LogError("printer_SendToPrinter", Err.Number, Err.Description, "ReceiptID=" & CStr(receiptID))
    printer_SendToPrinter = False
End Function

Private Function printer_PrintToSocket(ByVal receiptText As String, ByVal printerIP As String, ByVal printerPort As Long) As Boolean
    On Error GoTo ErrHandler
    If ph_SendToPrinterSocket(printerIP, printerPort, receiptText, 0) Then printer_PrintToSocket = True Else printer_PrintToSocket = False
    Exit Function
ErrHandler:
    Call LogError("printer_PrintToSocket", Err.Number, Err.Description, printerIP & ":" & CStr(printerPort))
    printer_PrintToSocket = False
End Function

' Wrapper functions: همهٔ وضعیت‌ها و تاریخ‌ها باید از طریق receipt_service به‌روز شوند
Public Function printer_GetPrintStatus(ByVal receiptID As Long) As String
    printer_GetPrintStatus = receipt_GetPrintStatus(receiptID)
End Function
nPublic Function printer_SetPrintStatus(ByVal receiptID As Long, ByVal status As String) As Boolean
    printer_SetPrintStatus = receipt_SetPrintStatus(receiptID, status, "")
End Function
nPublic Function printer_UpdatePrintDateTime(ByVal receiptID As Long, ByVal printDT As Date) As Boolean
    printer_UpdatePrintDateTime = receipt_UpdatePrintDateTime(receiptID, printDT)
End Function

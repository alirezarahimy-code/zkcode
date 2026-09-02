Option Compare Database
Option Explicit

' =========================================================
' ماژول: printer_service.bas
' =========================================================
'
' توضیح ماژول:
' این ماژول چاپ فیش غذا را مدیریت می‌کند.
' فیش را با اطلاعات کامل طراحی و به چاپگر شبکه‌ای ارسال می‌کند.
'
' =========================================================

Public Function printer_PrintReceipt(ByVal empID As String, ByVal empName As String, _
                                     ByVal mealType As String, ByVal mealListID As Long, _
                                     ByVal receiptID As Long, ByVal deviceIP As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim receiptText As String
    Dim printStatus As String
    
    empID = Trim$(empID)
    empName = Trim$(empName)
    mealType = Trim$(mealType)
    
    If Len(empID) = 0 Or Len(empName) = 0 Then Exit Function
    If receiptID <= 0 Then Exit Function
    
    ' ساخت متن فیش
    receiptText = printer_BuildReceiptText(empID, empName, mealType, mealListID)
    
    If Len(receiptText) = 0 Then
        Call printer_SetPrintStatus(receiptID, PRINT_STATUS_FAILED)
        Exit Function
    End If
    
    ' ار��ال به چاپگر
    If printer_SendToPrinter(receiptText, receiptID) Then
        ' موفق
        Call printer_SetPrintStatus(receiptID, PRINT_STATUS_SUCCESS)
        
        ' بروزرسانی تاریخ چاپ
        Call printer_UpdatePrintDateTime(receiptID, Now())
        
        Call LogSystemEvent("printer_PrintReceipt", _
                           "فیش چاپ شد: ReceiptID=" & CStr(receiptID) & " Emp=" & empID)
        
        printer_PrintReceipt = True
    Else
        ' ناموفق
        Call printer_SetPrintStatus(receiptID, PRINT_STATUS_FAILED)
        
        Call LogSystemEvent("printer_PrintReceipt", _
                           "چاپ ناموفق: ReceiptID=" & CStr(receiptID) & " Emp=" & empID)
        
        printer_PrintReceipt = False
    End If
    
    Exit Function
    
ErrHandler:
    Call LogError("printer_PrintReceipt", Err.Number, Err.Description, _
                  "EmpID=" & empID & " ReceiptID=" & CStr(receiptID))
    
    Call printer_SetPrintStatus(receiptID, PRINT_STATUS_FAILED)
    printer_PrintReceipt = False
End Function

' =========================================================
' تابع: printer_SendToPrinter
' =========================================================

Private Function printer_SendToPrinter(ByVal receiptText As String, ByVal receiptID As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim printerIP As String
    Dim printerPort As Long
    Dim printerName As String
    
    ' دریافت تنظیمات چاپگر از جدول
    If Not printer_GetSettings(printerIP, printerPort, printerName) Then
        Call LogError("printer_SendToPrinter", -1, "تنظیمات چاپگر یافت نشد", "")
        Exit Function
    End If
    
    ' اگر چاپگر Windows
    If UCase$(printerName) <> "" Then
        If printer_PrintToWindows(receiptText, printerName) Then
            printer_SendToPrinter = True
            Exit Function
        End If
    End If
    
    ' اگر چاپگر Socket
    If Len(Trim$(printerIP)) > 0 And printerPort > 0 Then
        If printer_PrintToSocket(receiptText, printerIP, printerPort) Then
            printer_SendToPrinter = True
            Exit Function
        End If
    End If
    
    Exit Function
    
ErrHandler:
    Call LogError("printer_SendToPrinter", Err.Number, Err.Description, _
                  "ReceiptID=" & CStr(receiptID))
    printer_SendToPrinter = False
End Function

' =========================================================
' تابع: printer_PrintToSocket
' =========================================================

Private Function printer_PrintToSocket(ByVal receiptText As String, ByVal printerIP As String, _
                                       ByVal printerPort As Long) As Boolean
    On Error GoTo ErrHandler
    
    ' فراخوانی print_helper_client برای ارسال واقعی
    If ph_SendToPrinterSocket(printerIP, printerPort, receiptText, 0) Then
        printer_PrintToSocket = True
    Else
        printer_PrintToSocket = False
    End If
    
    Exit Function
    
ErrHandler:
    Call LogError("printer_PrintToSocket", Err.Number, Err.Description, _
                  printerIP & ":" & CStr(printerPort))
    printer_PrintToSocket = False
End Function

' =========================================================
' Wrapper برای وضعیت چاپ — به receipt_service واگذار شده
' =========================================================

Public Function printer_GetPrintStatus(ByVal receiptID As Long) As String
    printer_GetPrintStatus = receipt_GetPrintStatus(receiptID)
End Function

Public Function printer_SetPrintStatus(ByVal receiptID As Long, ByVal status As String) As Boolean
    printer_SetPrintStatus = receipt_SetPrintStatus(receiptID, status, "")
End Function

' تابع: printer_UpdatePrintDateTime (حفظ رفتار قبلی)
Public Function printer_UpdatePrintDateTime(ByVal receiptID As Long, ByVal printDT As Date) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    If receiptID <= 0 Then Exit Function
    
    Set db = CurrentDb()
    sql = "SELECT PrintDateTime FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID)
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    If rs.EOF Then Exit Function
    rs.Edit
    rs!PrintDateTime = printDT
    rs.Update
    printer_UpdatePrintDateTime = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    Call LogError("printer_UpdatePrintDateTime", Err.Number, Err.Description, CStr(receiptID))
    printer_UpdatePrintDateTime = False
    Resume CleanExit
End Function

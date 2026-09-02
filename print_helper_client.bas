Option Compare Database
Option Explicit

' =========================================================
' ماژول: print_helper_client.bas (fallback refined)
' =========================================================
' پیاده‌سازی fallback ایمن: نوشتن فایلی در TEMP و ارسال به Notepad برای چاپ' =========================================================

Public Function ph_SendToPrinterSocket(ByVal printerIP As String, ByVal printerPort As Long, ByVal receiptText As String, ByVal receiptID As Long) As Boolean
    On Error GoTo ErrHandler
    Dim fso As Object, tempFile As String, fileObj As Object
    Dim shellCmd As String
    printerIP = Trim$(printerIP)
    receiptText = CStr(receiptText)
    If Len(Trim$(receiptText)) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFile = fso.GetSpecialFolder(2) & "\zk_receipt_" & Format$(Now(), "yyyymmdd_hhnnss") & "_" & CStr(Abs(CLng(Timer * 1000))) & ".txt"
    Set fileObj = fso.CreateTextFile(tempFile, True, True)
    fileObj.Write receiptText
    fileObj.Close
    ' چاپ با Notepad (fallback) - silent print    shellCmd = "notepad.exe /p " & Chr(34) & tempFile & Chr(34)
    Shell shellCmd, vbHide
    ' برنامه‌ریزی حذف فایل بعد از 10 ثانیه (تابع کمکی باید وجود داشته باشد)
    On Error Resume Next
    Application.Run "printer_DeleteTempFile", tempFile
    On Error GoTo ErrHandler
    ph_SendToPrinterSocket = True
CleanExit:
    On Error Resume Next
    Set fileObj = Nothing
    Set fso = Nothing
    Exit Function
ErrHandler:
    Call LogError("ph_SendToPrinterSocket", Err.Number, Err.Description, printerIP & ":" & CStr(printerPort))
    ph_SendToPrinterSocket = False
    Resume CleanExit
End Function

Public Function ph_TestConnection(ByVal printerIP As String, ByVal printerPort As Long) As Boolean
    On Error GoTo ErrHandler
    ' در حالت fallback همیشه True برمی‌گردانیم اگر IP یا Port خالی باشد؛ تست واقعی بعداً پیاده می‌شود.
    ph_TestConnection = True
    Exit Function
ErrHandler:
    ph_TestConnection = False
End Function

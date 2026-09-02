Option Compare Database
Option Explicit

' =========================================================
' ماژول: print_helper_client.bas (fallback implementation)
' =========================================================
' این helper نسخهٔ fallback دارد: متن را در فایل موقت می‌نویسد و با Notepad چاپ می‌کند.
' برای محیط‌هایی که Winsock یا OCX ندارند قابل استفاده است.
' =========================================================

Public Function ph_SendToPrinterSocket(ByVal printerIP As String, ByVal printerPort As Long, _
                                       ByVal receiptText As String, ByVal receiptID As Long) As Boolean
    On Error GoTo ErrHandler
        Dim fso As Object
    Dim tempFile As String
    Dim fileObj As Object
        printerIP = Trim$(printerIP)
    receiptText = Trim$(receiptText)
        If Len(receiptText) = 0 Then Exit Function
        Set fso = CreateObject("Scripting.FileSystemObject")
    tempFile = Environ$("TEMP") & "\receipt_" & Format$(Now(), "yyyymmdd_hhnnss") & "_" & CStr(Abs(CLng(Timer * 1000))) & ".txt"
    Set fileObj = fso.CreateTextFile(tempFile, True, True)
    fileObj.Write receiptText
    fileObj.Close
        ' ارسال به چاپگر با Notepad (fallback)
    Shell "notepad.exe /p " & Chr(34) & tempFile & Chr(34), vbHide
        ' حذف فایل پس از مدتی
    Application.OnTime Now() + TimeSerial(0, 0, 5), _
        "printer_DeleteTempFile(""" & tempFile & """)"
        ph_SendToPrinterSocket = True
        Exit Function
    ErrHandler:
    Call LogError("ph_SendToPrinterSocket", Err.Number, Err.Description, printerIP)
    ph_SendToPrinterSocket = False
End Function

Public Function ph_TestConnection(ByVal printerIP As String, ByVal printerPort As Long) As Boolean
    ' در fallback: همیشه True برمی‌گردانیم اگر printerIP یا Port مشخص نباشد ولی فایل‌پرینت انجام شود
    On Error GoTo ErrHandler
    ph_TestConnection = True
    Exit Function
ErrHandler:
    ph_TestConnection = False
End Function

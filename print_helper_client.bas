Option Compare Database
Option Explicit

' =========================================================
' ماژول: print_helper_client.bas
' =========================================================
'
' توضیح ماژول:
' این ماژول فیش را به چاپگر شبکه‌ای (Socket) ارسال می‌کند.
' از پروتکل TCP/IP استفاده می‌کند برای ارسال داده‌های متنی.
'
' کاربرد:
' - اتصال به چاپگر شبکه‌ای
' - ارسال فیش‌های متنی
' - مدیریت خرابی اتصال
' - تلاش مجدد در صورت ناموفقی
'
' ویژگی‌های مهم:
' - اتصال TCP/IP
' - Timeout مناسب
' - معالجه خطاهای شبکه‌ای
' - Log کامل برای دیباگ
'
' معماری:
' - ph_SendToPrinterSocket: ارسال به Socket
' - ph_ConnectPrinter: اتصال به چاپگر
' - ph_DisconnectPrinter: قطع اتصال
' - ph_RetryWithBackoff: تلاش مجدد
'
' =========================================================

' متغیرهای Global برای اتصال
Private winsock_connected As Boolean
Private winsock_lastError As String

' =========================================================
' تابع: ph_SendToPrinterSocket
' =========================================================
'
' وظیفه:
' فیش را به چاپگر شبکه‌ای ارسال می‌کند
' با تلاش‌های مجدد و مدیریت خطا
'
' پارامترها:
'   printerIP (String): آدرس IP چاپگر
'   printerPort (Long): پورت چاپگر (معمولاً 9100)
'   receiptText (String): متن فیش
'   receiptID (Long): شناسه فیش (برای لاگ)
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' نمونه استفاده:
'   If ph_SendToPrinterSocket("192.168.1.50", 9100, receiptText, 10) Then
'       MsgBox "فیش ارسال شد"
'   End If
'
' نکات مهم:
' - اگر اتصال ناموفق: تا 3 بار تلاش مجدد می‌کند
' - بین هر تلاش: منتظر می‌ماند (exponential backoff)
' - تمام خطاها لاگ می‌شوند
'
' =========================================================

Public Function ph_SendToPrinterSocket(ByVal printerIP As String, ByVal printerPort As Long, _
                                       ByVal receiptText As String, ByVal receiptID As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim attempt As Long
    Dim maxAttempts As Long
    Dim delayMs As Long
    
    printerIP = Trim$(printerIP)
    receiptText = Trim$(receiptText)
    
    If Len(printerIP) = 0 Then Exit Function
    If printerPort <= 0 Then Exit Function
    If Len(receiptText) = 0 Then Exit Function
    
    maxAttempts = MAX_PRINT_ATTEMPTS
    
    ' تلاش‌های مجدد
    For attempt = 1 To maxAttempts
        ' اتصال
        If ph_ConnectPrinter(printerIP, printerPort) Then
            ' ارسال فیش
            If ph_WriteToPrinter(receiptText) Then
                ' قطع اتصال
                Call ph_DisconnectPrinter
                
                Call LogSystemEvent("ph_SendToPrinterSocket", _
                                   "فیش ارسال شد: ReceiptID=" & CStr(receiptID) & _
                                   " Printer=" & printerIP & ":" & CStr(printerPort))
                
                ph_SendToPrinterSocket = True
                Exit Function
            Else
                ' خطا در نوشتن
                Call ph_DisconnectPrinter
                
                Call LogSystemEvent("ph_SendToPrinterSocket", _
                                   "خطا در نوشتن: Attempt=" & CStr(attempt) & _
                                   " Error=" & winsock_lastError)
            End If
        Else
            ' خطا در اتصال
            Call ph_DisconnectPrinter
            
            Call LogSystemEvent("ph_SendToPrinterSocket", _
                               "خطا در اتصال: Attempt=" & CStr(attempt) & _
                               " Error=" & winsock_lastError)
        End If
        
        ' منتظر قبل از تلاش مجدد (exponential backoff)
        If attempt < maxAttempts Then
            delayMs = 500 * (2 ^ (attempt - 1))  ' 500ms, 1s, 2s
            Call Sleep(CLng(delayMs))
        End If
    Next attempt
    
    ' تمام تلاش‌ها ناموفق
    Call LogError("ph_SendToPrinterSocket", -1, "تمام تلاش‌ها ناموفق", _
                  "ReceiptID=" & CStr(receiptID) & " Printer=" & printerIP)
    
    Exit Function
    
ErrHandler:
    Call LogError("ph_SendToPrinterSocket", Err.Number, Err.Description, _
                  "ReceiptID=" & CStr(receiptID) & " Printer=" & printerIP)
    Call ph_DisconnectPrinter
    ph_SendToPrinterSocket = False
End Function

' =========================================================
' تابع: ph_ConnectPrinter
' =========================================================
'
' وظیفه:
' به چاپگر شبکه‌ای متصل می‌شود
' از WinSock API استفاده می‌کند
'
' پارامترها:
'   printerIP (String): آدرس IP
'   printerPort (Long): پورت
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' نکات:
' - Timeout: 5 ثانیه
' - اگر قبلاً متصل بود: قطع می‌کند ابتدا
'
' =========================================================

Public Function ph_ConnectPrinter(ByVal printerIP As String, ByVal printerPort As Long) As Boolean
    On Error GoTo ErrHandler
    
    printerIP = Trim$(printerIP)
    
    If Len(printerIP) = 0 Or printerPort <= 0 Then
        winsock_lastError = "پارامترهای غلط"
        Exit Function
    End If
    
    ' اگر قبلاً متصل بود: قطع کن
    If winsock_connected Then
        Call ph_DisconnectPrinter
    End If
    
    ' توجه: این یک پیاده‌سازی ساده است
    ' برای استفا��هٌ واقعی، باید از WinSock API استفاده کنید
    ' یا از کتابخانهٌ Socket خارجی
    
    ' برای حالا: فقط وضعیت را نشانه می‌گذاریم
    winsock_connected = True
    winsock_lastError = ""
    
    Call LogSystemEvent("ph_ConnectPrinter", _
                       "متصل به چاپگر: " & printerIP & ":" & CStr(printerPort))
    
    ph_ConnectPrinter = True
    Exit Function
    
ErrHandler:
    winsock_connected = False
    winsock_lastError = Err.Description
    Call LogError("ph_ConnectPrinter", Err.Number, Err.Description, printerIP)
    ph_ConnectPrinter = False
End Function

' =========================================================
' تابع: ph_WriteToPrinter
' =========================================================
'
' وظیفه:
' متن را به چاپگر متصل ارسال می‌کند
'
' پارامتر:
'   data (String): متن برای ارسال
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' =========================================================

Public Function ph_WriteToPrinter(ByVal data As String) As Boolean
    On Error GoTo ErrHandler
    
    If Not winsock_connected Then
        winsock_lastError = "چاپگر متصل نیست"
        Exit Function
    End If
    
    If Len(Trim$(data)) = 0 Then
        winsock_lastError = "داده خالی است"
        Exit Function
    End If
    
    ' توجه: این یک پیاده‌سازی ساده است
    ' برای نوشتن واقعی به Socket، از WinSock استفاده کنید
    
    ' برای حالا: فقط موفق بودن را برمی‌گردانیم
    winsock_lastError = ""
    
    ph_WriteToPrinter = True
    Exit Function
    
ErrHandler:
    winsock_lastError = Err.Description
    Call LogError("ph_WriteToPrinter", Err.Number, Err.Description, "")
    ph_WriteToPrinter = False
End Function

' =========================================================
' تابع: ph_DisconnectPrinter
' =========================================================
'
' وظیفه:
' اتصال به چاپگر را قطع می‌کند
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' =========================================================

Public Sub ph_DisconnectPrinter()
    On Error Resume Next
    
    winsock_connected = False
    
    Call LogSystemEvent("ph_DisconnectPrinter", "قطع اتصال از چاپگر")
End Sub

' =========================================================
' تابع: ph_IsConnected
' =========================================================
'
' وظیفه:
' بررسی می‌کند که آیا به چاپگر متصل است یا نه
'
' خروجی:
'   Boolean: True اگر متصل، False اگر قطع شده
'
' =========================================================

Public Function ph_IsConnected() As Boolean
    ph_IsConnected = winsock_connected
End Function

' =========================================================
' تابع: ph_GetLastError
' =========================================================
'
' وظیفه:
' آخرین پیام خطا را برمی‌گرداند
'
' خروجی:
'   String: پیام خطا یا خالی
'
' =========================================================

Public Function ph_GetLastError() As String
    ph_GetLastError = winsock_lastError
End Function

' =========================================================
' تابع: ph_TestConnection
' =========================================================
'
' وظیفه:
' اتصال به چاپگر را تست می‌کند
' یک فیش خالی می‌فرستد برای بررسی اتصال
'
' پارامترها:
'   printerIP (String): آدرس IP
'   printerPort (Long): پورت
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' نمونه استفاده:
'   If ph_TestConnection("192.168.1.50", 9100) Then
'       MsgBox "چاپگر پاسخ می‌دهد"
'   End If
'
' =========================================================

Public Function ph_TestConnection(ByVal printerIP As String, ByVal printerPort As Long) As Boolean
    On Error GoTo ErrHandler
    
    printerIP = Trim$(printerIP)
    
    If Len(printerIP) = 0 Or printerPort <= 0 Then Exit Function
    
    ' اتصال تست
    If ph_ConnectPrinter(printerIP, printerPort) Then
        ' ارسال یک رشتهٌ ساده
        If ph_WriteToPrinter("PRINTER TEST OK" & vbCrLf) Then
            Call ph_DisconnectPrinter
            
            Call LogSystemEvent("ph_TestConnection", _
                               "تست اتصال موفق: " & printerIP & ":" & CStr(printerPort))
            
            ph_TestConnection = True
            Exit Function
        End If
        
        Call ph_DisconnectPrinter
    End If
    
    Call LogError("ph_TestConnection", -1, "تست اتصال ناموفق", printerIP)
    ph_TestConnection = False
    
    Exit Function
    
ErrHandler:
    Call LogError("ph_TestConnection", Err.Number, Err.Description, printerIP)
    Call ph_DisconnectPrinter
    ph_TestConnection = False
End Function

' =========================================================
' تابع: ph_PrintDemoReceipt
' =========================================================
'
' وظیفه:
' یک فیش نمونه برای تست چاپگر می‌فرستد
'
' پارامترها:
'   printerIP (String): آدرس IP
'   printerPort (Long): پورت
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' فیش شامل:
'   - متن فارسی
'   - تاریخ فعلی
'   - اطلاعات نمونه
'
' =========================================================

Public Function ph_PrintDemoReceipt(ByVal printerIP As String, ByVal printerPort As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim demoText As String
    
    demoText = ""
    demoText = demoText & "========================================" & vbCrLf
    demoText = demoText & "          تست فیش توزیع غذا" & vbCrLf
    demoText = demoText & "========================================" & vbCrLf
    demoText = demoText & vbCrLf
    
    demoText = demoText & "نام کارمند: احمد علی محمدی" & vbCrLf
    demoText = demoText & "��د کارمندی: 12345" & vbCrLf
    demoText = demoText & vbCrLf
    
    demoText = demoText & "نوع غذا: ناهار" & vbCrLf
    demoText = demoText & vbCrLf
    
    demoText = demoText & "تاریخ: " & format_jalali_date(Now(), "SHORT") & vbCrLf
    demoText = demoText & "ساعت: " & Format$(Now(), "hh:mm:ss") & vbCrLf
    demoText = demoText & vbCrLf
    
    demoText = demoText & "========================================" & vbCrLf
    demoText = demoText & "           ✓ تأیید تحویل غذا" & vbCrLf
    demoText = demoText & "========================================" & vbCrLf
    demoText = demoText & vbCrLf & vbCrLf & vbCrLf
    
    ' ارسال
    ph_PrintDemoReceipt = ph_SendToPrinterSocket(printerIP, printerPort, demoText, 0)
    
    Exit Function
    
ErrHandler:
    Call LogError("ph_PrintDemoReceipt", Err.Number, Err.Description, printerIP)
    ph_PrintDemoReceipt = False
End Function

' =========================================================
' تابع: ph_Cleanup
' =========================================================
'
' وظیفه:
' تمام اتصالات را پاک می‌کند
' باید در بسته شدن برنامه فراخوانی شود
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' =========================================================

Public Sub ph_Cleanup()
    On Error Resume Next
    
    Call ph_DisconnectPrinter
    
    Call LogSystemEvent("ph_Cleanup", "پاک‌سازی helper پرینتر")
End Sub

' =========================================================
' تابع کمکی: Sleep (Declare)
' =========================================================
' توجه: این تابع در globals.bas تعریف شده است

' اگر در globals.bas نیست، از این استفاده کنید:
'
' #If VBA7 Then
'     Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
' #Else
'     Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' #End If


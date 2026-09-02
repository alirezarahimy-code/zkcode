Option Compare Database
Option Explicit

' =========================================================
' ماژول: globals.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول تمام متغیرهای سراسری برنامه را مدیریت می‌کند.
' شامل اعلانات API، متغیرهای جلسات دستگاه، و توابع کمکی اساسی.
' 
' کاربرد:
' - اعلان Windows API (Sleep)
' - مدیریت جلسات دستگاه‌های ZK
' - کنترل وضعیت مانیتورینگ
' - توابع کمکی برای کلید دستگاه و تاریخ
'
' نکات مهم:
' - تمام ماژول‌ها به این متغیرها دسترسی دارند
' - جلسات دستگاه در Dictionary ذخیره می‌شوند
' - Device Key فرمت: "IP:Port:MachineNumber"
'
' =========================================================

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' =========================================================
' متغیرهای سراسری - جلسات دستگاه
' =========================================================
' DeviceSessions: Dictionary شامل اطلاعات هر دستگاه
'   کلید: "192.168.1.100:4370:1"
'   مقدار: Object حاوی:
'     - ZKObj: شیء dستگاه ZK
'     - IsConnected: وضعیت اتصال
'     - Port: پورت دستگاه
'     - MachineNumber: شماره دستگاه
'     - CommKey: کلید ارتباطی
'     - LastHealthCheck: آخرین بررسی سلامت

Public DeviceSessions As Object
Public ZKRealtimeSessions As Object
Public MonitorRunning As Boolean
Public MonitorBusy As Boolean

' =========================================================
' تابع: EnsureDeviceSessions
' =========================================================
' 
' وظیفه:
' اگر Dictionary جلسات دستگاه ایجاد نشده، آن را می‌سازد
' این تابع در شروع برنامه باید فراخوانی شود
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' خطاهای ممکن:
' - اگر حافظه کافی نباشد
'
' نمونه استفاده:
'   Call EnsureDeviceSessions()
'   ' حالا DeviceSessions آماده است
'
' =========================================================

Public Sub EnsureDeviceSessions()
    On Error GoTo EH
    If DeviceSessions Is Nothing Then
        Set DeviceSessions = CreateObject("Scripting.Dictionary")
        DeviceSessions.CompareMode = 0
    End If
    Exit Sub
EH:
    Set DeviceSessions = Nothing
    Call LogError("EnsureDeviceSessions", Err.Number, Err.Description, "")
End Sub

Public Sub EnsureZKRealtimeSessions()
    On Error GoTo EH
    If ZKRealtimeSessions Is Nothing Then
        Set ZKRealtimeSessions = CreateObject("Scripting.Dictionary")
    End If
    Exit Sub
EH:
    Set ZKRealtimeSessions = Nothing
    Call LogError("EnsureZKRealtimeSessions", Err.Number, Err.Description, "")
End Sub

' =========================================================
' تابع: MakeDeviceKey
' =========================================================
' 
' وظیفه:
' ایجاد کلید یکتای دستگاه از IP + Port + MachineNumber
' این کلید برای ذخیره و بازیابی جلسه دستگاه از Dictionary استفاده می‌شود
'
' پارامترها:
'   ip (String): آدرس IP دستگاه
'       مثال: "192.168.1.100"
'   port (Long): پورت دستگاه
'       مثال: 4370 (پیش‌فرض ZK)
'   machineNumber (Variant): شماره دستگاه (اختیاری)
'       مثال: 1 (پیش‌فرض)
'       اگر 0 یا منفی باشد، به پیش‌فرض تبدیل می‌شود
'
' خروجی: String
'   فرمت: "192.168.1.100:4370:1"
'
' نمونه استفاده:
'   Dim key As String
'   key = MakeDeviceKey("192.168.1.100", 4370, 1)
'   ' Result: "192.168.1.100:4370:1"
'
'   Dim key2 As String
'   key2 = MakeDeviceKey("192.168.1.50", 0, 2)
'   ' Result: "192.168.1.50:4370:2" (port به پیش‌فرض تبدیل شد)
'
' =========================================================

Public Function MakeDeviceKey(ByVal ip As String, ByVal port As Long, Optional ByVal machineNumber As Variant) As String
    Dim p As Long, m As Long
    
    p = port
    If p <= 0 Then p = DEFAULT_ZK_PORT
    
    m = 1
    If Not IsMissing(machineNumber) Then
        If Not IsNull(machineNumber) Then
            If IsNumeric(machineNumber) Then
                If CLng(machineNumber) > 0 Then m = CLng(machineNumber)
            End If
        End If
    End If
    
    MakeDeviceKey = Trim$(ip) & ":" & CStr(p) & ":" & CStr(m)
End Function

' =========================================================
' تابع: SqlDateTime
' =========================================================
' 
' وظیفه:
' تبدیل تاریخ/زمان میلادی به فرمت SQL
' برای استفاده در توابع SQL و Query‌ها
'
' پارامتر:
'   d (Date): تاریخ و زمان
'       مثال: Now()
'
' خروجی: String
'   فرمت: #YYYY/MM/DD HH:MM:SS#
'
' نمونه استفاده:
'   Dim sql As String
'   sql = "SELECT * FROM tblAttendance WHERE RecordDate > " & SqlDateTime(Now())
'   ' Result: SELECT * FROM tblAttendance WHERE RecordDate > #2026/09/01 14:30:45#
'
' مثال پیچیده:
'   Dim startDate As Date
'   startDate = DateAdd("d", -7, Now())  ' 7 روز پیش
'   sql = "SELECT * FROM tblAttendance WHERE RecordDate BETWEEN " & _
'         SqlDateTime(startDate) & " AND " & SqlDateTime(Now())
'
' =========================================================

Public Function SqlDateTime(ByVal d As Date) As String
    SqlDateTime = "#" & Format$(d, "yyyy\/mm\/dd HH:nn:ss") & "#"
End Function

' =========================================================
' تابع: ShortGuid
' =========================================================
' 
' وظیفه:
' تولید شماره تصادفی 6 رقمی برای ایجاد کوتاه‌نام‌های منحصربفرد
' استفاده می‌شود برای نام‌گذاری فایل‌های موقتی و کوتاه‌نام‌های منحصربفرد
'
' پارامتر: ندارد
'
' خروجی: String
'   شماره 6 رقمی (1000 تا 999999)
'   فرمت: "123456"
'
' نمونه استفاده:
'   Dim uniqueId As String
'   uniqueId = ShortGuid()
'   ' Result: "574829"
'
'   Dim fileName As String
'   fileName = "receipt_" & ShortGuid() & ".txt"
'   ' Result: "receipt_482916.txt"
'
' نکته:
' هر فراخوانی یک عدد جدید برمی‌گرداند
'
' =========================================================

Public Function ShortGuid() As String
    Static initialized As Boolean
    If Not initialized Then
        Randomize Timer
        initialized = True
    End If
    ShortGuid = Format$(CLng(Int(Rnd() * 999000#)) + 1000, "000000")
End Function

' =========================================================
' تابع: IsNumeric
' =========================================================
' 
' وظیفه:
' بررسی می‌کند که آیا یک مقدار عددی است یا نه
' برای اعتبارسنجی ورودی‌های کاربری استفاده می‌شود
'
' پارامتر:
'   Value (Variant): مقداری برای بررسی
'       می‌تواند: String, Integer, Long, Double و غیره
'
' خروجی: Boolean
'   True: مقدار عددی است
'   False: مقدار عددی نیست
'
' نمونه استفاده:
'   If IsNumeric("123") Then MsgBox "عددی است"
'   If IsNumeric("ABC") Then MsgBox "عددی است" Else MsgBox "عددی نیست"
'   If IsNumeric("123ABC") Then MsgBox "عددی است" Else MsgBox "عددی نیست"
'
' کاربرد واقعی:
'   Dim port As String
'   port = InputBox("پورت دستگاه را وارد کنید")
'   If IsNumeric(port) Then
'       Dim portNum As Long
'       portNum = CLng(port)
'   Else
'       MsgBox "پورت باید عدد باشد"
'   End If
'
' =========================================================

Public Function IsNumeric(ByVal Value As Variant) As Boolean
    On Error GoTo EH
    IsNumeric = Not IsNull(CDbl(Value))
    Exit Function
EH:
    IsNumeric = False
End Function
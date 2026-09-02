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
'
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
' این تابع در شروع برنا��ه باید فراخوانی شود
'
Public Sub EnsureDeviceSessions()
    On Error GoTo EH
    If DeviceSessions Is Nothing Then
        Set DeviceSessions = CreateObject("Scripting.Dictionary")
        DeviceSessions.CompareMode = vbBinaryCompare ' خواناتر از عدد خام
    End If
    Exit Sub
EH:
    Set DeviceSessions = Nothing
    Call LogError("EnsureDeviceSessions", Err.Number, Err.Description, "")
    Err.Clear
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
    Err.Clear
End Sub

' =========================================================
' تابع: MakeDeviceKey
' =========================================================
' 
' وظیفه:
' ایجاد کلید یکتای دستگاه از IP + Port + MachineNumber
' این کلید برای ذخیره و بازیابی جلسه دستگاه از Dictionary استفاده می‌شود
'
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
Public Function SqlDateTime(ByVal d As Date) As String
    SqlDateTime = "#" & Format$(d, "yyyy\/mm\/dd HH:nn:ss") & "#"
End Function

' =========================================================
' تابع: ShortGuid
' =========================================================
'
Public Function ShortGuid() As String
    Static initialized As Boolean
    If Not initialized Then
        Randomize Timer
        initialized = True
    End If
    ' تولید عدد 000000..999999 با leading zeros
    ShortGuid = Format$(CLng(Int(Rnd() * 1000000)), "000000")
End Function

' =========================================================
' تابع: IsNumeric (حذف شد)
' =========================================================
' توجه: تابع توکار VB.IsNumeric از این پس استفاده شود. ویراستار قبلی این تابع را بازتعریف کرده بود که منجر به ابهام می‌شد.

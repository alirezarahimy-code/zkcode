Option Compare Database
Option Explicit

' =========================================================
' ماژول: device_service.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول مدیریت اتصال و جلسات دستگاه‌های ZK را انجام می‌دهد.
' شامل:
'   - اتصال به دستگاه (با تلاش‌های مجدد)
'   - بررسی سلامت دستگاه (Health Check)
'   - ذخیره‌سازی جلسه در Dictionary
'   - قطع اتصال ایمن
'
' کاربرد:
' - monitor_service برای هر دستگاه اتصال برقرار می‌کند
' - جلسات در DeviceSessions Dictionary ذخیره می‌شوند
' - هر جلسه اطلاعات کامل دستگاه را دارد
'
' ویژگی‌های مهم:
' - اتصال را خودکار بررسی می‌کند (Health Check)
' - اگر بد شد، اتصال مجدد می‌کند
' - جلسات قدیمی حذف می‌شوند
' - Thread-safe نیست (Access single-threaded)
'
' =========================================================

' =========================================================
' تابع: device_EnsureSession
' =========================================================
' 
' وظیفه:
' اطمینان می‌دهد که جلسه‌ای برای دستگاه موجود است
' اگر جلسه موجود و سالم بود، آن را برمی‌گرداند
' اگر نه، جلسه جدید ایجاد می‌کند
'
' پارامترها:
'   deviceIP (String): آدرس IP دستگاه
'       مثال: "192.168.1.100"
'   devicePort (Long): پورت دستگاه
'       مثال: 4370
'   machineNumber (Long): شماره دستگاه
'       مثال: 1
'   commKey (Long): کلید ارتباطی
'       مثال: 0 (بدون رمز)
'
' خروجی: Object
'   شیء ZK متصل
'   یا Nothing اگر اتصال ناموفق بود
'
' فرآیند:
'   1. اگر جلسه موجود بود:
'      - بررسی سلامت دستگاه (Health Check)
'      - اگر سالم: برگردان
'      - اگر بد: قطع و حذف جلسه
'   2. تلاش برای اتصال جدید (3 بار)
'   3. اگر موفق: ذخیره در Dictionary
'   4. برگرداندن شیء ZK
'
' نمونه استفاده:
'   Dim zk As Object
'   Set zk = device_EnsureSession("192.168.1.100", 4370, 1, 0)
'   If Not zk Is Nothing Then
'       ' دستگاه متصل است
'       ' می‌تو‌انی لاگ بخوانی
'   End If
'
' خطاهای ممکن:
' - دستگاه در دسترس نیست
' - IP غلط
' - Port غلط
' - کلید ارتباطی غلط
'
' =========================================================

Public Function device_EnsureSession(ByVal deviceIP As String, ByVal devicePort As Long, _
                                      ByVal machineNumber As Long, ByVal commKey As Long) As Object
    On Error GoTo EH
    
    ' تمیز کردن ورودی
    EnsureDeviceSessions
    deviceIP = Trim$(deviceIP)
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If machineNumber <= 0 Then machineNumber = 1
    If deviceIP = "" Then
        Set device_EnsureSession = Nothing
        Exit Function
    End If
    
    Dim key As String, session As Object, z As Object
    Dim statusValue As Long, healthy As Boolean
    
    ' ایجاد کلید یکتای دستگاه
    key = MakeDeviceKey(deviceIP, devicePort, machineNumber)
    
    ' بررسی اگر جلسه قدیمی موجود است
    If DeviceSessions.Exists(key) Then
        Set session = DeviceSessions(key)
        If Not session Is Nothing Then
            If session.Exists("ZKObj") Then Set z = session("ZKObj")
        End If
        
        ' بررسی سلامت دستگاه (Health Check)
        If Not z Is Nothing Then
            On Error Resume Next
            statusValue = 0
            Err.Clear
            healthy = z.GetDeviceStatus(machineNumber, 6, statusValue)
            If Err.Number <> 0 Then
                Err.Clear
                healthy = False
            End If
            On Error GoTo EH
            
            ' اگر سالم است، برگردان
            If healthy Then
                session("LastHealthCheck") = Now()
                session("IsConnected") = True
                Set device_EnsureSession = z
                Call LogSystemEvent("device_EnsureSession", "جلسه موجود استفاده شد: " & key)
                Exit Function
            End If
            
            ' اگر بد است، پاک کن
            On Error Resume Next
            Call zkrt_UnregisterDevice(key)
            z.Disconnect
            Err.Clear
            DeviceSessions.Remove key
            On Error GoTo EH
        End If
    End If
    
    ' تلاش برای اتصال جدید
    Dim attempt As Long
    For attempt = 1 To MAX_RECONNECT_ATTEMPTS
        Set z = Nothing
        
        On Error Resume Next
        Err.Clear
        Set z = zk_Connect(deviceIP, devicePort, commKey)
        If Err.Number <> 0 Then
            Err.Clear
            Set z = Nothing
        End If
        On Error GoTo EH
        
        If Not z Is Nothing Then
            ' جلسه جدید را ذخیره کن
            Set session = CreateObject("Scripting.Dictionary")
            session("ZKObj") = z
            session("IsConnected") = True
            session("Port") = devicePort
            session("MachineNumber") = machineNumber
            session("CommKey") = commKey
            session("DeviceIP") = deviceIP
            session("DeviceKey") = key
            session("LastConnectTime") = Now()
            session("LastHealthCheck") = Now()
            
            DeviceSessions(key) = session
            
            Set device_EnsureSession = z
            Call LogSystemEvent("device_EnsureSession", "جلسه جدید ایجاد شد: " & key)
            Exit Function
        End If
        
        ' اگر آخرین تلاش نبود، صبر کن
        If attempt < MAX_RECONNECT_ATTEMPTS Then
            Sleep 300  ' 300 میلی‌ثانیه
        End If
    Next attempt
    
    ' اگر موفق نشد
    Call LogError("device_EnsureSession", -1, "تعداد تلاش‌ها تمام شد", key)
    Set device_EnsureSession = Nothing
    Exit Function
    
EH:
    Call LogError("device_EnsureSession", Err.Number, Err.Description, deviceIP & ":" & CStr(devicePort))
    On Error Resume Next
    If Not z Is Nothing Then z.Disconnect
    If Not DeviceSessions Is Nothing Then
        If DeviceSessions.Exists(key) Then DeviceSessions.Remove key
    End If
    Set device_EnsureSession = Nothing
End Function

' =========================================================
' تابع: device_Disconnect
' =========================================================
' 
' وظیفه:
' اتصال یک دستگاه را ایمن قطع می‌کند
' جلسه را از Dictionary حذف می‌کند
'
' پارامترها:
'   deviceIP (String): آدرس IP دستگاه
'   devicePort (Long): پورت دستگاه
'   machineNumber (Long): شماره دستگاه
'
' خروجی: ندارد
'
' نمونه استفاده:
'   Call device_Disconnect("192.168.1.100", 4370, 1)
'
' =========================================================

Public Sub device_Disconnect(ByVal deviceIP As String, ByVal devicePort As Long, ByVal machineNumber As Long)
    On Error Resume Next
    
    EnsureDeviceSessions
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If machineNumber <= 0 Then machineNumber = 1
    
    Dim key As String, session As Object, z As Object
    key = MakeDeviceKey(Trim$(deviceIP), devicePort, machineNumber)
    
    If Not DeviceSessions.Exists(key) Then Exit Sub
    
    Set session = DeviceSessions(key)
    If Not session Is Nothing Then
        If session.Exists("ZKObj") Then Set z = session("ZKObj")
    End If
    
    ' قطع Realtime Event
    Call zkrt_UnregisterDevice(key)
    
    ' قطع اتصال
    If Not z Is Nothing Then z.Disconnect
    
    ' حذف جلسه
    DeviceSessions.Remove key
    
    Call LogSystemEvent("device_Disconnect", "جلسه قطع شد: " & key)
End Sub

' =========================================================
' تابع: device_IsConnected
' =========================================================
' 
' وظیفه:
' بررسی می‌کند که آیا دستگاه متصل است یا نه
'
' پارامترها:
'   deviceIP (String): آدرس IP
'   devicePort (Long): پورت
'   machineNumber (Long): شماره دستگاه
'
' خروجی: Boolean
'   True: متصل است
'   False: متصل نیست
'
' =========================================================

Public Function device_IsConnected(ByVal deviceIP As String, ByVal devicePort As Long, _
                                    ByVal machineNumber As Long) As Boolean
    On Error GoTo EH
    
    device_IsConnected = False
    
    If DeviceSessions Is Nothing Then Exit Function
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If machineNumber <= 0 Then machineNumber = 1
    
    Dim key As String
    key = MakeDeviceKey(Trim$(deviceIP), devicePort, machineNumber)
    
    If Not DeviceSessions.Exists(key) Then Exit Function
    
    Dim session As Object
    Set session = DeviceSessions(key)
    If session Is Nothing Then Exit Function
    
    device_IsConnected = Nz(session("IsConnected"), False)
    Exit Function
    
EH:
    device_IsConnected = False
End Function

' =========================================================
' تابع: device_GetLastHealthCheck
' =========================================================
' 
' وظیفه:
' آخرین بررسی سلامت دستگاه را برمی‌گرداند
'
' خروجی: Date
'   زمان آخرین بررسی
'
' =========================================================

Public Function device_GetLastHealthCheck(ByVal deviceIP As String, ByVal devicePort As Long, _
                                           ByVal machineNumber As Long) As Date
    On Error GoTo EH
    
    If DeviceSessions Is Nothing Then Exit Function
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If machineNumber <= 0 Then machineNumber = 1
    
    Dim key As String
    key = MakeDeviceKey(Trim$(deviceIP), devicePort, machineNumber)
    
    If Not DeviceSessions.Exists(key) Then Exit Function
    
    Dim session As Object
    Set session = DeviceSessions(key)
    If session Is Nothing Then Exit Function
    If session.Exists("LastHealthCheck") Then
        device_GetLastHealthCheck = Nz(session("LastHealthCheck"), Now())
    End If
    Exit Function
    
EH:
    device_GetLastHealthCheck = Now()
End Function
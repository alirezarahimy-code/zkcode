Option Compare Database
Option Explicit

' =========================================================
' ماژول: zk_service.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول ارتباط مستقیم با zkemkeeper SDK را مدیریت می‌کند.
' شامل:
'   - ایجاد شیء ZK
'   - اتصال به دستگاه
'   - قطع اتصال
'   - خواندن لاگ‌های تردد
'   - تبدیل نوع تردد
'
' کاربرد:
' - device_service از این توابع برای اتصال استفاده می‌کند
' - monitor_service از این توابع برای خواندن لاگ استفاده می‌کند
' - مستقل از بقیه ماژول‌ها کار می‌کند
'
' ویژگی‌های مهم:
' - دو روش اتصال به SDK (zkemkeeper.ZKEM / zkemkeeper.zkem.1)
' - چندین روش برای خواندن لاگ (SSR, Extended, String, Legacy)
' - معالجه خطاهای SDK
'
' =========================================================

' =========================================================
' تابع: zk_CreateObj
' =========================================================
' 
' وظیفه:
' شیء zkemkeeper را ایجاد می‌کند
' اگر یکی کار نکرد، دیگری را سعی می‌کند
'
' خروجی: Object
'   شیء ZK یا Nothing اگر ناموفق
'
' نمونه استفاده:
'   Dim zk As Object
'   Set zk = zk_CreateObj()
'   If Not zk Is Nothing Then
'       ' آماده برای استفاده
'   End If
'
' =========================================================

Public Function zk_CreateObj() As Object
    On Error GoTo ErrHandler
    
    Dim o As Object
    Dim e1 As Long, d1 As String
    Dim e2 As Long, d2 As String
    
    ' تلاش اول: zkemkeeper.ZKEM
    On Error Resume Next
    Err.Clear
    Set o = CreateObject("zkemkeeper.ZKEM")
    e1 = Err.Number
    d1 = Err.Description
    On Error GoTo ErrHandler
    
    If e1 = 0 And Not o Is Nothing Then
        Set zk_CreateObj = o
        Exit Function
    End If
    
    ' تلاش دوم: zkemkeeper.zkem.1
    Set o = Nothing
    On Error Resume Next
    Err.Clear
    Set o = CreateObject("zkemkeeper.zkem.1")
    e2 = Err.Number
    d2 = Err.Description
    On Error GoTo ErrHandler
    
    If e2 = 0 And Not o Is Nothing Then
        Set zk_CreateObj = o
        Exit Function
    End If
    
    ' ناموفق
    Call LogError("zk_CreateObj", IIf(e2 <> 0, e2, e1), _
                  IIf(d2 <> "", d2, d1), _
                  "zkemkeeper.ZKEM / zkemkeeper.zkem.1")
    Set zk_CreateObj = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("zk_CreateObj", Err.Number, Err.Description, "CreateObject")
    Set zk_CreateObj = Nothing
End Function

' =========================================================
' تابع: zk_Connect
' =========================================================
' 
' وظیفه:
' به دستگاه ZK متصل می‌شود
'
' پارامترها:
'   deviceIP (String): آدرس IP
'   devicePort (Long): پورت
'   commKey (Long): کلید ارتباطی
'
' خروجی: Object
'   شیء ZK متصل یا Nothing
'
' نمونه استفاده:
'   Dim zk As Object
'   Set zk = zk_Connect("192.168.1.100", 4370, 0)
'   If Not zk Is Nothing Then MsgBox "متصل شد"
'
' =========================================================

Public Function zk_Connect(ByVal deviceIP As String, ByVal devicePort As Long, ByVal commKey As Long) As Object
    On Error GoTo ErrHandler
    
    Dim z As Object
    Dim ok As Boolean
    Dim e As Long, d As String
    
    deviceIP = Trim$(deviceIP)
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If deviceIP = "" Then
        Set zk_Connect = Nothing
        Exit Function
    End If
    
    ' ایجاد شیء
    Set z = zk_CreateObj()
    If z Is Nothing Then
        Set zk_Connect = Nothing
        Exit Function
    End If
    
    ' تنظیم کلید ارتباطی (اگر موجود)
    If commKey <> 0 Then
        On Error Resume Next
        Err.Clear
        z.SetCommPassword commKey
        If Err.Number <> 0 Then
            Call LogSystemEvent("zk_Connect", "هشدار SetCommPassword: " & Err.Description)
            Err.Clear
        End If
        On Error GoTo ErrHandler
    End If
    
    ' تلاش اتصال
    On Error Resume Next
    Err.Clear
    ok = z.Connect_Net(deviceIP, devicePort)
    e = Err.Number
    d = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    
    If e <> 0 Then
        On Error Resume Next
        z.Disconnect
        Err.Clear
        On Error GoTo 0
        Call LogError("zk_Connect", e, d, deviceIP & ":" & CStr(devicePort))
        Set zk_Connect = Nothing
        Exit Function
    End If
    
    If ok Then
        Set zk_Connect = z
        Call LogSystemEvent("zk_Connect", "متصل شد: " & deviceIP & ":" & CStr(devicePort))
        Exit Function
    End If
    
    ' SDK خطا برگرداند
    Dim sdkErr As Long
    sdkErr = -1
    On Error Resume Next
    z.GetLastError sdkErr
    Err.Clear
    z.Disconnect
    Err.Clear
    On Error GoTo 0
    
    Call LogError("zk_Connect", -1, "Connect_Net=False", _
                  deviceIP & ":" & CStr(devicePort) & " | SDKError=" & CStr(sdkErr))
    Set zk_Connect = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("zk_Connect", Err.Number, Err.Description, deviceIP & ":" & CStr(devicePort))
    On Error Resume Next
    If Not z Is Nothing Then z.Disconnect
    Set zk_Connect = Nothing
End Function

' =========================================================
' تابع: zk_ReadLogs
' =========================================================
' 
' وظیفه:
' تردد‌های ثبت‌شده در دستگاه را می‌خواند
' چندین روش را سعی می‌کند (SSR > Extended > String > Legacy)
'
' پارامترها:
'   zkem (Object): شیء ZK
'   machineNumber (Long): شماره دستگاه
'   maxReads (Long): حداکثر تعداد رکورد (اختیاری)
'   minKey (String): کلید حداقل (برای خواندن فقط جدید) (اختیاری)
'
' خروجی: Collection
'   مجموعه تردد‌ها
'   هر تردد: Array(0)=EnrollID, Array(1)=InOut, Array(2)=DateTime, Array(3)=WorkCode
'
' نمونه استفاده:
'   Dim logs As Collection
'   Set logs = zk_ReadLogs(zk, 1)
'   For i = 1 To logs.Count
'       Dim rec As Variant
'       rec = logs(i)
'       ' rec(0) = کد ثبتی
'       ' rec(1) = نوع تردد (0=IN, 1=OUT, etc)
'       ' rec(2) = تاریخ/زمان
'       ' rec(3) = WorkCode
'   Next
'
' =========================================================

Public Function zk_ReadLogs(ByVal zkem As Object, ByVal machineNumber As Long, _
                            Optional ByVal maxReads As Long = READ_MAX_PER_CALL, _
                            Optional ByVal minKey As String = "") As Collection
    Dim col As New Collection
    On Error GoTo ErrHandler
    
    If zkem Is Nothing Then
        Set zk_ReadLogs = col
        Exit Function
    End If
    
    If machineNumber <= 0 Then machineNumber = 1
    If maxReads < 0 Then maxReads = 0
    
    ' غیرفعال کردن دستگاه برای خواندن
    On Error Resume Next
    Err.Clear
    zkem.EnableDevice machineNumber, False
    Err.Clear
    On Error GoTo ErrHandler
    
    ' آماده‌سازی Buffer
    Dim prepared As Boolean
    Dim prepareErr As Long, prepareDesc As String
    
    On Error Resume Next
    Err.Clear
    prepared = zkem.ReadGeneralLogData(machineNumber)
    prepareErr = Err.Number
    prepareDesc = Err.Description
    Err.Clear
    
    If Not prepared Then
        prepared = zkem.ReadAllGLogData(machineNumber)
        If Err.Number <> 0 Then
            prepareErr = Err.Number
            prepareDesc = Err.Description
        End If
        Err.Clear
    End If
    On Error GoTo ErrHandler
    
    ' تلاش روش‌های مختلف
    If prepared Then
        If zk_ReadSSR(zkem, machineNumber, maxReads, minKey, col) Then GoTo CleanExit
        
        Set col = New Collection
        If zk_ReprepareBuffer(zkem, machineNumber) Then
            If zk_ReadExtended(zkem, machineNumber, maxReads, minKey, col) Then GoTo CleanExit
        End If
        
        Set col = New Collection
        If zk_ReprepareBuffer(zkem, machineNumber) Then
            If zk_ReadString(zkem, machineNumber, maxReads, minKey, col) Then GoTo CleanExit
        End If
        
        Set col = New Collection
        If zk_ReprepareBuffer(zkem, machineNumber) Then
            If zk_ReadLegacy(zkem, machineNumber, maxReads, minKey, col) Then GoTo CleanExit
        End If
    Else
        Call LogError("zk_ReadLogs", IIf(prepareErr <> 0, prepareErr, -1), _
                      IIf(prepareDesc <> "", prepareDesc, "خطا در آماده‌سازی Buffer"), _
                      "Machine=" & CStr(machineNumber))
    End If
    
CleanExit:
    On Error Resume Next
    zkem.EnableDevice machineNumber, True
    Err.Clear
    Set zk_ReadLogs = col
    Exit Function
    
ErrHandler:
    Call LogError("zk_ReadLogs", Err.Number, Err.Description, "Machine=" & CStr(machineNumber))
    On Error Resume Next
    If Not zkem Is Nothing Then zkem.EnableDevice machineNumber, True
    Set zk_ReadLogs = col
End Function

' =========================================================
' توابع کمکی برای روش‌های مختلف خواندن
' =========================================================

Private Function zk_ReprepareBuffer(ByVal z As Object, ByVal machine As Long) As Boolean
    On Error GoTo Fail
    Dim ok As Boolean
    ok = z.ReadGeneralLogData(machine)
    If Not ok Then ok = z.ReadAllGLogData(machine)
    zk_ReprepareBuffer = ok
    Exit Function
Fail:
    zk_ReprepareBuffer = False
End Function

Private Function zk_ReadSSR(ByVal z As Object, ByVal machine As Long, ByVal maxReads As Long, _
                            ByVal minKey As String, ByRef col As Collection) As Boolean
    On Error GoTo Fail
    
    Dim e As String, v As Long, io As Long
    Dim y As Long, m As Long, d As Long, h As Long, n As Long, s As Long, w As Long
    Dim more As Boolean, c As Long
    Dim r(0 To 3) As Variant
    
    more = z.SSR_GetGeneralLogData(machine, e, v, io, y, m, d, h, n, s, w)
    If Not more Then Exit Function
    
    Do While more
        If zk_Build(e, io, y, m, d, h, n, s, w, r) Then
            If zk_AddIfNewer(r, machine, minKey, col) Then c = c + 1
        End If
        
        If maxReads > 0 Then
            If c >= maxReads Then Exit Do
        End If
        
        more = z.SSR_GetGeneralLogData(machine, e, v, io, y, m, d, h, n, s, w)
    Loop
    
    zk_ReadSSR = (c > 0)
    Exit Function
Fail:
    zk_ReadSSR = False
End Function

Private Function zk_ReadExtended(ByVal z As Object, ByVal machine As Long, ByVal maxReads As Long, _
                                 ByVal minKey As String, ByRef col As Collection) As Boolean
    On Error GoTo Fail
    
    Dim e As Long, v As Long, io As Long
    Dim y As Long, m As Long, d As Long, h As Long, n As Long, s As Long, w As Long, reserved As Long
    Dim more As Boolean, c As Long
    Dim r(0 To 3) As Variant
    
    more = z.GetGeneralExtLogData(machine, e, v, io, y, m, d, h, n, s, w, reserved)
    If Not more Then Exit Function
    
    Do While more
        If zk_Build(CStr(e), io, y, m, d, h, n, s, w, r) Then
            If zk_AddIfNewer(r, machine, minKey, col) Then c = c + 1
        End If
        
        If maxReads > 0 Then
            If c >= maxReads Then Exit Do
        End If
        
        more = z.GetGeneralExtLogData(machine, e, v, io, y, m, d, h, n, s, w, reserved)
    Loop
    
    zk_ReadExtended = (c > 0)
    Exit Function
Fail:
    zk_ReadExtended = False
End Function

Private Function zk_ReadString(ByVal z As Object, ByVal machine As Long, ByVal maxReads As Long, _
                               ByVal minKey As String, ByRef col As Collection) As Boolean
    On Error GoTo Fail
    
    Dim e As Long, v As Long, io As Long, dt As String
    Dim more As Boolean, c As Long, t As Date
    Dim r(0 To 3) As Variant
    
    more = z.GetGeneralLogDataStr(machine, e, v, io, dt)
    If Not more Then Exit Function
    
    Do While more
        If zk_ParseDate(dt, t) Then
            r(0) = CStr(e)
            r(1) = CStr(io)
            r(2) = t
            r(3) = 0
            If zk_AddIfNewer(r, machine, minKey, col) Then c = c + 1
        End If
        
        If maxReads > 0 Then
            If c >= maxReads Then Exit Do
        End If
        
        more = z.GetGeneralLogDataStr(machine, e, v, io, dt)
    Loop
    
    zk_ReadString = (c > 0)
    Exit Function
Fail:
    zk_ReadString = False
End Function

Private Function zk_ReadLegacy(ByVal z As Object, ByVal machine As Long, ByVal maxReads As Long, _
                               ByVal minKey As String, ByRef col As Collection) As Boolean
    On Error GoTo Fail
    
    Dim tm As Long, e As Long, em As Long, v As Long, io As Long
    Dim y As Long, m As Long, d As Long, h As Long, n As Long
    Dim more As Boolean, c As Long
    Dim r(0 To 3) As Variant
    
    more = z.GetGeneralLogData(machine, tm, e, em, v, io, y, m, d, h, n)
    If Not more Then Exit Function
    
    Do While more
        If zk_Build(CStr(e), io, y, m, d, h, n, 0, 0, r) Then
            If zk_AddIfNewer(r, machine, minKey, col) Then c = c + 1
        End If
        
        If maxReads > 0 Then
            If c >= maxReads Then Exit Do
        End If
        
        more = z.GetGeneralLogData(machine, tm, e, em, v, io, y, m, d, h, n)
    Loop
    
    zk_ReadLegacy = (c > 0)
    Exit Function
Fail:
    zk_ReadLegacy = False
End Function

' =========================================================
' توابع کمکی برای ساخت و تبدیل رکوردها
' =========================================================

Private Function zk_AddIfNewer(ByRef rec As Variant, ByVal machine As Long, ByVal minKey As String, _
                               ByRef col As Collection) As Boolean
    On Error GoTo Fail
    
    Dim candidateKey As String
    candidateKey = CStr(machine) & "|" & CStr(rec(0)) & "|" & _
                   Format$(rec(2), "yyyy-mm-dd HH:nn:ss") & "|" & _
                   CStr(rec(1)) & "|" & CStr(rec(3))
    
    If Len(Trim$(minKey)) = 0 Then
        col.Add rec
        zk_AddIfNewer = True
        Exit Function
    End If
    
    If cursor_IsAtOrAfterTimestamp(minKey, candidateKey) Then
        col.Add rec
        zk_AddIfNewer = True
    End If
    Exit Function
    
Fail:
    zk_AddIfNewer = False
End Function

Private Function zk_Build(ByVal enroll As String, ByVal io As Long, ByVal y As Long, _
                          ByVal m As Long, ByVal d As Long, ByVal h As Long, _
                          ByVal n As Long, ByVal s As Long, ByVal work As Long, _
                          ByRef r As Variant) As Boolean
    On Error GoTo Fail
    
    enroll = Trim$(enroll)
    If enroll = "" Then Exit Function
    If y < 1900 Or y > 2200 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    If h < 0 Or h > 23 Then Exit Function
    If n < 0 Or n > 59 Then Exit Function
    If s < 0 Or s > 59 Then Exit Function
    
    r(0) = enroll
    r(1) = CStr(io)
    r(2) = DateSerial(CInt(y), CInt(m), CInt(d)) + TimeSerial(CInt(h), CInt(n), CInt(s))
    r(3) = work
    zk_Build = True
    Exit Function
    
Fail:
    zk_Build = False
End Function

Private Function zk_ParseDate(ByVal value As String, ByRef outDate As Date) As Boolean
    On Error GoTo Fail
    
    Dim s As String
    Dim y As Long, m As Long, d As Long, h As Long, n As Long, sec As Long
    
    s = Trim$(value)
    If Len(s) < 10 Then Exit Function
    
    y = Val(Left$(s, 4))
    m = Val(Mid$(s, 6, 2))
    d = Val(Mid$(s, 9, 2))
    
    If Len(s) >= 19 Then
        h = Val(Mid$(s, 12, 2))
        n = Val(Mid$(s, 15, 2))
        sec = Val(Mid$(s, 18, 2))
    End If
    
    If y < 1900 Or y > 2200 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    If h < 0 Or h > 23 Then Exit Function
    If n < 0 Or n > 59 Then Exit Function
    If sec < 0 Or sec > 59 Then Exit Function
    
    outDate = DateSerial(CInt(y), CInt(m), CInt(d)) + TimeSerial(CInt(h), CInt(n), CInt(sec))
    zk_ParseDate = True
    Exit Function
    
Fail:
    zk_ParseDate = False
End Function

' =========================================================
' تابع: zk_MapAttendanceTypeVal
' =========================================================
' 
' وظیفه:
' شماره نوع تردد را به متن تبدیل می‌کند
'
' پارامتر:
'   value (Long): شماره نوع تردد
'       0 = IN, 1 = OUT, 2 = BREAK_OUT, 3 = BREAK_IN, etc
'
' خروجی: String
'   نام نوع تردد
'
' =========================================================

Public Function zk_MapAttendanceTypeVal(ByVal value As Long) As String
    Select Case value
        Case 0: zk_MapAttendanceTypeVal = "IN"
        Case 1: zk_MapAttendanceTypeVal = "OUT"
        Case 2: zk_MapAttendanceTypeVal = "BREAK_OUT"
        Case 3: zk_MapAttendanceTypeVal = "BREAK_IN"
        Case 4: zk_MapAttendanceTypeVal = "OT_IN"
        Case 5: zk_MapAttendanceTypeVal = "OT_OUT"
        Case Else: zk_MapAttendanceTypeVal = "UNKNOWN_" & CStr(value)
    End Select
End Function
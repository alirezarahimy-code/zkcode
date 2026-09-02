Option Compare Database
Option Explicit

' =========================================================
' ماژول: zk_realtime_service.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول رویدادهای Real-Time دستگاه ZK را مدیریت می‌کند.
' هنگامی که کاربر اثر انگشت می‌زند، دستگاه فوری رویداد آنی را فعال می‌کند.
'
' کاربرد:
' - دریافت فوری رویدادهای تردد
' - پردازش فوری بدون انتظار برای Polling
' - بسیار سریع‌تر از Polling
'
' ویژگی‌های مهم:
' - از COM Events zkemkeeper استفاده می‌کند
' - لیسنر (Listener) برای هر دستگاه
' - پردازش فوری تردد
' - معالجه خطاهای زمان اجرا
'
' معماری:
' - zkrt_RegisterDevice: ایجاد Listener برای دستگاه
' - zkrt_UnregisterDevice: حذف Listener
' - zkrt_HandleAttendance: پردازش تردد دریافت شده
' - clsZKRealtime: کلاس Listener
'
' =========================================================

' =========================================================
' تابع: zkrt_EnsureSessions
' =========================================================
' 
' وظیفه:
' اطمینان می‌دهد که Dictionary جلسات Realtime موجود است
' هر Realtime Listener در این Dictionary ذخیره می‌شود
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' =========================================================

Public Sub zkrt_EnsureSessions()
    On Error GoTo ErrHandler
    
    If ZKRealtimeSessions Is Nothing Then
        Set ZKRealtimeSessions = CreateObject("Scripting.Dictionary")
    End If
    
    Exit Sub
    
ErrHandler:
    Set ZKRealtimeSessions = Nothing
    Call LogError("zkrt_EnsureSessions", Err.Number, Err.Description, "")
End Sub

' =========================================================
' تابع: zkrt_RegisterDevice
' =========================================================
' 
' وظیفه:
' یک Listener (شنونده رویداد) برای دستگاه ایجاد می‌کند
' Listener به تمام رویدادهای تردد گوش می‌دهد
'
' پارامترها:
'   deviceKey (String): کلید یکتای دستگاه
'       فرمت: "192.168.1.100:4370:1"
'   z (Object): شیء ZK
'   deviceIP (String): آدرس IP
'   devicePort (Long): پورت
'   machineNumber (Long): شماره دستگاه
'   commKey (Long): کلید ارتباطی
'
' خروجی: Boolean
'   True: ثبت موفق
'   False: ثبت ناموفق
'
' نمونه استفاده:
'   If zkrt_RegisterDevice(key, zk, "192.168.1.100", 4370, 1, 0) Then
'       MsgBox "Listener ثبت شد"
'   End If
'
' نکات مهم:
' - این تابع در monitor_service فراخوانی می‌شود
' - فقط یک بار برای هر دستگاه فراخوانی شود
' - اگر دستگاه قطع شد و دوباره متصل شد، دوباره فراخوانی شود
'
' =========================================================

Public Function zkrt_RegisterDevice(ByVal deviceKey As String, ByVal z As Object, _
                                     ByVal deviceIP As String, ByVal devicePort As Long, _
                                     ByVal machineNumber As Long, ByVal commKey As Long) As Boolean
    On Error GoTo ErrHandler
    
    zkrt_EnsureSessions
    
    If ZKRealtimeSessions Is Nothing Then Exit Function
    If z Is Nothing Then Exit Function
    If Len(Trim$(deviceKey)) = 0 Then Exit Function
    
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If machineNumber <= 0 Then machineNumber = 1
    
    ' اگر این دستگاه قبلاً ثبت شده، بازگردان
    If ZKRealtimeSessions.Exists(deviceKey) Then
        Dim existing As clsZKRealtime
        Set existing = ZKRealtimeSessions(deviceKey)
        
        If Not existing Is Nothing Then
            zkrt_RegisterDevice = True
            Exit Function
        End If
        
        ZKRealtimeSessions.Remove deviceKey
    End If
    
    ' ایجاد Listener جدید
    Dim listener As clsZKRealtime
    Set listener = New clsZKRealtime
    
    If Not listener.Attach(z, deviceKey, deviceIP, devicePort, machineNumber, commKey) Then
        Exit Function
    End If
    
    ' ثبت در Dictionary
    ZKRealtimeSessions.Add deviceKey, listener
    
    Call LogSystemEvent("zkrt_RegisterDevice", "Listener ثبت شد: " & deviceKey)
    
    zkrt_RegisterDevice = True
    Exit Function
    
ErrHandler:
    Call LogError("zkrt_RegisterDevice", Err.Number, Err.Description, deviceKey)
    zkrt_RegisterDevice = False
End Function

' =========================================================
' تابع: zkrt_UnregisterDevice
' =========================================================
' 
' وظیفه:
' Listener یک دستگاه را حذف می‌کند
' رویدادهای دستگاه دیگر دریافت نمی‌شوند
'
' پارامتر:
'   deviceKey (String): کلید دستگاه
'
' خروجی: ندارد
'
' نمونه استفاده:
'   Call zkrt_UnregisterDevice(key)
'
' =========================================================

Public Sub zkrt_UnregisterDevice(ByVal deviceKey As String)
    On Error Resume Next
    
    zkrt_EnsureSessions
    
    If ZKRealtimeSessions Is Nothing Then Exit Sub
    If Not ZKRealtimeSessions.Exists(deviceKey) Then Exit Sub
    
    Dim listener As clsZKRealtime
    Set listener = ZKRealtimeSessions(deviceKey)
    
    If Not listener Is Nothing Then
        listener.Detach
    End If
    
    ZKRealtimeSessions.Remove deviceKey
    
    Call LogSystemEvent("zkrt_UnregisterDevice", "Listener حذف شد: " & deviceKey)
End Sub

' =========================================================
' تابع: zkrt_StopAll
' =========================================================
' 
' وظیفه:
' تمام Listener‌ها را حذف می‌کند
' معمولاً هنگام بسته شدن برنامه فراخوانی می‌شود
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' =========================================================

Public Sub zkrt_StopAll()
    On Error Resume Next
    
    zkrt_EnsureSessions
    
    If ZKRealtimeSessions Is Nothing Then Exit Sub
    
    Dim key As Variant
    Dim listener As clsZKRealtime
    
    For Each key In ZKRealtimeSessions.Keys
        Set listener = Nothing
        Set listener = ZKRealtimeSessions(key)
        If Not listener Is Nothing Then listener.Detach
    Next key
    
    ZKRealtimeSessions.RemoveAll
    
    Call LogSystemEvent("zkrt_StopAll", "تمام Listener‌ها حذف شدند")
End Sub

' =========================================================
' تابع: zkrt_IsRegistered
' =========================================================
' 
' وظیفه:
' بررسی می‌کند که آیا Listener برای دستگاه ثبت شده یا نه
'
' پارامتر:
'   deviceKey (String): کلید دستگاه
'
' خروجی: Boolean
'   True: ثبت شده است
'   False: ثبت نشده است
'
' =========================================================

Public Function zkrt_IsRegistered(ByVal deviceKey As String) As Boolean
    On Error Resume Next
    
    zkrt_EnsureSessions
    
    If ZKRealtimeSessions Is Nothing Then Exit Function
    zkrt_IsRegistered = ZKRealtimeSessions.Exists(deviceKey)
End Function

' =========================================================
' تابع: zkrt_HandleAttendance
' =========================================================
' 
' وظیفه:
' رویداد تردد دریافت شده را پردازش می‌کند
' این تابع توسط clsZKRealtime هنگام دریافت تردد فراخوانی می‌شود
'
' پارامترها:
'   deviceKey (String): کلید دستگاه
'   deviceIP (String): IP دستگاه
'   devicePort (Long): پورت دستگاه
'   machineNumber (Long): شماره دستگاه
'   enroll (String): کد ثبتی کاربر
'   isInvalid (Long): آیا تردد معتبر است (0=معتبر)
'   attState (Long): نوع تردد (0=IN, 1=OUT, etc)
'   verifyMethod (Long): روش تأیید (0=Fingerprint, etc)
'   y, m, d, h, mn, sc (Long): تاریخ و زمان تردد
'   workCode (Long): کد کار
'
' خروجی: ندارد
'
' فرآیند:
'   1. اگر تردد نامعتبر است (isInvalid <> 0): بازگردان
'   2. اعتبار‌سنجی تاریخ/زمان
'   3. ایجاد رکورد تردد جدید
'   4. فراخوانی proc_ProcessRecordAtomic برای پردازش فوری
'
' =========================================================

Public Sub zkrt_HandleAttendance(ByVal deviceKey As String, ByVal deviceIP As String, _
                                  ByVal devicePort As Long, ByVal machineNumber As Long, _
                                  ByVal enroll As String, ByVal isInvalid As Long, _
                                  ByVal attState As Long, ByVal verifyMethod As Long, _
                                  ByVal y As Long, ByVal mo As Long, ByVal da As Long, _
                                  ByVal hr As Long, ByVal mn As Long, ByVal sc As Long, _
                                  ByVal workCode As Long)
    On Error GoTo ErrHandler
    
    ' اگر تردد نامعتبر است، نپذیر
    If isInvalid <> 0 Then Exit Sub
    
    enroll = Trim$(enroll)
    If enroll = "" Then Exit Sub
    
    ' اعتبار‌سنجی تاریخ/زمان
    If y < 1900 Or y > 2200 Then Exit Sub
    If mo < 1 Or mo > 12 Then Exit Sub
    If da < 1 Or da > 31 Then Exit Sub
    If hr < 0 Or hr > 23 Then Exit Sub
    If mn < 0 Or mn > 59 Then Exit Sub
    If sc < 0 Or sc > 59 Then Exit Sub
    
    ' ایجاد تاریخ/زمان
    Dim attDT As Date
    attDT = DateSerial(CInt(y), CInt(mo), CInt(da)) + _
            TimeSerial(CInt(hr), CInt(mn), CInt(sc))
    
    ' ایجاد کلید رکورد
    Dim rawKey As String
    rawKey = CStr(machineNumber) & "|" & enroll & "|" & _
             Format$(attDT, "yyyy-mm-dd HH:nn:ss") & "|" & _
             CStr(attState) & "|" & CStr(workCode)
    
    ' درج رکورد تردد
    Dim recID As Long
    recID = att_InsertRaw(enroll, deviceIP, devicePort, machineNumber, attDT, _
                          zk_MapAttendanceTypeVal(attState), rawKey, "NEW")
    
    If recID <= 0 Then
        Call LogSystemEvent("zkrt_HandleAttendance", _
                           "درج ناموفق: " & deviceKey & " RawKey=" & rawKey)
        Exit Sub
    End If
    
    ' فراخوانی پردازش فوری (Atomic Processing)
    ' اگر موفق در Claim شود (یعنی هنوز کس دیگری در حال پردازش نیست)
    If db_TryClaimRecordAtomic(recID, "PROCESSING", _
                              "'NEW','PRINT_FAILED','NO_EMPLOYEE','WAITING_FOR_MEAL','RECEIPT_FAILED','MEAL_FINALIZE_FAILED'") Then
        Call proc_ProcessRecordAtomic(recID)
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("zkrt_HandleAttendance", Err.Number, Err.Description, _
                  "Device=" & deviceKey & " | Enroll=" & enroll)
End Sub
Option Compare Database
Option Explicit

' =========================================================
' ماژول: monitor_service.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول Polling (پایش دوره‌ای) دستگاه‌های ZK را مدیریت می‌کند.
' هر چند ثانیه یک بار، تمام دستگاه‌های فعال را بررسی می‌کند.
'
' کاربرد:
' - اتصال به تمام دستگاه‌های ZK
' - خواندن لاگ‌های تردد جدید
' - Fallback برای رویدادهای Real-Time
' - بازیابی تردد‌های از دست رفته
'
' ویژگی‌های مهم:
' - Thread-safe (تک‌ریسمان)
' - Cursor-based (خواندن فقط تردد‌های جدید)
' - Atomic Processing (هر تردد یک بار پردازش شود)
' - Recovery (بازیابی چاپ‌های گیر‌کرده)
'
' معماری:
' - monitor_Start: شروع مانیتور
' - monitor_Stop: توقف مانیتور
' - monitor_Tick: هر دوره اجرا شود (از ماژول UI)
' - monitor_ProcessDeviceLogs: پردازش لاگ‌های دستگاه
'
' =========================================================

' =========================================================
' تابع: monitor_Start
' =========================================================
' 
' وظیفه:
' مانیتور را شروع می‌کند
' متغیرهای سراسری را مقداردهی می‌کند
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' نمونه استفاده:
'   Call monitor_Start()
'
' نکات مهم:
' - باید یک بار در شروع برنامه فراخوانی شود
' - بعد از اینکه DB تماماً آماده شد
'
' =========================================================

Public Sub monitor_Start()
    On Error GoTo ErrHandler
    
    EnsureDeviceSessions
    If DeviceSessions Is Nothing Then
        Err.Raise vbObjectError + 2100, "monitor_Start", "متغیر جلسات آماده نشد"
    End If
    
    MonitorRunning = True
    MonitorBusy = False
    
    Call LogSystemEvent("monitor_Start", "مانیتور شروع شد")
    Exit Sub
    
ErrHandler:
    MonitorRunning = False
    MonitorBusy = False
    Call LogError("monitor_Start", Err.Number, Err.Description, "")
End Sub

' =========================================================
' تابع: monitor_Stop
' =========================================================
' 
' وظیفه:
' مانیتور را متوقف می‌کند
' تمام اتصالات را قطع می‌کند
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' نمونه استفاده:
'   Call monitor_Stop()
'
' فرآیند:
'   1. متوقف کردن مانیتور
'   2. حذف تمام Realtime Listeners
'   3. قطع اتصال تمام دستگاه‌ها
'   4. پاک کردن جلسات
'
' =========================================================

Public Sub monitor_Stop()
    On Error Resume Next
    
    MonitorRunning = False
    MonitorBusy = False
    
    EnsureDeviceSessions
    
    ' قطع رویدادهای Realtime
    Call zkrt_StopAll()
    
    ' قطع اتصالات
    Dim k As Variant, session As Object, z As Object
    For Each k In DeviceSessions.Keys
        Set session = Nothing
        Set z = Nothing
        Set session = DeviceSessions(k)
        If Not session Is Nothing Then
            If session.Exists("ZKObj") Then Set z = session("ZKObj")
            If Not z Is Nothing Then z.Disconnect
        End If
    Next k
    
    DeviceSessions.RemoveAll
    
    Call LogSystemEvent("monitor_Stop", "مانیتور متوقف شد")
End Sub

' =========================================================
' تابع: monitor_Tick
' =========================================================
' 
' وظیفه:
' یک دوره پایش را اجرا می‌کند
' باید هر 5 ثانیه یک بار از Timer فراخوانی شود
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' فرآیند:
'   1. اگر مانیتور متوقف است: بازگردان
'   2. اگر مانیتور مشغول است: بازگردان (از تداخل جلوگیری)
'   3. بازیابی چاپ‌های گیر‌کرده
'   4. برای هر دستگاه:
'      a. اتصال برقرار کن
'      b. لاگ‌ها را بخوان
'      c. لاگ‌ها را پردازش کن
'      d. Realtime Listener ثبت کن
'   5. پردازش تردد‌های باقی‌مانده
'
' نمونه استفاده:
'   ' از Timer فراخوانی شود
'   Call monitor_Tick()
'
' نکات مهم:
' - این تابع باید از Timer فراخوانی شود
' - اگر قبلی هنوز در حال اجرا است، بازگردان
' - خطاهای احتمالی معالجه می‌شود
'
' =========================================================

Public Sub monitor_Tick()
    On Error GoTo ErrHandler
    
    If Not MonitorRunning Then Exit Sub
    If MonitorBusy Then Exit Sub
    
    MonitorBusy = True
    
    ' بازیابی چاپ‌های گیر‌کرده
    Call proc_RecoverStuckPrinting()
    
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb()
    
    ' خواندن تمام دستگاه‌های فعال
    Set rs = db.OpenRecordset( _
        "SELECT DeviceID,DeviceIP,DevicePort,MachineNumber,CommKey FROM " & TABLE_ZK_DEVICES & " " & _
        "WHERE Nz(IsActive,True)=True ORDER BY DeviceID", dbOpenSnapshot)
    
    Do While Not rs.EOF
        Dim ip As String, deviceKey As String, prevKey As String
        Dim port As Long, machine As Long, commKey As Long
        Dim z As Object, logs As Collection, state As Variant
        Dim initialSync As Boolean
        
        ip = Trim$(Nz(rs!DeviceIP, ""))
        port = Nz(rs!DevicePort, DEFAULT_ZK_PORT)
        machine = Nz(rs!MachineNumber, 1)
        commKey = Nz(rs!CommKey, 0)
        
        If port <= 0 Then port = DEFAULT_ZK_PORT
        If machine <= 0 Then machine = 1
        
        If ip <> "" Then
            deviceKey = MakeDeviceKey(ip, port, machine)
            
            ' به‌روزرسانی تلاش اتصال
            Call monitor_UpdateConnectionAttempt(db, ip, port, machine)
            
            ' دریافت وضعیت دستگاه
            state = cursor_GetState(deviceKey)
            initialSync = True
            prevKey = ""
            
            ' بررسی اگر دستگاه قبلاً همگام‌سازی شده
            If Not IsNull(state) Then
                If IsArray(state) Then
                    If UBound(state) >= 2 Then
                        initialSync = IsNull(state(2))
                    End If
                    
                    If Not initialSync Then
                        If UBound(state) >= 1 Then prevKey = Nz(state(1), "")
                        
                        If Len(prevKey) = 0 And UBound(state) >= 0 Then
                            If Not IsNull(state(0)) Then
                                prevKey = CStr(machine) & "|__CURSOR__|" & _
                                          Format$(state(0), "yyyy-mm-dd HH:nn:ss") & "|0|0"
                            End If
                        End If
                    End If
                End If
            End If
            
            ' اتصال و خواندن لاگ
            Set z = device_EnsureSession(ip, port, machine, commKey)
            If Not z Is Nothing Then
                ' ثبت Realtime Listener
                Call zkrt_RegisterDevice(deviceKey, z, ip, port, machine, commKey)
                
                ' خواندن لاگ‌ها
                If initialSync Then
                    ' همگام‌سازی اولیه: فقط آخرین رکورد
                    Set logs = zk_ReadLogs(z, machine, INITIAL_SYNC_MAX_READS, "")
                Else
                    ' خواندن تردد‌های جدید
                    Set logs = zk_ReadLogs(z, machine, READ_MAX_PER_CALL, prevKey)
                End If
                
                ' پردازش لاگ‌ها
                If Not logs Is Nothing Then
                    Call monitor_ProcessDeviceLogs(ip, port, machine, deviceKey, logs, initialSync)
                End If
            End If
        Else
            Call LogError("monitor_Tick", -1, "دستگاه بدون IP", "DeviceID=" & CStr(Nz(rs!DeviceID, 0)))
        End If
        
        rs.MoveNext
    Loop
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' پردازش تردد‌های باقی‌مانده
    Call proc_ProcessPendingBatch()
    
    MonitorBusy = False
    Exit Sub
    
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    MonitorBusy = False
    
    Call LogError("monitor_Tick", en, ed, "")
End Sub

' =========================================================
' تابع: monitor_UpdateConnectionAttempt
' =========================================================
' 
' وظیفه:
' به‌روزرسانی آخرین تلاش اتصال برای دستگاه
'
' =========================================================

Private Sub monitor_UpdateConnectionAttempt(ByVal db As DAO.Database, ByVal deviceIP As String, _
                                            ByVal devicePort As Long, ByVal machineNumber As Long)
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset( _
        "SELECT TOP 1 LastConnectionAttempt FROM " & TABLE_ZK_DEVICES & " " & _
        "WHERE DeviceIP='" & Replace(deviceIP, "'", "''") & "' " & _
        "AND DevicePort=" & CStr(devicePort) & " " & _
        "AND MachineNumber=" & CStr(machineNumber), dbOpenDynaset)
    
    If Not rs.EOF Then
        rs.Edit
        rs!LastConnectionAttempt = Now()
        rs.Update
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    Call LogError("monitor_UpdateConnectionAttempt", Err.Number, Err.Description, _
                  deviceIP & ":" & CStr(devicePort) & ":" & CStr(machineNumber))
    Resume CleanExit
End Sub

' =========================================================
' تابع: monitor_ProcessDeviceLogs
' =========================================================
' 
' وظیفه:
' لاگ‌های دریافت شده از دستگاه را پردازش می‌کند
'
' پارامترها:
'   ip (String): IP دستگاه
'   port (Long): پورت
'   machine (Long): شماره دستگاه
'   deviceKey (String): کلید دستگاه
'   logs (Collection): مجموعه لاگ‌ها
'   initialSync (Boolean): اولین همگام‌سازی؟
'
' فرآیند:
'   1. برای هر لاگ:
'      a. بررسی معتبر بودن
'      b. اگر اولین همگام: فقط درج (بدون چاپ)
'      c. اگر بعد از اولین: درج و پردازش فوری
'   2. بروزرسانی Cursor (آخرین موقعیت خوانده‌شده)
'   3. ذخیره وضعیت دستگاه
'
' =========================================================

Private Sub monitor_ProcessDeviceLogs(ByVal ip As String, ByVal port As Long, ByVal machine As Long, _
                                       ByVal deviceKey As String, ByVal logs As Collection, _
                                       Optional ByVal initialSync As Boolean = False)
    On Error GoTo ErrHandler
    
    Dim state As Variant, prevKey As String
    Dim maxKey As String, maxDT As Date
    Dim syncOK As Boolean
    Dim i As Long, rec As Variant
    Dim enroll As String, candidateKey As String
    Dim io As Long, work As Long, dt As Date, id As Long
    
    prevKey = ""
    If Not initialSync Then
        state = cursor_GetState(deviceKey)
        If Not IsNull(state) Then
            If IsArray(state) Then
                If UBound(state) >= 1 Then prevKey = Nz(state(1), "")
            End If
        End If
    End If
    
    maxKey = prevKey
    maxDT = DateSerial(1970, 1, 1)
    syncOK = True
    
    ' پردازش هر لاگ
    For i = 1 To logs.Count
        rec = logs(i)
        
        If monitor_IsValidLogRecord(rec) Then
            enroll = Trim$(CStr(rec(0)))
            io = CLng(rec(1))
            dt = rec(2)
            work = CLng(rec(3))
            
            candidateKey = CStr(machine) & "|" & enroll & "|" & Format$(dt, "yyyy-mm-dd HH:nn:ss") & _
                          "|" & CStr(io) & "|" & CStr(work)
            
            If Len(maxKey) = 0 Or cursor_IsNewerKey(maxKey, candidateKey) Then
                maxKey = candidateKey
                maxDT = dt
            End If
            
            If initialSync Then
                ' اولین همگام: فقط درج (بدون چاپ)
                id = att_InsertRaw(enroll, ip, port, machine, dt, zk_MapAttendanceTypeVal(io), _
                                  candidateKey, "HISTORICAL")
                If id <= 0 Then syncOK = False
            ElseIf Len(prevKey) = 0 Or cursor_IsAtOrAfterTimestamp(prevKey, candidateKey) Then
                ' خواندن جدید: درج و پردازش فوری
                id = att_InsertRaw(enroll, ip, port, machine, dt, zk_MapAttendanceTypeVal(io), _
                                  candidateKey, "NEW")
                
                If id > 0 Then
                    ' تلاش برای Atomic Claim و پردازش
                    If db_TryClaimRecordAtomic(id, "PROCESSING", _
                                              "'NEW','PRINT_FAILED','NO_EMPLOYEE','WAITING_FOR_MEAL','RECEIPT_FAILED','MEAL_FINALIZE_FAILED'") Then
                        Call proc_ProcessRecordAtomic(id)
                    End If
                Else
                    syncOK = False
                    Call LogSystemEvent("monitor_ProcessDeviceLogs", _
                                       "درج ناموفق: Device=" & deviceKey & " Key=" & candidateKey)
                End If
            End If
        Else
            Call LogSystemEvent("monitor_ProcessDeviceLogs", _
                               "رکورد نامعتبر: Device=" & deviceKey & " Index=" & CStr(i))
        End If
    Next i
    
    ' بروزرسانی Cursor
    If initialSync Then
        If syncOK Then
            If Len(maxKey) > 0 Then
                Call cursor_UpdateState(deviceKey, ip, port, machine, maxDT, maxKey, "")
            Else
                Call cursor_UpdateState(deviceKey, ip, port, machine, Now(), "", "")
            End If
        Else
            Call cursor_UpdateState(deviceKey, ip, port, machine, Now(), "", _
                                   "همگام‌سازی اولیه ناموفق")
        End If
    ElseIf syncOK And Len(maxKey) > 0 Then
        If Len(prevKey) = 0 Or cursor_IsNewerKey(prevKey, maxKey) Then
            Call cursor_UpdateState(deviceKey, ip, port, machine, maxDT, maxKey, "")
        End If
    ElseIf Not syncOK Then
        Call cursor_UpdateState(deviceKey, ip, port, machine, maxDT, maxKey, _
                               "همگام‌سازی ناموفق")
    End If
    
    Exit Sub
    
ErrHandler:
    Call LogError("monitor_ProcessDeviceLogs", Err.Number, Err.Description, _
                  "Device=" & deviceKey)
End Sub

' =========================================================
' تابع: monitor_IsValidLogRecord
' =========================================================
' 
' وظیفه:
' بررسی می‌کند که آیا لاگ معتبر است یا نه
'
' =========================================================

Private Function monitor_IsValidLogRecord(ByVal rec As Variant) As Boolean
    On Error GoTo EH
    
    If Not IsArray(rec) Then Exit Function
    If UBound(rec) < 3 Then Exit Function
    If Trim$(CStr(Nz(rec(0), ""))) = "" Then Exit Function
    If Not IsNumeric(rec(1)) Then Exit Function
    If Not IsDate(rec(2)) Then Exit Function
    If Not IsNumeric(rec(3)) Then Exit Function
    
    monitor_IsValidLogRecord = True
    Exit Function
    
EH:
    monitor_IsValidLogRecord = False
End Function

' =========================================================
' تابع کمکی: IsDate
' =========================================================

Private Function IsDate(ByVal value As Variant) As Boolean
    On Error GoTo EH
    IsDate = Not IsNull(CDate(value))
    Exit Function
EH:
    IsDate = False
End Function
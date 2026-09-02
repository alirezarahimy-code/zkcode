Option Compare Database
Option Explicit

' =========================================================
' ماژول: logger.bas
' =========================================================
' 
' توضیح ماژول:
' سیستم ثبت لاگ (Logging) برنامه را مدیریت می‌کند.
' تمام رویدادهای مهم، خطاها، و اتفاقات در جداول ثبت می‌شوند.
' این لاگ‌ها برای دیباگ و عیب‌یابی بسیار مهم هستند.
'
' کاربرد:
' - ثبت رویدادهای سیستم (System Events)
' - ثبت خطاهای برنامه (Error Logging)
' - ثبت رویدادهای تردد (Attendance Events)
' - پاد‌و‌پیش‌رو (Fallback) در TEMP اگر DB ناپایدار بود
'
' ویژگی‌ها:
' - اگر جدول DB نتوانست، به فایل در TEMP می‌نویسد
' - تمام لاگ‌ها با Timestamp ثبت می‌شوند
' - خطاها به صورت کامل ثبت می‌شوند (شماره، توضیح، منبع)
' - اطلاعات اضافی (Additional Info) می‌تواند ذخیره شود
'
' =========================================================

' =========================================================
' تابع: LogSystemEvent
' =========================================================
' 
' وظیفه:
' هر رویداد سیستم‌ای را در جدول tblSystemLogs ثبت می‌کند
' استفاده می‌شود برای پیگیری رویدادهای نرمال و مهم
'
' پارامترها:
'   source (String): منبع رویداد
'       مثال: "monitor_Tick", "device_EnsureSession", "printer_PrintReceipt"
'   msg (String): توضیح رویداد
'       مثال: "شروع پایش دستگاه", "متصل شد: 192.168.1.100:4370"
'
' خروجی: ندارد
'
' خطاهای ممکن:
' - اگر جدول وجود نداشت، فایل TEMP می‌شود (Fallback)
' - اگر TEMP نوشتنی نبود، ساکت (بدون خطا)
'
' نمونه استفاده:
'   Call LogSystemEvent("monitor_Start", "مانیتور شروع شد")
'   Call LogSystemEvent("device_Connect", "اتصال موفق: " & ip)
'   Call LogSystemEvent("printer_PrintReceipt", "فیش چاپ شد")
'
' بهترین عملکرد:
' - هر عملیات مهم باید لاگ شود
' - متن توضیحی باید واضح و اختصاری باشد
' - اگر خطا داشت: LogError را استفاده کنید، نه LogSystemEvent
'
' =========================================================

Public Sub LogSystemEvent(ByVal source As String, ByVal msg As String)
    On Error Resume Next
    
    ' منبع و پیام را تمیز کن
    source = Trim$(Left$(source, 100))
    msg = Trim$(msg)
    
    ' اگر خالی است، لاگ نکن
    If Len(source) = 0 Then Exit Sub
    
    ' سعی کن در DB بنویس
    If TryWriteDbSystemLog(source, msg) Then Exit Sub
    
    ' اگر DB ناموفق بود، به TEMP بنویس
    AppendFallback _
        "ZK_SystemLog.txt", _
        Format$(Now(), "yyyy-mm-dd HH:nn:ss") & " | " & source & " | " & msg
End Sub

' =========================================================
' تابع: LogError
' =========================================================
' 
' وظیفه:
' هر خطای برنامه را در جدول tblErrorLogs ثبت می‌کند
' این لاگ‌ها برای عیب‌یابی و درک علت خطا استفاده می‌شوند
'
' پارامترها:
'   procName (String): نام تابع یا ماژول که خطا رخ داده
'       مثال: "emp_ResolveByDeviceEnroll", "db_InsertAttendanceRaw"
'   errNumber (Long): شماره خطا (Err.Number)
'       مثال: 11 (Division by Zero), 13 (Type Mismatch)
'   errDesc (String): توضیح خطا (Err.Description)
'       مثال: "Type Mismatch", "Object variable not set"
'   additional (String): اطلاعات اضافی (اختیاری)
'       مثال: "DeviceIP=192.168.1.100 | Port=4370"
'
' خروجی: ندارد
'
' خطاهای ممکن:
' - اگر جدول وجود نداشت، فایل TEMP می‌شود (Fallback)
'
' نمونه استفاده:
'   On Error GoTo EH
'   Dim x As Integer
'   x = 1 / 0  ' خطای تقسیم بر صفر
'   Exit Sub
'   EH:
'       Call LogError("MyFunction", Err.Number, Err.Description, "x=" & x)
'
' نمونه واقعی:
'   On Error GoTo EH
'   Dim rs As DAO.Recordset
'   Set rs = db.OpenRecordset("SELECT * FROM tblNonExistent")
'   Exit Sub
'   EH:
'       LogError "device_EnsureSession", Err.Number, Err.Description, _
'                "DeviceIP=" & ip & " Port=" & port
'
' نکات مهم:
' - همیشه Err.Number و Err.Description را بدون تغییر ثبت کنید
' - Additional Info باید دقیق و مفید باشد
' - هرگز LogError را بدون On Error GoTo قبل از آن فراخوانی نکن
'
' =========================================================

Public Sub LogError(ByVal procName As String, ByVal errNumber As Long, ByVal errDesc As String, Optional ByVal additional As String = "")
    On Error Resume Next
    
    ' پارامترها را تمیز کن
    procName = Trim$(Left$(procName, 255))
    errDesc = Trim$(errDesc)
    additional = Trim$(additional)
    
    ' اگر خالی است، لاگ نکن
    If Len(procName) = 0 Then Exit Sub
    
    ' سعی کن در DB بنویس
    If TryWriteDbErrorLog(procName, errNumber, errDesc, additional) Then Exit Sub
    
    ' اگر DB ناموفق بود، به TEMP بنویس
    AppendFallback _
        "ZK_ErrorLog.txt", _
        Format$(Now(), "yyyy-mm-dd HH:nn:ss") & " | " & procName & _
        " | Err:" & CStr(errNumber) & " | " & errDesc & _
        " | " & additional
End Sub

' =========================================================
' تابع: LogAttendanceEvent
' =========================================================
' 
' وظیفه:
' رویدادهای تردد را برای پایش آنی (Real-Time Monitoring) ثبت می‌کند
' این رویدادها درفرم پایش نمایش داده می‌شوند
'
' پارامترها:
'   enrollID (String): کد ثبتی کاربر در دستگاه
'       مثال: "1234"
'   result (String): نتیجه پردازش تردد
'       مثال: "PRINTED", "NO_MEAL_ORDER", "NO_EMPLOYEE"
'   details (String): جزئیات بیشتر
'       مثال: "علی احمدی - فیش چاپ شد"
'
' خروجی: ندارد
'
' کاربرد:
' - فرم پایش آنی این رویدادها را نمایش می‌دهد
' - کاربر می‌تواند رویدادهای اخیر را ببیند
'
' نمونه استفاده:
'   Call LogAttendanceEvent("1234", "PRINTED", "علی احمدی - فیش چاپ شد")
'   Call LogAttendanceEvent("5678", "NO_MEAL_ORDER", "غذا سفارش نداده")
'   Call LogAttendanceEvent("9012", "NO_EMPLOYEE", "کارمند یافت نشد")
'
' =========================================================

Public Sub LogAttendanceEvent(ByVal enrollID As String, ByVal result As String, ByVal details As String)
    On Error Resume Next
    
    enrollID = Trim$(enrollID)
    result = Trim$(result)
    details = Trim$(details)
    
    If Len(enrollID) = 0 Or Len(result) = 0 Then Exit Sub
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    
    On Error Resume Next
    If TableExists(TABLE_LIVE_MONITORING) Then
        Set rs = db.OpenRecordset(TABLE_LIVE_MONITORING, dbOpenDynaset)
        If Not rs Is Nothing Then
            rs.AddNew
            rs!EnrollID = enrollID
            rs!Result = result
            rs!Details = details
            rs!LogDateTime = Now()
            rs.Update
        End If
    End If
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' =========================================================
' تابع کمکی: TryWriteDbSystemLog
' =========================================================
' 
' وظیفه:
' سعی می‌کند رویداد را در جدول DB ثبت کند
' اگر موفق بود: True برمی‌گرداند
' اگر ناموفق بود: False برمی‌گرداند (Fallback سعی خواهد کرد)
'
' =========================================================

Private Function TryWriteDbSystemLog(ByVal source As String, ByVal msg As String) As Boolean
    On Error GoTo EH
    
    If Not TableExists(TABLE_SYSTEM_LOGS) Then Exit Function
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset(TABLE_SYSTEM_LOGS, dbOpenDynaset)
    
    If rs Is Nothing Then Exit Function
    
    rs.AddNew
    rs!LogDate = Now()
    rs!Source = source
    rs!Message = msg
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    TryWriteDbSystemLog = True
    Exit Function
    
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    TryWriteDbSystemLog = False
End Function

' =========================================================
' تابع کمکی: TryWriteDbErrorLog
' =========================================================
' 
' وظیفه:
' سعی می‌کند خطا را در جدول DB ثبت کند
' اگر موفق بود: True برمی‌گرداند
' اگر ناموفق بود: False برمی‌گرداند (Fallback سعی خواهد کرد)
'
' =========================================================

Private Function TryWriteDbErrorLog(ByVal procName As String, ByVal errNumber As Long, ByVal errDesc As String, Optional ByVal additional As String = "") As Boolean
    On Error GoTo EH
    
    If Not TableExists(TABLE_ERROR_LOGS) Then Exit Function
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset(TABLE_ERROR_LOGS, dbOpenDynaset)
    
    If rs Is Nothing Then Exit Function
    
    rs.AddNew
    rs!ErrorDate = Now()
    rs!ErrNumber = errNumber
    rs!ErrDescription = errDesc
    rs!Procedure = procName
    rs!AdditionalInfo = additional
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    TryWriteDbErrorLog = True
    Exit Function
    
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    TryWriteDbErrorLog = False
End Function

' =========================================================
' تابع کمکی: AppendFallback
' =========================================================
' 
' وظیفه:
' اگر DB ثبت نتوانست، به فایل TextLog در TEMP بنویسد
' این تابع آخرین چاره (Last Resort) است
'
' =========================================================

Private Sub AppendFallback(ByVal fileName As String, ByVal lineText As String)
    On Error Resume Next
    
    Dim filePath As String
    Dim f As Integer
    
    filePath = Environ$("TEMP") & "\" & fileName
    f = FreeFile
    
    Open filePath For Append As #f
    Print #f, lineText
    Close #f
End Sub

' =========================================================
' تابع کمکی: TableExists
' =========================================================
' 
' وظیفه:
' بررسی می‌کند آیا جدول موجود است یا نه
'
' =========================================================

Private Function TableExists(ByVal TableName As String) As Boolean
    On Error GoTo EH
    
    Dim td As DAO.TableDef
    
    For Each td In CurrentDb().TableDefs
        If StrComp(td.Name, TableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next td
    
    Exit Function
    
EH:
    TableExists = False
End Function

' =========================================================
' ثابت برای جدول پایش آنی
' =========================================================

Private Const TABLE_LIVE_MONITORING As String = "tblLiveMonitoring"
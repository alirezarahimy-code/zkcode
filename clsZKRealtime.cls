Option Compare Database
Option Explicit

' =========================================================
' کلاس: clsZKRealtime
' =========================================================
'
' توضیح کلاس:
' این کلاس یک Listener برای دستگاه ZK است.
' هنگام اثرانگشت زدن کاربر، فوری رویداد را دریافت می‌کند.
'
' کاربرد:
' - دریافت رویدادهای Real-Time از دستگاه ZK
' - پردازش فوری بدون انتظار برای Polling
' - بسیار سریع‌تر از Polling (< 1 ثانیه)
'
' ویژگی‌های مهم:
' - استفاده از COM Events (zkemkeeper.dll)
' - یک Listener برای هر دستگاه
' - معالجه تمام خطاهای رابط‌گاه
' - Log کامل برای تشخیص عیب
'
' رویدادهای تحت نظارت:
' - OnAttendance: رویداد تردد
' - OnEnrollStatus: تغییر شماره ثبتی
' - OnNewUser: کاربر جدید
' - OnDeleteUser: حذف کاربر
'
' معماری:
' - Attach: اتصال به دستگاه
' - Detach: قطع اتصال
' - زوار رویدادهای COM
'
' =========================================================

Option Explicit

Private WithEvents zkDevice As zkemkeeper.CZKEM

Private deviceKey As String
Private deviceIP As String
Private devicePort As Long
Private machineNumber As Long
Private commKey As Long
Private isAttached As Boolean

' =========================================================
' تابع: Attach
' =========================================================
'
' وظیفه:
' Listener را به دستگاه ZK متصل می‌کند
' رویدادهای COM را فعال می‌کند
'
' پارامترها:
'   zk (Object): شیء zkemkeeper.CZKEM متصل‌شده
'   devKey (String): کلید دستگاه (IP:Port:Machine)
'   ip (String): آدرس IP
'   port (Long): پورت
'   machine (Long): شماره دستگاه
'   cKey (Long): کلید ارتباطی
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' نمونه استفاده:
'   Dim listener As clsZKRealtime
'   Set listener = New clsZKRealtime
'   If listener.Attach(zkObj, "192.168.1.100:4370:1", "192.168.1.100", 4370, 1, 0) Then
'       ' متصل شد
'   End If
'
' =========================================================

Public Function Attach(ByVal zk As Object, ByVal devKey As String, ByVal ip As String, _
                       ByVal port As Long, ByVal machine As Long, ByVal cKey As Long) As Boolean
    On Error GoTo ErrHandler
    
    If zk Is Nothing Then Exit Function
    If Len(Trim$(devKey)) = 0 Then Exit Function
    
    Set zkDevice = zk
    deviceKey = devKey
    deviceIP = ip
    devicePort = port
    machineNumber = machine
    commKey = cKey
    isAttached = True
    
    Call LogSystemEvent("clsZKRealtime.Attach", "Listener متصل شد: " & deviceKey)
    
    Attach = True
    Exit Function
    
ErrHandler:
    Call LogError("clsZKRealtime.Attach", Err.Number, Err.Description, devKey)
    Attach = False
End Function

' =========================================================
' تابع: Detach
' =========================================================
'
' وظیفه:
' Listener را از دستگاه قطع می‌کند
' رویدادهای COM را غیرفعال می‌کند
'
' پارامتر: ندارد
'
' خروجی: ندارد
'
' =========================================================

Public Sub Detach()
    On Error Resume Next
    
    isAttached = False
    Set zkDevice = Nothing
    
    Call LogSystemEvent("clsZKRealtime.Detach", "Listener قطع شد: " & deviceKey)
End Sub

' =========================================================
' رویداد: zkDevice_OnAttendance
' =========================================================
'
' وظیفه:
' هنگام تردد (اثر انگشت) دریافت می‌شود
' فوری به zkrt_HandleAttendance فراخوانی می‌شود
'
' پارامترها:
'   dwEnrollID: کد ثبتی کاربر
'   dwEmpID: شماره کارمندی
'   IsInvalid: اعتبار تردد (0=معتبر)
'   AttState: نوع تردد (0=ورود، 1=خروج)
'   VerifyFaceMask: ماسک صورت
'   VerifyMethod: روش تأیید (0=اثر انگشت)
'   dwYear: سال
'   dwMonth: ماه
'   dwDay: روز
'   dwHour: ساعت
'   dwMinute: دقیقه
'   dwSecond: ثانیه
'   dwWorkCode: کد کار
'
' =========================================================

Private Sub zkDevice_OnAttendance(ByVal dwEnrollID As Long, ByVal dwEmpID As Long, _
                                  ByVal IsInvalid As Long, ByVal AttState As Long, _
                                  ByVal VerifyFaceMask As Long, ByVal VerifyMethod As Long, _
                                  ByVal dwYear As Long, ByVal dwMonth As Long, _
                                  ByVal dwDay As Long, ByVal dwHour As Long, _
                                  ByVal dwMinute As Long, ByVal dwSecond As Long, _
                                  ByVal dwWorkCode As Long)
    On Error GoTo ErrHandler
    
    ' انتقال رویداد به ماژول مرکزی
    Call zkrt_HandleAttendance(deviceKey, deviceIP, devicePort, machineNumber, _
                               CStr(dwEnrollID), IsInvalid, AttState, VerifyMethod, _
                               dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkCode)
    
    Exit Sub
    
ErrHandler:
    Call LogError("zkDevice_OnAttendance", Err.Number, Err.Description, _
                  "Device=" & deviceKey & " Enroll=" & CStr(dwEnrollID))
End Sub

' =========================================================
' رویداد: zkDevice_OnEnrollStatus
' =========================================================
'
' وظیفه:
' هنگام تغییر شماره ثبتی فراخوانی می‌شود
'
' =========================================================

Private Sub zkDevice_OnEnrollStatus(ByVal dwEnrollID As Long, ByVal bEnrollStatus As Long)
    On Error Resume Next
    
    Call LogSystemEvent("clsZKRealtime.OnEnrollStatus", _
                       "تغییر ثبتی: Device=" & deviceKey & " EnrollID=" & CStr(dwEnrollID))
End Sub

' =========================================================
' رویداد: zkDevice_OnNewUser
' =========================================================
'
' وظیفه:
' هنگام اضافه شدن کاربر جدید فراخوانی می‌شود
'
' =========================================================

Private Sub zkDevice_OnNewUser(ByVal dwEnrollID As Long)
    On Error Resume Next
    
    Call LogSystemEvent("clsZKRealtime.OnNewUser", _
                       "کاربر جدید: Device=" & deviceKey & " EnrollID=" & CStr(dwEnrollID))
End Sub

' =========================================================
' رویداد: zkDevice_OnDeleteUser
' =========================================================
'
' وظیفه:
' هنگام حذف کاربر فراخوانی می‌شود
'
' =========================================================

Private Sub zkDevice_OnDeleteUser(ByVal dwEnrollID As Long)
    On Error Resume Next
    
    Call LogSystemEvent("clsZKRealtime.OnDeleteUser", _
                       "حذف کاربر: Device=" & deviceKey & " EnrollID=" & CStr(dwEnrollID))
End Sub

' =========================================================
' تابع: IsAttached
' =========================================================
'
' وظیفه:
' بررسی می‌کند که آیا Listener متصل است یا نه
'
' خروجی:
'   Boolean: True اگر متصل، False اگر قطع شده
'
' =========================================================

Public Function IsAttached() As Boolean
    IsAttached = isAttached
End Function

' =========================================================
' تابع: GetDeviceKey
' =========================================================
'
' وظیفه:
' کلید دستگاه را برمی‌گرداند
'
' خروجی:
'   String: کلید دستگاه
'
' =========================================================

Public Function GetDeviceKey() As String
    GetDeviceKey = deviceKey
End Function

' =========================================================
' Destructor
' =========================================================

Private Sub Class_Terminate()
    On Error Resume Next
    Call Detach
End Sub

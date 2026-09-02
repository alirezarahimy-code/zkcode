Option Compare Database
Option Explicit

' =========================================================
' ماژول: cfg.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول تمام ثابت‌های (Constants) پیکربندی برنامه را مدیریت می‌کند.
' تمام مقادیر ثابت در یک جا تعریف شده‌اند تا تغییر آن‌ها ساده باشد.
'
' کاربرد:
' - تنظیمات ارتباط با دستگاه ZK
' - تنظیمات خواندن لاگ‌ها
' - تنظیمات پردازش و چاپ
' - انواع چاپگر و وضعیت‌ها
' - اسامی تمام جداول
' - نوع‌های تردد
'
' مزایا:
' - اگر نیاز به تغییر داشت، فقط این فایل را تغییر می‌دهیم
' - تمام برنامه از این مقادیر استفاده می‌کند
' - جلوگیری از خطاهای تایپی (Typo)
'
' =========================================================

' =========================================================
' تنظیمات ارتباط دستگاه ZK
' =========================================================
' DEFAULT_ZK_PORT: پورت پیش‌فرض دستگاه‌های ZK
' MAX_RECONNECT_ATTEMPTS: تعداد تلاش دوباره برای اتصال مجدد
' DEVICE_TIMEOUT_SECONDS: مهلت زمانی برای پاسخ دستگاه

Public Const DEFAULT_ZK_PORT As Long = 4370
Public Const MAX_RECONNECT_ATTEMPTS As Long = 3
Public Const DEVICE_TIMEOUT_SECONDS As Long = 30

' =========================================================
' تنظیمات خواندن لاگ‌های تردد از دستگاه
' =========================================================
' READ_MAX_PER_CALL: حداکثر تعداد رکورد تردد خوانده‌شده در هر بار
' INITIAL_SYNC_MAX_READS: حداکثر تعداد برای اولین همگام‌سازی
'   (0 = فقط نقطه شروع را برقرار کن، تاریخ را نخوان)

Public Const READ_MAX_PER_CALL As Long = 2000
Public Const INITIAL_SYNC_MAX_READS As Long = 0

' =========================================================
' تنظیمات پردازش تردد و چاپ
' =========================================================
' PROCESS_BATCH_SIZE: تعداد رکورد‌های پردازش‌شده در هر دوره
' MAX_PRINT_ATTEMPTS: حداکثر تلاش برای چاپ هر فیش
' PRINTING_STUCK_TIMEOUT_MINUTES: مهلت‌زمانی قبل از شمردن چاپ "گیر‌کرده"

Public Const PROCESS_BATCH_SIZE As Long = 100
Public Const MAX_PRINT_ATTEMPTS As Long = 3
Public Const PRINTING_STUCK_TIMEOUT_MINUTES As Long = 10

' =========================================================
' انواع چاپگر
' =========================================================
' PRINTER_TYPE_WINDOWS: چاپگر سیستم Windows
' PRINTER_TYPE_FOLDER: ذخیره فیش در پوشه مشخص (UNC Path)
' PRINTER_TYPE_SOCKET: ارسال به چاپگر از طریق TCP Socket

Public Const PRINTER_TYPE_WINDOWS As String = "WINDOWS"
Public Const PRINTER_TYPE_FOLDER As String = "FOLDER"
Public Const PRINTER_TYPE_SOCKET As String = "SOCKET"

' =========================================================
' وضعیت‌های چاپ
' =========================================================
' PRINT_STATUS_PENDING: منتظر چاپ
' PRINT_STATUS_SUBMITTED: ارسال شده به چاپگر
' PRINT_STATUS_SUCCESS: چاپ موفق
' PRINT_STATUS_FAILED: چاپ ناموفق
' PRINT_STATUS_UNKNOWN: نتیجه نامعلوم

Public Const PRINT_STATUS_PENDING As String = "PENDING"
Public Const PRINT_STATUS_SUBMITTED As String = "SUBMITTED"
Public Const PRINT_STATUS_SUCCESS As String = "SUCCESS"
Public Const PRINT_STATUS_FAILED As String = "FAILED"
Public Const PRINT_STATUS_UNKNOWN As String = "UNKNOWN"

' =========================================================
' نتایج پردازش تردد
' =========================================================
' NEW: تردد جدید (هنوز پردازش نشده)
' PROCESSING: در حال پردازش
' PRINTED: فیش چاپ شد
' PRINT_FAILED: چاپ ناموفق
' NO_EMPLOYEE: کارمند یافت نشد
' WAITING_FOR_MEAL: غذا سفارش نشده یا منتظر تأیید
' ALREADY_PRINTED: فیش قبلاً چاپ شده است
' RECEIPT_FAILED: خطا در ایجاد فیش
' MEAL_FINALIZE_FAILED: خطا در تکمیل غذا
' PRINT_UNKNOWN: نتیجه چاپ نامعلوم

Public Const RESULT_NEW As String = "NEW"
Public Const RESULT_PROCESSING As String = "PROCESSING"
Public Const RESULT_PRINTED As String = "PRINTED"
Public Const RESULT_PRINT_FAILED As String = "PRINT_FAILED"
Public Const RESULT_NO_EMPLOYEE As String = "NO_EMPLOYEE"
Public Const RESULT_NO_MEAL_ORDER As String = "WAITING_FOR_MEAL"
Public Const RESULT_ALREADY_PRINTED As String = "ALREADY_PRINTED"
Public Const RESULT_RECEIPT_FAILED As String = "RECEIPT_FAILED"
Public Const RESULT_MEAL_FINALIZE_FAILED As String = "MEAL_FINALIZE_FAILED"
Public Const RESULT_PRINT_UNKNOWN As String = "PRINT_UNKNOWN"

' =========================================================
' انواع تردد
' =========================================================
' IN: ورود
' OUT: خروج
' BREAK_OUT: خروج برای استراحت
' BREAK_IN: بازگشت از استراحت
' OT_IN: ورود برای کار اضافی
' OT_OUT: خروج کار اضافی

Public Const ATT_IN As String = "IN"
Public Const ATT_OUT As String = "OUT"
Public Const ATT_BREAK_OUT As String = "BREAK_OUT"
Public Const ATT_BREAK_IN As String = "BREAK_IN"
Public Const ATT_OT_IN As String = "OT_IN"
Public Const ATT_OT_OUT As String = "OT_OUT"

' =========================================================
' اسامی جداول
' =========================================================
' هر جدول یک نقش خاص دارد:
' - tblEmployees: اطلاعات کارمندان
' - tblFoodOrders: سفارشات غذا
' - tblZKDevices: دستگاه‌های ZK
' - tblAttendanceRecords: تمام تردد‌های ثبت‌شده
' - tblPrintedReceipts: فیش‌های چاپ‌شده
' - tblDeviceState: وضعیت هر دستگاه
' - tblSystemLogs: لاگ‌های سیستم
' - tblErrorLogs: لاگ‌های خطا
' - tblDailyMealList: لیست غذای روزانه
' - tblPrinterSettings: تنظیمات چاپگر
' - tblUserTemplates: الگوهای کاربری

Public Const TABLE_EMPLOYEES As String = "tblEmployees"
Public Const TABLE_FOOD_ORDERS As String = "tblFoodOrders"
Public Const TABLE_ZK_DEVICES As String = "tblZKDevices"
Public Const TABLE_ATTENDANCE As String = "tblAttendanceRecords"
Public Const TABLE_RECEIPTS As String = "tblPrintedReceipts"
Public Const TABLE_DEVICE_STATE As String = "tblDeviceState"
Public Const TABLE_SYSTEM_LOGS As String = "tblSystemLogs"
Public Const TABLE_ERROR_LOGS As String = "tblErrorLogs"
Public Const TABLE_DAILY_MEALS As String = "tblDailyMealList"
Public Const TABLE_PRINTER_SETTINGS As String = "tblPrinterSettings"
Public Const TABLE_TEMPLATES As String = "tblUserTemplates"
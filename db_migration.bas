Option Compare Database
Option Explicit

' =========================================================
' ماژول: db_migration.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول تمام جداول پایگاه‌داده را ایجاد و مدیریت می‌کند.
' هنگام اولین اجرای برنامه (یا به‌روزرسانی)، این ماژول
' اطمینان می‌دهد که تمام جداول و فیلدهای لازم موجود هستند.
'
' کاربرد:
' - ایجاد جداول (اگر وجود نداشت)
' - اضافه کردن فیلدهای جدید (به جداول موجود)
' - ایجاد Indexes برای سرعت
' - بهینه‌سازی نوع‌های داده (Data Types)
'
' ویژگی‌ها:
' - Idempotent: می‌تواند چند بار اجرا شود، مشکل نیست
' - تمام جداول با کلید اصلی (Primary Key) دارند
' - Indexes برای بهتری Search و Join
' - مدیریت NULL values برای Unique Indexes
'
' =========================================================

' =========================================================
' تابع: db_migration_CreateOrUpdateSchema
' =========================================================
' 
' وظیفه:
' تمام جداول و ساختار پایگاه‌داده را ایجاد یا بروزرسانی می‌کند
' این تابع باید در شروع برنامه فراخوانی شود
'
' پارامتر: ندارد
'
' خروجی: Boolean
'   True: ایجاد/بروزرسانی موفق
'   False: خطا رخ داد
'
' خطاهای ممکن:
' - اگر دسترسی نوشتن نبود
' - اگر فضای دیسک کافی نبود
' - اگر فایل DB قفل شده بود
'
' نمونه استفاده:
'   If db_migration_CreateOrUpdateSchema() Then
'       MsgBox "پایگاه‌داده آماده شد"
'   Else
'       MsgBox "خطا در ایجاد پایگاه‌داده"
'   End If
'
' نکات مهم:
' - این تابع فقط یک بار در شروع برنامه اجرا شود
' - اطلاعات موجود محفوظ می‌ماند
' - فقط جداول و فیلدهای جدید اضافه می‌شود
'
' =========================================================

Public Function db_migration_CreateOrUpdateSchema() As Boolean
    On Error GoTo ErrorHandler
    
    Application.Echo False, "در حال ایجاد یا بروزرسانی پایگاه‌داده..."
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' ایجاد تمام جداول
    EnsureEmployeesTable db
    EnsureFoodOrdersTable db
    EnsureZKDevicesTable db
    EnsureAttendanceTable db
    EnsurePrintedReceiptsTable db
    EnsureDeviceStateTable db
    EnsureSystemLogsTable db
    EnsureErrorLogsTable db
    EnsureDailyMealListTable db
    EnsurePrinterSettingsTable db
    EnsureUserTemplatesTable db
    EnsureLiveMonitoringTable db
    
    ' نرمال‌سازی NULL values برای Unique Indexes
    On Error Resume Next
    db.Execute "UPDATE tblEmployees SET DeviceUserID=Null WHERE DeviceUserID=''", dbFailOnError
    db.Execute "UPDATE tblAttendanceRecords SET RawKey=Null WHERE RawKey=''", dbFailOnError
    db.Execute "UPDATE tblDeviceState SET DeviceKey=Null WHERE DeviceKey=''", dbFailOnError
    On Error GoTo ErrorHandler
    
    ' ایجاد Indexes برای سرعت
    EnsureUniqueIndex db, TABLE_EMPLOYEES, "idxEmployees_DeviceUserID", "DeviceUserID", True
    EnsureUniqueIndex db, TABLE_ATTENDANCE, "idxAttendance_RawKey", "RawKey", True
    EnsureUniqueIndex db, TABLE_RECEIPTS, "idxReceipts_AttendanceUnique", "AttendanceRecordID", True
    EnsureUniqueIndex db, TABLE_DEVICE_STATE, "idxDeviceState_Key", "DeviceKey", True
    EnsureUniqueCompositeIndex db, TABLE_TEMPLATES, "idxTemplates_EmployeeVersion", "EmployeeID", "VersionNo"
    EnsureIndex db, TABLE_DAILY_MEALS, "idxMeal_EmployeeDate", "EmployeeID", "MealDate"
    
    Application.Echo True
    
    Call LogSystemEvent("db_migration", "پایگاه‌داده آماده شد")
    db_migration_CreateOrUpdateSchema = True
    
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    Application.Echo True
    Call LogError("db_migration_CreateOrUpdateSchema", Err.Number, Err.Description, "")
    db_migration_CreateOrUpdateSchema = False
    Set db = Nothing
End Function

' =========================================================
' جداول و فیلدهای آن‌ها
' =========================================================

Private Sub EnsureEmployeesTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_EMPLOYEES) Then
        EnsureColumn db, TABLE_EMPLOYEES, "DeviceUserID TEXT(50)"
        EnsureColumn db, TABLE_EMPLOYEES, "FullName TEXT(255)"
        EnsureColumn db, TABLE_EMPLOYEES, "NationalCode TEXT(50)"
        EnsureColumn db, TABLE_EMPLOYEES, "CardNumber TEXT(50)"
        EnsureColumn db, TABLE_EMPLOYEES, "IsActive YESNO"
        EnsureColumn db, TABLE_EMPLOYEES, "EnrollDate DATETIME"
        EnsureColumn db, TABLE_EMPLOYEES, "CreatedDate DATETIME"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_EMPLOYEES & "] (" & _
          "EmployeeID TEXT(50) PRIMARY KEY, " & _
          "DeviceUserID TEXT(50), " & _
          "FullName TEXT(255), " & _
          "NationalCode TEXT(50), " & _
          "CardNumber TEXT(50), " & _
          "IsActive YESNO DEFAULT True, " & _
          "EnrollDate DATETIME, " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول کارمندان ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureEmployeesTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureFoodOrdersTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_FOOD_ORDERS) Then
        EnsureColumn db, TABLE_FOOD_ORDERS, "UserID TEXT(50)"
        EnsureColumn db, TABLE_FOOD_ORDERS, "NationalCode TEXT(20)"
        EnsureColumn db, TABLE_FOOD_ORDERS, "FirstName TEXT(100)"
        EnsureColumn db, TABLE_FOOD_ORDERS, "LastName TEXT(150)"
        EnsureColumn db, TABLE_FOOD_ORDERS, "MealType TEXT(100)"
        EnsureColumn db, TABLE_FOOD_ORDERS, "OrderDate DATETIME"
        EnsureColumn db, TABLE_FOOD_ORDERS, "ReserveDate DATETIME"
        EnsureColumn db, TABLE_FOOD_ORDERS, "ReserveTime DATETIME"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_FOOD_ORDERS & "] (" & _
          "OrderID AUTOINCREMENT PRIMARY KEY, " & _
          "UserID TEXT(50), " & _
          "NationalCode TEXT(20), " & _
          "FirstName TEXT(100), " & _
          "LastName TEXT(150), " & _
          "MealType TEXT(100), " & _
          "OrderDate DATETIME, " & _
          "ReserveDate DATETIME, " & _
          "ReserveTime DATETIME, " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول سفارشات غذا ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureFoodOrdersTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureZKDevicesTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_ZK_DEVICES) Then
        EnsureColumn db, TABLE_ZK_DEVICES, "DeviceName TEXT(255)"
        EnsureColumn db, TABLE_ZK_DEVICES, "DeviceIP TEXT(50)"
        EnsureColumn db, TABLE_ZK_DEVICES, "DevicePort LONG"
        EnsureColumn db, TABLE_ZK_DEVICES, "MachineNumber LONG"
        EnsureColumn db, TABLE_ZK_DEVICES, "CommKey LONG"
        EnsureColumn db, TABLE_ZK_DEVICES, "IsActive YESNO"
        EnsureColumn db, TABLE_ZK_DEVICES, "LastConnectionAttempt DATETIME"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_ZK_DEVICES & "] (" & _
          "DeviceID AUTOINCREMENT PRIMARY KEY, " & _
          "DeviceName TEXT(255), " & _
          "DeviceIP TEXT(50), " & _
          "DevicePort LONG DEFAULT " & DEFAULT_ZK_PORT & ", " & _
          "MachineNumber LONG DEFAULT 1, " & _
          "CommKey LONG DEFAULT 0, " & _
          "IsActive YESNO DEFAULT True, " & _
          "LastConnectionAttempt DATETIME, " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول دستگاه‌های ZK ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureZKDevicesTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureAttendanceTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_ATTENDANCE) Then
        EnsureColumn db, TABLE_ATTENDANCE, "DeviceEnrollID TEXT(50)"
        EnsureColumn db, TABLE_ATTENDANCE, "EmployeeID TEXT(50)"
        EnsureColumn db, TABLE_ATTENDANCE, "FullName TEXT(255)"
        EnsureColumn db, TABLE_ATTENDANCE, "AttendanceDateTime DATETIME"
        EnsureColumn db, TABLE_ATTENDANCE, "AttendanceType TEXT(50)"
        EnsureColumn db, TABLE_ATTENDANCE, "DeviceIP TEXT(50)"
        EnsureColumn db, TABLE_ATTENDANCE, "DevicePort LONG"
        EnsureColumn db, TABLE_ATTENDANCE, "DeviceMachineNumber LONG"
        EnsureColumn db, TABLE_ATTENDANCE, "DeviceKey TEXT(150)"
        EnsureColumn db, TABLE_ATTENDANCE, "RawKey TEXT(255)"
        EnsureColumn db, TABLE_ATTENDANCE, "IsProcessed YESNO"
        EnsureColumn db, TABLE_ATTENDANCE, "ProcessingResult TEXT(50)"
        EnsureColumn db, TABLE_ATTENDANCE, "ReceiptID LONG"
        EnsureColumn db, TABLE_ATTENDANCE, "PrintAttempts LONG"
        EnsureColumn db, TABLE_ATTENDANCE, "LastPrintAttempt DATETIME"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_ATTENDANCE & "] (" & _
          "RecordID AUTOINCREMENT PRIMARY KEY, " & _
          "DeviceEnrollID TEXT(50), " & _
          "EmployeeID TEXT(50), " & _
          "FullName TEXT(255), " & _
          "AttendanceDateTime DATETIME, " & _
          "AttendanceType TEXT(50), " & _
          "DeviceIP TEXT(50), " & _
          "DevicePort LONG, " & _
          "DeviceMachineNumber LONG, " & _
          "DeviceKey TEXT(150), " & _
          "RawKey TEXT(255), " & _
          "IsProcessed YESNO DEFAULT False, " & _
          "ProcessingResult TEXT(50) DEFAULT 'NEW', " & _
          "ReceiptID LONG, " & _
          "PrintAttempts LONG DEFAULT 0, " & _
          "LastPrintAttempt DATETIME, " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول تردد‌ها ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureAttendanceTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsurePrintedReceiptsTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_RECEIPTS) Then
        EnsureColumn db, TABLE_RECEIPTS, "EmployeeID TEXT(50)"
        EnsureColumn db, TABLE_RECEIPTS, "MealListID LONG"
        EnsureColumn db, TABLE_RECEIPTS, "AttendanceRecordID LONG"
        EnsureColumn db, TABLE_RECEIPTS, "PendingDateTime DATETIME"
        EnsureColumn db, TABLE_RECEIPTS, "PrintDateTime DATETIME"
        EnsureColumn db, TABLE_RECEIPTS, "PrinterName TEXT(255)"
        EnsureColumn db, TABLE_RECEIPTS, "PrinterSharePath TEXT(255)"
        EnsureColumn db, TABLE_RECEIPTS, "PrinterIP TEXT(50)"
        EnsureColumn db, TABLE_RECEIPTS, "PrinterPort LONG"
        EnsureColumn db, TABLE_RECEIPTS, "PrinterType TEXT(50)"
        EnsureColumn db, TABLE_RECEIPTS, "IsSuccessful YESNO"
        EnsureColumn db, TABLE_RECEIPTS, "PrintStatus TEXT(30)"
        EnsureColumn db, TABLE_RECEIPTS, "ErrorMessage MEMO"
        EnsureColumn db, TABLE_RECEIPTS, "IPAddress TEXT(50)"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_RECEIPTS & "] (" & _
          "ReceiptID AUTOINCREMENT PRIMARY KEY, " & _
          "EmployeeID TEXT(50), " & _
          "MealListID LONG, " & _
          "AttendanceRecordID LONG, " & _
          "PendingDateTime DATETIME, " & _
          "PrintDateTime DATETIME, " & _
          "PrinterName TEXT(255), " & _
          "PrinterSharePath TEXT(255), " & _
          "PrinterIP TEXT(50), " & _
          "PrinterPort LONG, " & _
          "PrinterType TEXT(50), " & _
          "IsSuccessful YESNO DEFAULT False, " & _
          "PrintStatus TEXT(30) DEFAULT 'PENDING', " & _
          "ErrorMessage MEMO, " & _
          "IPAddress TEXT(50), " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول فیش‌های چاپ‌شده ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsurePrintedReceiptsTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureDeviceStateTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_DEVICE_STATE) Then
        EnsureColumn db, TABLE_DEVICE_STATE, "DeviceKey TEXT(150)"
        EnsureColumn db, TABLE_DEVICE_STATE, "DeviceIP TEXT(50)"
        EnsureColumn db, TABLE_DEVICE_STATE, "DevicePort LONG"
        EnsureColumn db, TABLE_DEVICE_STATE, "MachineNumber LONG"
        EnsureColumn db, TABLE_DEVICE_STATE, "LastProcessedDateTime DATETIME"
        EnsureColumn db, TABLE_DEVICE_STATE, "LastProcessedKey TEXT(255)"
        EnsureColumn db, TABLE_DEVICE_STATE, "LastSuccessfulSync DATETIME"
        EnsureColumn db, TABLE_DEVICE_STATE, "LastSyncAttempt DATETIME"
        EnsureColumn db, TABLE_DEVICE_STATE, "LastSyncError TEXT(255)"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_DEVICE_STATE & "] (" & _
          "StateID AUTOINCREMENT PRIMARY KEY, " & _
          "DeviceKey TEXT(150), " & _
          "DeviceIP TEXT(50), " & _
          "DevicePort LONG, " & _
          "MachineNumber LONG, " & _
          "LastProcessedDateTime DATETIME, " & _
          "LastProcessedKey TEXT(255), " & _
          "LastSuccessfulSync DATETIME, " & _
          "LastSyncAttempt DATETIME, " & _
          "LastSyncError TEXT(255), " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول وضعیت دستگاه ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureDeviceStateTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureSystemLogsTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_SYSTEM_LOGS) Then
        EnsureColumn db, TABLE_SYSTEM_LOGS, "LogDate DATETIME"
        EnsureColumn db, TABLE_SYSTEM_LOGS, "Source TEXT(100)"
        EnsureColumn db, TABLE_SYSTEM_LOGS, "Message MEMO"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_SYSTEM_LOGS & "] (" & _
          "LogID AUTOINCREMENT PRIMARY KEY, " & _
          "LogDate DATETIME DEFAULT Now(), " & _
          "Source TEXT(100), " & _
          "Message MEMO" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول لاگ‌های سیستم ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureSystemLogsTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureErrorLogsTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_ERROR_LOGS) Then
        EnsureColumn db, TABLE_ERROR_LOGS, "ErrorDate DATETIME"
        EnsureColumn db, TABLE_ERROR_LOGS, "ErrNumber LONG"
        EnsureColumn db, TABLE_ERROR_LOGS, "ErrDescription MEMO"
        EnsureColumn db, TABLE_ERROR_LOGS, "[Procedure] TEXT(255)"
        EnsureColumn db, TABLE_ERROR_LOGS, "AdditionalInfo MEMO"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_ERROR_LOGS & "] (" & _
          "ErrorID AUTOINCREMENT PRIMARY KEY, " & _
          "ErrorDate DATETIME DEFAULT Now(), " & _
          "ErrNumber LONG, " & _
          "ErrDescription MEMO, " & _
          "[Procedure] TEXT(255), " & _
          "AdditionalInfo MEMO" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول لاگ‌های خطا ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureErrorLogsTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureDailyMealListTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_DAILY_MEALS) Then
        EnsureColumn db, TABLE_DAILY_MEALS, "MealDate DATETIME"
        EnsureColumn db, TABLE_DAILY_MEALS, "EmployeeID TEXT(50)"
        EnsureColumn db, TABLE_DAILY_MEALS, "FullName TEXT(255)"
        EnsureColumn db, TABLE_DAILY_MEALS, "MealType TEXT(100)"
        EnsureColumn db, TABLE_DAILY_MEALS, "Quantity LONG"
        EnsureColumn db, TABLE_DAILY_MEALS, "IsConfirmed YESNO"
        EnsureColumn db, TABLE_DAILY_MEALS, "MealDelivered YESNO"
        EnsureColumn db, TABLE_DAILY_MEALS, "DeliveryReceiptID LONG"
        EnsureColumn db, TABLE_DAILY_MEALS, "DeliveredDateTime DATETIME"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_DAILY_MEALS & "] (" & _
          "MealListID AUTOINCREMENT PRIMARY KEY, " & _
          "MealDate DATETIME, " & _
          "EmployeeID TEXT(50), " & _
          "FullName TEXT(255), " & _
          "MealType TEXT(100), " & _
          "Quantity LONG DEFAULT 1, " & _
          "IsConfirmed YESNO DEFAULT True, " & _
          "MealDelivered YESNO DEFAULT False, " & _
          "DeliveryReceiptID LONG, " & _
          "DeliveredDateTime DATETIME, " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول لیست غذای روزانه ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureDailyMealListTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsurePrinterSettingsTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_PRINTER_SETTINGS) Then
        EnsureColumn db, TABLE_PRINTER_SETTINGS, "PrinterName TEXT(255)"
        EnsureColumn db, TABLE_PRINTER_SETTINGS, "PrinterSharePath TEXT(255)"
        EnsureColumn db, TABLE_PRINTER_SETTINGS, "PrinterIP TEXT(50)"
        EnsureColumn db, TABLE_PRINTER_SETTINGS, "PrinterPort LONG"
        EnsureColumn db, TABLE_PRINTER_SETTINGS, "PrinterType TEXT(50)"
        EnsureColumn db, TABLE_PRINTER_SETTINGS, "IsDefault YESNO"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_PRINTER_SETTINGS & "] (" & _
          "PrinterID AUTOINCREMENT PRIMARY KEY, " & _
          "PrinterName TEXT(255), " & _
          "PrinterSharePath TEXT(255), " & _
          "PrinterIP TEXT(50), " & _
          "PrinterPort LONG DEFAULT 9100, " & _
          "PrinterType TEXT(50) DEFAULT '" & PRINTER_TYPE_WINDOWS & "', " & _
          "IsDefault YESNO DEFAULT False, " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول تنظیمات چاپگر ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsurePrinterSettingsTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureUserTemplatesTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_TEMPLATES) Then
        EnsureColumn db, TABLE_TEMPLATES, "EmployeeID TEXT(50)"
        EnsureColumn db, TABLE_TEMPLATES, "TemplateBase64 MEMO"
        EnsureColumn db, TABLE_TEMPLATES, "TemplateType TEXT(50)"
        EnsureColumn db, TABLE_TEMPLATES, "VersionNo LONG"
        EnsureColumn db, TABLE_TEMPLATES, "IsActive YESNO"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_TEMPLATES & "] (" & _
          "TemplateID AUTOINCREMENT PRIMARY KEY, " & _
          "EmployeeID TEXT(50), " & _
          "TemplateBase64 MEMO, " & _
          "TemplateType TEXT(50), " & _
          "VersionNo LONG DEFAULT 1, " & _
          "IsActive YESNO DEFAULT True, " & _
          "CreatedDate DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول الگوهای کاربری ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureUserTemplatesTable", Err.Number, Err.Description, "")
End Sub

Private Sub EnsureLiveMonitoringTable(ByVal db As DAO.Database)
    On Error Resume Next
    If TableExists(TABLE_LIVE_MONITORING) Then
        EnsureColumn db, TABLE_LIVE_MONITORING, "EnrollID TEXT(50)"
        EnsureColumn db, TABLE_LIVE_MONITORING, "Result TEXT(100)"
        EnsureColumn db, TABLE_LIVE_MONITORING, "Details MEMO"
        EnsureColumn db, TABLE_LIVE_MONITORING, "LogDateTime DATETIME"
        Exit Sub
    End If
    On Error GoTo EH
    
    Dim sql As String
    sql = "CREATE TABLE [" & TABLE_LIVE_MONITORING & "] (" & _
          "MonitorID AUTOINCREMENT PRIMARY KEY, " & _
          "EnrollID TEXT(50), " & _
          "Result TEXT(100), " & _
          "Details MEMO, " & _
          "LogDateTime DATETIME DEFAULT Now()" & _
          ")"
    
    db.Execute sql, dbFailOnError
    Call LogSystemEvent("db_migration", "جدول پایش آنی ایجاد شد")
    Exit Sub
EH:
    Call LogError("EnsureLiveMonitoringTable", Err.Number, Err.Description, "")
End Sub

' =========================================================
' توابع کمکی: Indexes
' =========================================================

Private Sub EnsureUniqueIndex(ByVal db As DAO.Database, ByVal tableName As String, _
                              ByVal indexName As String, ByVal fieldName As String, _
                              Optional ByVal ignoreNulls As Boolean = False)
    On Error GoTo EH
    
    Dim td As DAO.TableDef, idx As DAO.Index, f As DAO.Field
    Dim deleteExisting As Boolean
    
    Set td = db.TableDefs(tableName)
    
    For Each idx In td.Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            If idx.Unique And idx.Fields.Count = 1 Then
                If StrComp(idx.Fields(0).Name, fieldName, vbTextCompare) = 0 Then
                    On Error Resume Next
                    idx.IgnoreNulls = ignoreNulls
                    On Error GoTo EH
                    Exit Sub
                End If
            End If
            deleteExisting = True
            Exit For
        End If
    Next idx
    
    If deleteExisting Then td.Indexes.Delete indexName
    
    Set idx = td.CreateIndex(indexName)
    idx.Unique = True
    On Error Resume Next
    idx.IgnoreNulls = ignoreNulls
    On Error GoTo EH
    
    Set f = idx.CreateField(fieldName)
    idx.Fields.Append f
    td.Indexes.Append idx
    
    Exit Sub
    
EH:
    Call LogError("EnsureUniqueIndex", Err.Number, Err.Description, tableName & "." & indexName)
End Sub

Private Sub EnsureUniqueCompositeIndex(ByVal db As DAO.Database, ByVal tableName As String, _
                                       ByVal indexName As String, ByVal field1 As String, _
                                       ByVal field2 As String)
    On Error GoTo EH
    
    Dim td As DAO.TableDef, idx As DAO.Index, f As DAO.Field
    Dim deleteExisting As Boolean
    
    Set td = db.TableDefs(tableName)
    
    For Each idx In td.Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            If idx.Unique And idx.Fields.Count = 2 Then
                If StrComp(idx.Fields(0).Name, field1, vbTextCompare) = 0 And _
                   StrComp(idx.Fields(1).Name, field2, vbTextCompare) = 0 Then
                    Exit Sub
                End If
            End If
            deleteExisting = True
            Exit For
        End If
    Next idx
    
    If deleteExisting Then td.Indexes.Delete indexName
    
    Set idx = td.CreateIndex(indexName)
    idx.Unique = True
    
    Set f = idx.CreateField(field1)
    idx.Fields.Append f
    Set f = idx.CreateField(field2)
    idx.Fields.Append f
    
    td.Indexes.Append idx
    
    Exit Sub
    
EH:
    Call LogError("EnsureUniqueCompositeIndex", Err.Number, Err.Description, tableName & "." & indexName)
End Sub

Private Sub EnsureIndex(ByVal db As DAO.Database, ByVal tableName As String, _
                        ByVal indexName As String, ByVal field1 As String, ByVal field2 As String)
    On Error Resume Next
    
    Dim td As DAO.TableDef, idx As DAO.Index, f As DAO.Field
    
    Set td = db.TableDefs(tableName)
    
    For Each idx In td.Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then Exit Sub
    Next idx
    
    Set idx = td.CreateIndex(indexName)
    idx.Unique = False
    
    Set f = idx.CreateField(field1)
    idx.Fields.Append f
    Set f = idx.CreateField(field2)
    idx.Fields.Append f
    
    td.Indexes.Append idx
End Sub

' =========================================================
' توابع کمکی: Column & Table Management
' =========================================================

Private Sub EnsureColumn(ByVal db As DAO.Database, ByVal tableName As String, ByVal columnSql As String)
    On Error Resume Next
    
    Dim columnName As String
    Dim parts() As String
    
    parts = Split(columnSql, " ")
    columnName = Replace(Replace(parts(0), "[", ""), "]", "")
    
    If ColumnExists(tableName, columnName) Then Exit Sub
    
    On Error GoTo EH
    db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN " & columnSql, dbFailOnError
    Exit Sub
    
EH:
    Call LogError("EnsureColumn", Err.Number, Err.Description, tableName & "." & columnName)
End Sub

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

Private Function ColumnExists(ByVal TableName As String, ByVal ColumnName As String) As Boolean
    On Error GoTo EH
    
    Dim f As DAO.Field
    
    For Each f In CurrentDb().TableDefs(TableName).Fields
        If StrComp(f.Name, ColumnName, vbTextCompare) = 0 Then
            ColumnExists = True
            Exit Function
        End If
    Next f
    
    Exit Function
    
EH:
    ColumnExists = False
End Function
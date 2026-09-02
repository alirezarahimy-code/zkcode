Option Compare Database
Option Explicit

' =========================================================
' ماژول: csv_utils.bas
' =========================================================
'
' توضیح ماژول:
' این ماژول صادرات گزارش‌ها به فرمت CSV انجام می‌دهد.
' گزارش‌های تردد، سفارشات، و چاپ‌های موفق به فایل تبدیل می‌شوند.
'
' کاربرد:
' - صادرات گزارش تردد‌ها
' - صادرات گزارش سفارشات غذا
' - صادرات گزارش چاپ‌های موفق
' - استفاده برای تحلیل و بررسی
'
' ویژگی‌های مهم:
' - فارسی کامل
' - Encoding UTF-8
' - فرمت‌بندی صحیح CSV
' - حفاظت از ویرگول و نقل‌قول
'
' معماری:
' - export_AttendanceReport: صادرات تردد‌ها
' - export_MealOrdersReport: صادرات سفارشات
' - export_SuccessfulPrints: صادرات چاپ‌های موفق
' - csv_EscapeField: محافظت از داده‌های CSV
'
' =========================================================

' =========================================================
' تابع: export_AttendanceReport
' =========================================================
'
' وظیفه:
' تمام تردد‌های یک روز (یا بازهٌ زمانی) را به CSV صادر می‌کند
'
' پارامترها:
'   fromDate (Date): تاریخ شروع
'   toDate (Date): تاریخ پایان
'   filePath (String): مسیر فایل خروجی
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' نمونه استفاده:
'   If export_AttendanceReport(Date(), Date(), "C:\Reports\Attendance.csv") Then
'       MsgBox "گزارش صادر شد"
'   End If
'
' فایل خروجی:
'   شماره رکورد,کد ثبتی,نام کامل,تاریخ شمسی,ساعت,نوع تردد,نتیجهٌ پردازش,وضعیت چاپ
'
' =========================================================

Public Function export_AttendanceReport(ByVal fromDate As Date, ByVal toDate As Date, _
                                        ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fso As Object
    Dim file As Object
    Dim line As String
    Dim jDate As String, jTime As String
    Dim count As Long
    
    Set db = CurrentDb()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' دریافت تردد‌ها
    sql = "SELECT RecordID, DeviceEnrollID, FullName, AttendanceDateTime, " & _
          "AttendanceType, ProcessingResult, CreatedDate " & _
          "FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE AttendanceDateTime >= " & SqlDateTime(fromDate) & " " & _
          "AND AttendanceDateTime < " & SqlDateTime(toDate + 1) & " " & _
          "ORDER BY AttendanceDateTime DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' ایجاد فایل
    Set file = fso.CreateTextFile(filePath, True, True)
    
    ' نوشتن Header (سرتیتر)
    line = "شماره رکورد,کد ثبتی,نام کامل,تاریخ شمسی,ساعت,نوع تردد,نتیجهٌ پردازش"
    file.WriteLine line
    
    ' نوشتن داده‌ها
    count = 0
    Do While Not rs.EOF
        jDate = format_jalali_date(Nz(rs!AttendanceDateTime, Now()), "SHORT")
        jTime = Format$(Nz(rs!AttendanceDateTime, Now()), "hh:mm:ss")
        
        line = csv_EscapeField(CStr(Nz(rs!RecordID, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!DeviceEnrollID, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!FullName, ""))) & "," & _
               csv_EscapeField(jDate) & "," & _
               csv_EscapeField(jTime) & "," & _
               csv_EscapeField(CStr(Nz(rs!AttendanceType, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!ProcessingResult, "")))
        
        file.WriteLine line
        count = count + 1
        rs.MoveNext
    Loop
    
    file.Close
    
    Call LogSystemEvent("export_AttendanceReport", _
                       "گزارش تردد صادر شد: " & CStr(count) & " رکورد به " & filePath)
    
    export_AttendanceReport = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set file = Nothing
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("export_AttendanceReport", Err.Number, Err.Description, filePath)
    export_AttendanceReport = False
    Resume CleanExit
End Function

' =========================================================
' تابع: export_MealOrdersReport
' =========================================================
'
' وظیفه:
' تمام سفارشات غذای یک روز را به CSV صادر می‌کند
'
' پارامترها:
'   fromDate (Date): تاریخ شروع
'   toDate (Date): تاریخ پایان
'   filePath (String): مسیر فایل خروجی
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' فایل خروجی:
'   کد کارمندی,نام کامل,کد ملی,نوع غذا,تاریخ رزرو,تاریخ سفارش
'
' =========================================================

Public Function export_MealOrdersReport(ByVal fromDate As Date, ByVal toDate As Date, _
                                        ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fso As Object
    Dim file As Object
    Dim line As String
    Dim jDate As String
    Dim count As Long
    
    Set db = CurrentDb()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' دریافت سفارشات
    sql = "SELECT EmployeeID, FirstName, LastName, NationalCode, MealType, " & _
          "ReserveDate, CreatedDate " & _
          "FROM " & TABLE_MEAL_ORDERS & " " & _
          "WHERE ReserveDate >= " & SqlDateTime(fromDate) & " " & _
          "AND ReserveDate < " & SqlDateTime(toDate + 1) & " " & _
          "ORDER BY ReserveDate DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' ایجاد فایل
    Set file = fso.CreateTextFile(filePath, True, True)
    
    ' نوشتن Header
    line = "کد کارمندی,نام,نام خانوادگی,کد ملی,نوع غذا,تاریخ رزرو,تاریخ سفارش"
    file.WriteLine line
    
    ' نوشتن داده‌ها
    count = 0
    Do While Not rs.EOF
        jDate = format_jalali_date(Nz(rs!ReserveDate, Now()), "SHORT")
        
        line = csv_EscapeField(CStr(Nz(rs!EmployeeID, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!FirstName, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!LastName, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!NationalCode, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!MealType, ""))) & "," & _
               csv_EscapeField(jDate) & "," & _
               csv_EscapeField(Format$(Nz(rs!CreatedDate, Now()), "yyyy/mm/dd hh:mm:ss"))
        
        file.WriteLine line
        count = count + 1
        rs.MoveNext
    Loop
    
    file.Close
    
    Call LogSystemEvent("export_MealOrdersReport", _
                       "گزارش سفارشات صادر شد: " & CStr(count) & " رکورد به " & filePath)
    
    export_MealOrdersReport = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set file = Nothing
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("export_MealOrdersReport", Err.Number, Err.Description, filePath)
    export_MealOrdersReport = False
    Resume CleanExit
End Function

' =========================================================
' تابع: export_SuccessfulPrints
' =========================================================
'
' وظیفه:
' تمام چاپ‌های موفق فیش‌ها را به CSV صادر می‌کند
'
' پارامترها:
'   fromDate (Date): تاریخ شروع
'   toDate (Date): تاریخ پایان
'   filePath (String): مسیر فایل خروجی
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' فایل خروجی:
'   شماره فیش,کد کارمندی,نام کامل,نوع غذا,تاریخ چاپ,ساعت چاپ
'
' =========================================================

Public Function export_SuccessfulPrints(ByVal fromDate As Date, ByVal toDate As Date, _
                                        ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fso As Object
    Dim file As Object
    Dim line As String
    Dim jDate As String, jTime As String
    Dim count As Long
    
    Set db = CurrentDb()
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' دریافت چاپ‌های موفق
    sql = "SELECT ReceiptID, EmployeeID, EmployeeFullName, MealType, PrintDateTime " & _
          "FROM " & TABLE_RECEIPTS & " " & _
          "WHERE PrintStatus='" & PRINT_STATUS_SUCCESS & "' " & _
          "AND PrintDateTime >= " & SqlDateTime(fromDate) & " " & _
          "AND PrintDateTime < " & SqlDateTime(toDate + 1) & " " & _
          "ORDER BY PrintDateTime DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' ایجاد فایل
    Set file = fso.CreateTextFile(filePath, True, True)
    
    ' نوشتن Header
    line = "شماره فیش,کد کارمندی,نام کامل,نوع غذا,تاریخ چاپ,ساعت چاپ"
    file.WriteLine line
    
    ' نوشتن داده‌ها
    count = 0
    Do While Not rs.EOF
        jDate = format_jalali_date(Nz(rs!PrintDateTime, Now()), "SHORT")
        jTime = Format$(Nz(rs!PrintDateTime, Now()), "hh:mm:ss")
        
        line = csv_EscapeField(CStr(Nz(rs!ReceiptID, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!EmployeeID, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!EmployeeFullName, ""))) & "," & _
               csv_EscapeField(CStr(Nz(rs!MealType, ""))) & "," & _
               csv_EscapeField(jDate) & "," & _
               csv_EscapeField(jTime)
        
        file.WriteLine line
        count = count + 1
        rs.MoveNext
    Loop
    
    file.Close
    
    Call LogSystemEvent("export_SuccessfulPrints", _
                       "گزارش چاپ‌های موفق صادر شد: " & CStr(count) & " رکورد به " & filePath)
    
    export_SuccessfulPrints = True
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set file = Nothing
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("export_SuccessfulPrints", Err.Number, Err.Description, filePath)
    export_SuccessfulPrints = False
    Resume CleanExit
End Function

' =========================================================
' تابع کمکی: csv_EscapeField
' =========================================================
'
' وظیفه:
' یک فیلد را برای استفاده در CSV آماده می‌کند
' اگر فیلد شامل ویرگول یا نقل‌قول باشد، آن را در نقل‌قول محصور می‌کند
' و نقل‌قول‌های داخلی را دوبل می‌کند
'
' پارامتر:
'   fieldValue (String): مقدار فیلد
'
' خروجی:
'   String: فیلد آماده‌شده برای CSV
'
' نمونه:
'   csv_EscapeField("علی") → "علی"
'   csv_EscapeField("علی، احمد") → """علی، احمد"""
'   csv_EscapeField("علی "" احمد") → """علی "" احمد"""
'
' =========================================================

Private Function csv_EscapeField(ByVal fieldValue As String) As String
    On Error GoTo ErrHandler
    
    fieldValue = Trim$(fieldValue)
    
    ' اگر شامل ویرگول، نقل‌قول، یا newline باشد
    If InStr(1, fieldValue, ",", vbTextCompare) > 0 Or _
       InStr(1, fieldValue, """", vbTextCompare) > 0 Or _
       InStr(1, fieldValue, vbCrLf, vbTextCompare) > 0 Then
        
        ' دوبل کردن نقل‌قول‌ها
        fieldValue = Replace(fieldValue, """", """""")
        
        ' محصور کردن در نقل‌قول
        csv_EscapeField = """" & fieldValue & """"
    Else
        csv_EscapeField = fieldValue
    End If
    
    Exit Function
    
ErrHandler:
    csv_EscapeField = fieldValue
End Function

' =========================================================
' تابع: export_DailyReport
' =========================================================
'
' وظیفه:
' گزارش روزانهٌ کامل را صادر می‌کند
' شامل تمام تردد‌ها، سفارشات، و چاپ‌های موفق
'
' پارامترها:
'   reportDate (Date): تاریخ گزارش
'   folderPath (String): مسیر پوشهٌ خروجی
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' فایل‌های ایجاد‌شده:
'   Attendance_1403_05_15.csv
'   MealOrders_1403_05_15.csv
'   SuccessfulPrints_1403_05_15.csv
'   DailyReport_1403_05_15.txt
'
' =========================================================

Public Function export_DailyReport(ByVal reportDate As Date, ByVal folderPath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim jDate As String
    Dim attendanceFile, mealFile, printsFile, summaryFile As String
    Dim fso As Object
    Dim file As Object
    Dim totalAttendance, totalMeals, totalPrints As Long
    Dim line As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' فرمت تاریخ
    jDate = format_jalali_date(reportDate, "SHORT")
    jDate = Replace(jDate, "/", "_")
    
    ' مسیرهای فایل‌ها
    attendanceFile = folderPath & "\Attendance_" & jDate & ".csv"
    mealFile = folderPath & "\MealOrders_" & jDate & ".csv"
    printsFile = folderPath & "\SuccessfulPrints_" & jDate & ".csv"
    summaryFile = folderPath & "\DailyReport_" & jDate & ".txt"
    
    ' صادر کردن گزارش‌ها
    If Not export_AttendanceReport(reportDate, reportDate, attendanceFile) Then
        Exit Function
    End If
    
    If Not export_MealOrdersReport(reportDate, reportDate, mealFile) Then
        Exit Function
    End If
    
    If Not export_SuccessfulPrints(reportDate, reportDate, printsFile) Then
        Exit Function
    End If
    
    ' ایجاد گزارش خلاصه
    Set file = fso.CreateTextFile(summaryFile, True, True)
    
    totalAttendance = db_GetTodayAttendanceCount()
    totalMeals = db_GetTodayMealOrderCount()
    totalPrints = db_GetSuccessfulRecordsCount()
    
    file.WriteLine "گزارش روزانهٌ توزیع غذا"
    file.WriteLine "=" & String(40, "=")
    file.WriteLine ""
    
    file.WriteLine "تاریخ: " & format_jalali_date(reportDate, "FULL")
    file.WriteLine ""
    
    file.WriteLine "خلاصهٌ گزارش:"
    file.WriteLine "-" & String(40, "-")
    file.WriteLine "کل تردد‌ها: " & CStr(totalAttendance)
    file.WriteLine "کل سفارش��ت غذا: " & CStr(totalMeals)
    file.WriteLine "چاپ‌های موفق: " & CStr(totalPrints)
    file.WriteLine ""
    
    file.WriteLine "فایل‌های تفصیلی:"
    file.WriteLine "1. " & attendanceFile
    file.WriteLine "2. " & mealFile
    file.WriteLine "3. " & printsFile
    file.WriteLine ""
    
    file.WriteLine "زمان تولید: " & Now()
    
    file.Close
    
    Call LogSystemEvent("export_DailyReport", _
                       "گزارش روزانه صادر شد: " & folderPath)
    
    export_DailyReport = True
    
CleanExit:
    Set file = Nothing
    Set fso = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("export_DailyReport", Err.Number, Err.Description, folderPath)
    export_DailyReport = False
    Resume CleanExit
End Function

' =========================================================
' تابع کمکی: db_GetTodayAttendanceCount
' =========================================================

Private Function db_GetTodayAttendanceCount() As Long
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim count As Long
    
    Set db = CurrentDb()
    
    sql = "SELECT COUNT(*) as cnt FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE DateValue(AttendanceDateTime)=DateValue(Now())"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        count = Nz(rs!cnt, 0)
    End If
    
    db_GetTodayAttendanceCount = count
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    db_GetTodayAttendanceCount = 0
End Function

' =========================================================
' تابع کمکی: db_GetTodayMealOrderCount
' =========================================================

Private Function db_GetTodayMealOrderCount() As Long
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim count As Long
    
    Set db = CurrentDb()
    
    sql = "SELECT COUNT(DISTINCT EmployeeID) as cnt FROM " & TABLE_MEAL_ORDERS & " " & _
          "WHERE DateValue(ReserveDate)=DateValue(Now())"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        count = Nz(rs!cnt, 0)
    End If
    
    db_GetTodayMealOrderCount = count
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    db_GetTodayMealOrderCount = 0
End Function

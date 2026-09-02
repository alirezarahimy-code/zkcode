Option Compare Database
Option Explicit

' =========================================================
' فرم: frmPrinterSettings
' =========================================================
'
' توضیح فرم:
' این فرم تنظیمات چاپگر را مدیریت می‌کند.
' شامل IP، پورت، نام چاپگر Windows و غیره.
'
' کاربرد:
' - تنظیم IP چاپگر شبکه‌ای
' - تنظیم پورت (معمولاً 9100)
' - انتخاب چاپگر Windows
' - تست اتصال
' - ذخیره تنظیمات
'
' کنترل‌ها:
' - txtPrinterIP: آدرس IP
' - txtPrinterPort: پورت
' - cmbPrinterName: نام چاپگر
' - cmdTest: تست اتصال
' - cmdSave: ذخیره
' - cmdClose: بستن
'
' رویدادها:
' - Form_Load: بارگزاری
' - cmdTest_Click: تست
' - cmdSave_Click: ذخیره
'
' =========================================================

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Me.Caption = "تنظیمات چاپگر"
    
    ' بارگزاری تنظیمات موجود
    Call LoadPrinterSettings
    
    ' بارگزاری لیست چاپگرهای Windows
    Call LoadWindowsPrinters
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در بارگزاری فرم: " & Err.Description, vbCritical
End Sub

' =========================================================
' زیربرنامه: LoadPrinterSettings
' =========================================================
'
' وظیفه:
' تنظیمات چاپگر ذخیره‌شده را بارگزاری می‌کند
'
' =========================================================

Private Sub LoadPrinterSettings()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb()
    
    ' جدول تنظیمات وجود دارد؟
    If Not TableExists("tblPrinterSettings") Then
        ' تنظیمات پیش‌فرض
        Me.txtPrinterIP.Value = "192.168.1.50"
        Me.txtPrinterPort.Value = 9100
        Me.cmbPrinterName.Value = ""
        Exit Sub
    End If
    
    sql = "SELECT TOP 1 PrinterIP, PrinterPort, PrinterName FROM tblPrinterSettings"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        Me.txtPrinterIP.Value = Nz(rs!PrinterIP, "192.168.1.50")
        Me.txtPrinterPort.Value = Nz(rs!PrinterPort, 9100)
        Me.cmbPrinterName.Value = Nz(rs!PrinterName, "")
    Else
        ' پیش‌فرض
        Me.txtPrinterIP.Value = "192.168.1.50"
        Me.txtPrinterPort.Value = 9100
        Me.cmbPrinterName.Value = ""
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    Me.txtPrinterIP.Value = "192.168.1.50"
    Me.txtPrinterPort.Value = 9100
    Resume CleanExit
End Sub

' =========================================================
' زیربرنامه: LoadWindowsPrinters
' =========================================================
'
' وظیفه:
' لیست چاپگرهای Windows را بارگزاری می‌کند
'
' =========================================================

Private Sub LoadWindowsPrinters()
    On Error GoTo ErrHandler
    
    ' پاک کردن لیست
    If Me.cmbPrinterName.ListCount > 0 Then
        Me.cmbPrinterName.Clear
    End If
    
    ' اضافه کردن گزینهٌ خالی
    Me.cmbPrinterName.AddItem ""
    
    ' توجه: برای دریافت لیست چاپگرهای Windows
    ' نیاز به WMI یا Windows API است
    ' برای حالا: چاپگرهای معمول را اضافه می‌کنیم
    
    ' مثال‌های معمول
    Me.cmbPrinterName.AddItem "HP LaserJet"
    Me.cmbPrinterName.AddItem "Canon Printer"
    Me.cmbPrinterName.AddItem "Xerox"
    Me.cmbPrinterName.AddItem "Epson"
    Me.cmbPrinterName.AddItem "Brother"
    
    Exit Sub
    
ErrHandler:
    ' خطا در بارگزاری
    Me.cmbPrinterName.AddItem ""
End Sub

' =========================================================
' رویداد: cmdTest_Click
' =========================================================

Private Sub cmdTest_Click()
    On Error GoTo ErrHandler
    
    Dim printerIP As String
    Dim printerPort As Long
    
    printerIP = Trim$(Me.txtPrinterIP.Value)
    printerPort = CLng(Nz(Me.txtPrinterPort.Value, 9100))
    
    If Len(printerIP) = 0 Then
        MsgBox "لطفاً IP چاپگر را وارد کنید", vbExclamation
        Me.txtPrinterIP.SetFocus
        Exit Sub
    End If
    
    If printerPort <= 0 Then
        MsgBox "پورت نامعتبر است", vbExclamation
        Me.txtPrinterPort.SetFocus
        Exit Sub
    End If
    
    ' تست اتصال
    Call LogSystemEvent("frmPrinterSettings", "در حال تست چاپگر: " & printerIP & ":" & CStr(printerPort))
    
    If ph_TestConnection(printerIP, printerPort) Then
        MsgBox "✓ چاپگر پاسخ می‌دهد", vbInformation
    Else
        MsgBox "✗ چاپگر پاسخ نمی‌دهد" & vbCrLf & "خطا: " & ph_GetLastError(), vbCritical
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

' =========================================================
' رویداد: cmdSave_Click
' =========================================================

Private Sub cmdSave_Click()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim printerIP As String
    Dim printerPort As Long
    Dim printerName As String
    
    printerIP = Trim$(Me.txtPrinterIP.Value)
    printerPort = CLng(Nz(Me.txtPrinterPort.Value, 9100))
    printerName = Trim$(Me.cmbPrinterName.Value)
    
    ' بررسی IP
    If Len(printerIP) = 0 Then
        MsgBox "لطفاً IP را وارد کنید", vbExclamation
        Me.txtPrinterIP.SetFocus
        Exit Sub
    End If
    
    If printerPort <= 0 Then
        MsgBox "پورت نامعتبر است", vbExclamation
        Me.txtPrinterPort.SetFocus
        Exit Sub
    End If
    
    Set db = CurrentDb()
    
    ' جدول تنظیمات موجود است؟
    If Not TableExists("tblPrinterSettings") Then
        ' ایجاد جدول
        Call CreatePrinterSettingsTable(db)
    End If
    
    ' حذف تنظیمات قدیم
    sql = "DELETE FROM tblPrinterSettings"
    db.Execute sql
    
    ' درج تنظیمات جدید
    sql = "INSERT INTO tblPrinterSettings (PrinterIP, PrinterPort, PrinterName) " & _
          "VALUES ('" & printerIP & "', " & CStr(printerPort) & ", '" & printerName & "')"
    
    db.Execute sql
    
    MsgBox "تنظیمات ذخیره شد", vbInformation
    
    Call LogSystemEvent("frmPrinterSettings", _
                       "تنظیمات چاپگر ذخیره شد: " & printerIP & ":" & CStr(printerPort))
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در ذخیره: " & Err.Description, vbCritical
End Sub

' =========================================================
' زیربرنامه: CreatePrinterSettingsTable
' =========================================================

Private Sub CreatePrinterSettingsTable(ByVal db As DAO.Database)
    On Error GoTo ErrHandler
    
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    
    Set td = db.CreateTableDef("tblPrinterSettings")
    
    ' SettingID
    Set fld = td.CreateField("SettingID", dbAutoIncrement)
    fld.Attributes = dbAutoIncrField
    td.Fields.Append fld
    
    ' PrinterIP
    Set fld = td.CreateField("PrinterIP", dbText, 50)
    td.Fields.Append fld
    
    ' PrinterPort
    Set fld = td.CreateField("PrinterPort", dbLong)
    td.Fields.Append fld
    
    ' PrinterName
    Set fld = td.CreateField("PrinterName", dbText, 255)
    td.Fields.Append fld
    
    db.TableDefs.Append td
    
    Exit Sub
    
ErrHandler:
    ' جدول شاید قبلاً موجود است
End Sub

' =========================================================
' رویداد: cmdClose_Click
' =========================================================

Private Sub cmdClose_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name
End Sub

' =========================================================
' تابع کمکی: TableExists
' =========================================================

Private Function TableExists(ByVal tableName As String) As Boolean
    On Error GoTo EH
    Dim td As DAO.TableDef
    For Each td In CurrentDb().TableDefs
        If StrComp(td.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next td
    Exit Function
EH:
    TableExists = False
End Function

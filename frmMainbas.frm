' فرم اصلی برنامه - توسط Form Designer ایجاد شود
' این کد برای Private Module فرم است

Option Compare Database
Option Explicit

' =========================================================
' فرم: frmMain
' =========================================================
' 
' توضیح:
' فرم اصلی برنامه است که کاربر اولین بار با آن مواجه می‌شود.
' شامل:
'   - دکمه شروع/توقف مانیتور
'   - نمایش وضعیت
'   - دکمه‌های دسترسی به سایر فرم‌ها
'   - نمایش لاگ‌های سیستم
'
' ویژگی‌ها:
' - Timer برای بروزرسانی وضعیت هر 5 ثانیه
' - نمایش رنگی برای وضعیت
' - Quick Links به فرم‌های دیگر
'
' =========================================================

Dim timerID As Long

Private Sub Form_Load()
    On Error GoTo EH
    
    ' تنظیمات فرم
    Me.Caption = "سیستم توزیع غذا با دستگاه حضور و غیاب"
    Me.Width = 9000
    Me.Height = 6000
    
    ' ایجاد Controls
    Call CreateControls()
    
    ' راه‌اندازی دستگاه
    If db_migration_CreateOrUpdateSchema() Then
        MsgBox "پایگاه‌داده آماده شد", vbInformation
    Else
        MsgBox "خطا در آماده‌سازی پایگاه‌داده", vbCritical
        Exit Sub
    End If
    
    ' شروع مانیتور
    Call monitor_Start()
    
    ' شروع Timer
    timerID = SetTimer(Me.hWnd, 1, 5000, AddressOf TimerProc)
    
    Call UpdateStatus()
    
    Exit Sub
EH:
    MsgBox "خطا در بارگزاری: " & Err.Description, vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' توقف Timer
    If timerID <> 0 Then
        Call KillTimer(Me.hWnd, timerID)
    End If
    
    ' توقف مانیتور
    Call monitor_Stop()
    
End Sub

Private Sub CreateControls()
    On Error GoTo EH
    
    ' عنوان
    Dim lblTitle As Control
    Set lblTitle = Me.Controls.Add("Forms.Label.1")
    With lblTitle
        .Name = "lblTitle"
        .Caption = "🖥️ سیستم توزیع غذا با دستگاه حضور و غیاب"
        .Left = 100
        .Top = 100
        .Width = 8800
        .Height = 400
        .FontSize = 16
        .FontBold = True
        .ForeColor = 2070783 ' آبی
    End With
    
    ' دکمه شروع/توقف
    Dim cmdStartStop As Control
    Set cmdStartStop = Me.Controls.Add("Forms.CommandButton.1")
    With cmdStartStop
        .Name = "cmdStartStop"
        .Caption = "▶️ شروع مانیتور"
        .Left = 100
        .Top = 600
        .Width = 2000
        .Height = 500
        .FontSize = 11
        .FontBold = True
        .BackColor = 5296274 ' سبز
        .ForeColor = 16777215 ' سفید
    End With
    
    ' Label وضعیت
    Dim lblStatus As Control
    Set lblStatus = Me.Controls.Add("Forms.Label.1")
    With lblStatus
        .Name = "lblStatus"
        .Caption = "⚫ متوقف"
        .Left = 2200
        .Top = 600
        .Width = 2000
        .Height = 500
        .FontSize = 11
        .FontBold = True
    End With
    
    ' لیست دستگاه‌ها
    Dim lstDevices As Control
    Set lstDevices = Me.Controls.Add("Forms.ListBox.1")
    With lstDevices
        .Name = "lstDevices"
        .Left = 100
        .Top = 1200
        .Width = 4000
        .Height = 2500
        .ColumnHeads = True
    End With
    
    ' دکمه تنظیمات
    Dim cmdSettings As Control
    Set cmdSettings = Me.Controls.Add("Forms.CommandButton.1")
    With cmdSettings
        .Name = "cmdSettings"
        .Caption = "⚙️ تنظیمات"
        .Left = 4200
        .Top = 1200
        .Width = 1800
        .Height = 400
        .FontSize = 10
    End With
    
    ' دکمه مانیتور آنی
    Dim cmdMonitoring As Control
    Set cmdMonitoring = Me.Controls.Add("Forms.CommandButton.1")
    With cmdMonitoring
        .Name = "cmdMonitoring"
        .Caption = "📊 پایش آنی"
        .Left = 4200
        .Top = 1700
        .Width = 1800
        .Height = 400
        .FontSize = 10
    End With
    
    ' دکمه گزارش‌ها
    Dim cmdReports As Control
    Set cmdReports = Me.Controls.Add("Forms.CommandButton.1")
    With cmdReports
        .Name = "cmdReports"
        .Caption = "📈 گزارش‌ها"
        .Left = 4200
        .Top = 2200
        .Width = 1800
        .Height = 400
        .FontSize = 10
    End With
    
    ' دکمه کارمندان
    Dim cmdEmployees As Control
    Set cmdEmployees = Me.Controls.Add("Forms.CommandButton.1")
    With cmdEmployees
        .Name = "cmdEmployees"
        .Caption = "👥 کارمندان"
        .Left = 4200
        .Top = 2700
        .Width = 1800
        .Height = 400
        .FontSize = 10
    End With
    
    Exit Sub
EH:
    MsgBox "خطا در ایجاد کنترل‌ها: " & Err.Description, vbCritical
End Sub

Private Sub UpdateStatus()
    On Error Resume Next
    
    If MonitorRunning Then
        Me.lblStatus.Caption = "🟢 درحال اجرا"
        Me.cmdStartStop.Caption = "⏹️ توقف مانیتور"
        Me.cmdStartStop.BackColor = 255 ' قرمز
    Else
        Me.lblStatus.Caption = "⚫ متوقف"
        Me.cmdStartStop.Caption = "▶️ شروع مانیتور"
        Me.cmdStartStop.BackColor = 5296274 ' سبز
    End If
    
    ' بروزرسانی لیست دستگاه‌ها
    Call RefreshDevicesList()
    
End Sub

Private Sub RefreshDevicesList()
    On Error GoTo EH
    
    Dim rs As DAO.Recordset
    Dim db As DAO.Database
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT DeviceName, DeviceIP, DevicePort, IsActive FROM " & TABLE_ZK_DEVICES & " ORDER BY DeviceIP")
    
    Me.lstDevices.RowSource = ""
    Me.lstDevices.ColumnCount = 4
    Me.lstDevices.ColumnWidths = "2cm;3cm;1.5cm;1cm"
    
    Dim sql As String
    sql = "SELECT DeviceName, DeviceIP & ':' & DevicePort AS Address, IsActive FROM " & TABLE_ZK_DEVICES & " ORDER BY DeviceIP"
    Me.lstDevices.RowSource = sql
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

' رویدادهای دکمه‌ها
Private Sub cmdStartStop_Click()
    On Error GoTo EH
    
    If MonitorRunning Then
        Call monitor_Stop()
    Else
        Call monitor_Start()
    End If
    
    Call UpdateStatus()
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub cmdSettings_Click()
    On Error GoTo EH
    DoCmd.OpenForm "frmDeviceSetup"
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub cmdMonitoring_Click()
    On Error GoTo EH
    DoCmd.OpenForm "frmLiveMonitoring"
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub cmdReports_Click()
    On Error GoTo EH
    DoCmd.OpenForm "frmReports"
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub cmdEmployees_Click()
    On Error GoTo EH
    DoCmd.OpenForm "frmEmployees"
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    
    ' اجرای Tick مانیتور
    Call monitor_Tick()
    
    ' بروزرسانی وضعیت
    Call UpdateStatus()
    
End Sub
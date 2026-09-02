Option Compare Database
Option Explicit

' =========================================================
' فرم: frmDeviceSetup
' =========================================================
' 
' توضیح:
' مدیریت دستگاه‌های ZK
' اضافه کردن، ویرایش، حذف، تست اتصال
'
' =========================================================

Private Sub Form_Load()
    On Error GoTo EH
    
    Me.Caption = "معرفی دستگاه‌های ZK"
    Me.Width = 8000
    Me.Height = 5000
    
    Call CreateControls()
    Call RefreshDevicesList()
    
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub CreateControls()
    ' لیست دستگاه‌ها
    Dim lstDevices As Control
    Set lstDevices = Me.Controls.Add("Forms.ListBox.1")
    With lstDevices
        .Name = "lstDevices"
        .Left = 100
        .Top = 100
        .Width = 7800
        .Height = 2500
        .ColumnHeads = True
    End With
    
    ' فیلدهای ورودی
    Dim lblIP As Control
    Set lblIP = Me.Controls.Add("Forms.Label.1")
    With lblIP
        .Caption = "IP Address:"
        .Left = 100
        .Top = 2700
        .Width = 1500
    End With
    
    Dim txtIP As Control
    Set txtIP = Me.Controls.Add("Forms.TextBox.1")
    With txtIP
        .Name = "txtIP"
        .Left = 1700
        .Top = 2700
        .Width = 2000
    End With
    
    Dim lblPort As Control
    Set lblPort = Me.Controls.Add("Forms.Label.1")
    With lblPort
        .Caption = "Port:"
        .Left = 100
        .Top = 3200
        .Width = 1500
    End With
    
    Dim txtPort As Control
    Set txtPort = Me.Controls.Add("Forms.TextBox.1")
    With txtPort
        .Name = "txtPort"
        .Left = 1700
        .Top = 3200
        .Width = 2000
        .Value = "4370"
    End With
    
    ' دکمه‌ها
    Dim cmdAdd As Control
    Set cmdAdd = Me.Controls.Add("Forms.CommandButton.1")
    With cmdAdd
        .Name = "cmdAdd"
        .Caption = "➕ اضافه"
        .Left = 100
        .Top = 3800
        .Width = 1500
    End With
    
    Dim cmdTest As Control
    Set cmdTest = Me.Controls.Add("Forms.CommandButton.1")
    With cmdTest
        .Name = "cmdTest"
        .Caption = "🔌 تست"
        .Left = 1700
        .Top = 3800
        .Width = 1500
    End With
    
    Dim cmdDelete As Control
    Set cmdDelete = Me.Controls.Add("Forms.CommandButton.1")
    With cmdDelete
        .Name = "cmdDelete"
        .Caption = "🗑️ حذف"
        .Left = 3300
        .Top = 3800
        .Width = 1500
    End With
    
    Dim cmdClose As Control
    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1")
    With cmdClose
        .Name = "cmdClose"
        .Caption = "❌ بستن"
        .Left = 6300
        .Top = 3800
        .Width = 1500
    End With
    
End Sub

Private Sub RefreshDevicesList()
    On Error Resume Next
    
    Dim sql As String
    sql = "SELECT DeviceName, DeviceIP & ':' & DevicePort AS Address, IsActive FROM " & TABLE_ZK_DEVICES & " ORDER BY DeviceIP"
    
    Me.lstDevices.RowSource = sql
    Me.lstDevices.ColumnCount = 3
    Me.lstDevices.ColumnWidths = "3cm;3cm;2cm"
    
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo EH
    
    Dim ip As String, port As Long
    ip = Trim$(Me.txtIP.Value)
    port = Val(Me.txtPort.Value)
    
    If ip = "" Or port <= 0 Then
        MsgBox "IP و Port را بررسی کنید", vbExclamation
        Exit Sub
    End If
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset(TABLE_ZK_DEVICES, dbOpenDynaset)
    
    rs.AddNew
    rs!DeviceName = "Device_" & ip
    rs!DeviceIP = ip
    rs!DevicePort = port
    rs!MachineNumber = 1
    rs!CommKey = 0
    rs!IsActive = True
    rs!CreatedDate = Now()
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    MsgBox "دستگاه اضافه شد", vbInformation
    Call RefreshDevicesList()
    
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub cmdTest_Click()
    On Error GoTo EH
    
    Dim ip As String, port As Long
    ip = Trim$(Me.txtIP.Value)
    port = Val(Me.txtPort.Value)
    
    If ip = "" Then
        MsgBox "IP را وارد کنید", vbExclamation
        Exit Sub
    End If
    
    If port <= 0 Then port = DEFAULT_ZK_PORT
    
    Dim zk As Object
    Set zk = zk_Connect(ip, port, 0)
    
    If Not zk Is Nothing Then
        MsgBox "✓ اتصال موفق", vbInformation
        zk.Disconnect
    Else
        MsgBox "✗ اتصال ناموفق", vbCritical
    End If
    
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo EH
    
    If Me.lstDevices.ListIndex < 0 Then
        MsgBox "دستگاهی را انتخاب کنید", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("حذف می‌کنید؟", vbYesNo) <> vbYes Then Exit Sub
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' حذف (از DB بخوان و حذف کن)
    ' ...کد حذف...
    
    Call RefreshDevicesList()
    Exit Sub
    
EH:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close
End Sub
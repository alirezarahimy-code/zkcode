Option Compare Database
Option Explicit

Dim timerID As Long

Private Sub Form_Load()
    Me.Caption = "📊 پایش آنی تردد‌ها"
    Me.Width = 10000
    Me.Height = 5000
    
    Call CreateControls()
    
    ' شروع Timer برای Refresh هر 2 ثانیه
    timerID = SetTimer(Me.hWnd, 1, 2000, AddressOf TimerProc)
    
    Call RefreshData()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If timerID <> 0 Then
        Call KillTimer(Me.hWnd, timerID)
    End If
End Sub

Private Sub CreateControls()
    ' لیست تردد‌ها
    Dim lstMonitoring As Control
    Set lstMonitoring = Me.Controls.Add("Forms.ListBox.1")
    With lstMonitoring
        .Name = "lstMonitoring"
        .Left = 100
        .Top = 100
        .Width = 9800
        .Height = 4200
        .ColumnHeads = True
        .ColumnCount = 5
        .ColumnWidths = "2cm;3cm;3cm;2cm;2cm"
    End With
    
    ' دکمه بروزرسانی
    Dim cmdRefresh As Control
    Set cmdRefresh = Me.Controls.Add("Forms.CommandButton.1")
    With cmdRefresh
        .Name = "cmdRefresh"
        .Caption = "🔄 بروزرسانی"
        .Left = 100
        .Top = 4400
        .Width = 1500
    End With
    
    ' دکمه بستن
    Dim cmdClose As Control
    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1")
    With cmdClose
        .Name = "cmdClose"
        .Caption = "❌ بستن"
        .Left = 8400
        .Top = 4400
        .Width = 1500
    End With
End Sub

Private Sub RefreshData()
    On Error Resume Next
    
    Dim sql As String
    sql = "SELECT TOP 100 " & _
          "Format(LogDateTime,'hh:mm:ss') AS [ساعت], " & _
          "EnrollID AS [کد], " & _
          "Details AS [جزئیات], " & _
          "Result AS [نتیجه], " & _
          "Format(LogDateTime,'yyyy/mm/dd') AS [تاریخ] " & _
          "FROM " & TABLE_LIVE_MONITORING & " " & _
          "ORDER BY LogDateTime DESC"
    
    Me.lstMonitoring.RowSource = sql
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshData()
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close
End Sub

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    Call RefreshData()
End Sub
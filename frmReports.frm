Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.Caption = "📈 گزارش‌های سیستم"
    Me.Width = 10000
    Me.Height = 6000
    
    Call CreateControls()
    Call RefreshData()
End Sub

Private Sub CreateControls()
    ' فیلترها
    Dim lblDate As Control
    Set lblDate = Me.Controls.Add("Forms.Label.1")
    With lblDate
        .Caption = "تاریخ:"
        .Left = 100
        .Top = 100
        .Width = 1000
    End With
    
    Dim txtDate As Control
    Set txtDate = Me.Controls.Add("Forms.TextBox.1")
    With txtDate
        .Name = "txtDate"
        .Left = 1200
        .Top = 100
        .Width = 2000
        .Value = Format(Now(), "yyyy/mm/dd")
    End With
    
    ' لیست نتایج
    Dim lstResults As Control
    Set lstResults = Me.Controls.Add("Forms.ListBox.1")
    With lstResults
        .Name = "lstResults"
        .Left = 100
        .Top = 600
        .Width = 9800
        .Height = 4500
        .ColumnHeads = True
        .ColumnCount = 6
        .ColumnWidths = "2cm;2cm;2cm;2cm;2cm;2cm"
    End With
    
    ' خلاصه
    Dim lblSummary As Control
    Set lblSummary = Me.Controls.Add("Forms.Label.1")
    With lblSummary
        .Name = "lblSummary"
        .Left = 100
        .Top = 5200
        .Width = 9800
        .Height = 600
        .BorderStyle = 1
    End With
    
End Sub

Private Sub RefreshData()
    On Error GoTo EH
    
    Dim sql As String
    sql = "SELECT " & _
          "EnrollID, FullName, AttendanceDateTime, ProcessingResult, PrintAttempts, " & _
          "IIf(ProcessingResult='PRINTED','✓','✗') AS Status " & _
          "FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE Int(AttendanceDateTime)=" & SqlDateTime(CDate(Me.txtDate.Value)) & " " & _
          "ORDER BY AttendanceDateTime"
    
    Me.lstResults.RowSource = sql
    Me.lstResults.ColumnCount = 6
    Me.lstResults.ColumnWidths = "1.5cm;3cm;2.5cm;2.5cm;1.5cm;1cm"
    
    ' خلاصه
    Dim dbSummary As String
    dbSummary = "✓ موفق: 0 | ✗ ناموفق: 0 | ⏳ درانتظار: 0"
    Me.lblSummary.Caption = dbSummary
    
    Exit Sub
EH:
    MsgBox "خطا: " & Err.Description
End Sub
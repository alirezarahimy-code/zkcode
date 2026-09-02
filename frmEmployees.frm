Option Compare Database
Option Explicit

Private Sub Form_Load()
    Me.Caption = "👥 لیست کارمندان"
    Me.Width = 8000
    Me.Height = 5000
    
    Call CreateControls()
    Call RefreshData()
End Sub

Private Sub CreateControls()
    ' لیست کارمندان
    Dim lstEmployees As Control
    Set lstEmployees = Me.Controls.Add("Forms.ListBox.1")
    With lstEmployees
        .Name = "lstEmployees"
        .Left = 100
        .Top = 100
        .Width = 7800
        .Height = 3500
        .ColumnHeads = True
        .ColumnCount = 5
        .ColumnWidths = "1.5cm;2cm;2cm;2cm;1.5cm"
    End With
    
    ' دکمه بستن
    Dim cmdClose As Control
    Set cmdClose = Me.Controls.Add("Forms.CommandButton.1")
    With cmdClose
        .Caption = "❌ بستن"
        .Left = 6300
        .Top = 3700
        .Width = 1600
    End With
End Sub

Private Sub RefreshData()
    On Error Resume Next
    
    Dim sql As String
    sql = "SELECT EmployeeID, FullName, NationalCode, " & _
          "IIf(IsActive,True,False) AS [فعال] FROM " & TABLE_EMPLOYEES & " " & _
          "ORDER BY FullName"
    
    Me.lstEmployees.RowSource = sql
End Sub

Private Sub cmdClose_Click()
    DoCmd.Close
End Sub
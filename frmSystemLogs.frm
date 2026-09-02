Option Compare Database
Option Explicit

' =========================================================
' فرم: frmSystemLogs
' =========================================================
'
' توضیح فرم:
' این فرم لاگ‌های سیستمی را نمایش می‌دهد.
' شامل تمام رویدادها، اتصالات، و عملیات سیستم.
'
' کاربرد:
' - نمایش لاگ‌های سیستمی
' - فیلتر بر اساس تاریخ
' - جستجو در لاگ‌ها
' - صادرات لاگ‌ها
'
' کنترل‌ها:
' - lstSystemLogs: لیستی از لاگ‌ها
' - txtSearch: جستجو
' - dtFromDate, dtToDate: فیلتر تاریخ
' - cmdSearch, cmdRefresh, cmdExport, cmdClose
'
' رویدادها:
' - Form_Load: بارگزاری
' - cmdSearch_Click: جستجو
' - cmdRefresh_Click: بروزرسانی
' - cmdExport_Click: صادرات
'
' =========================================================

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Me.Caption = "لاگ‌های سیستم"
    
    ' تنظیم تاریخ‌های پیش‌فرض
    Me.dtFromDate.Value = DateAdd("d", -7, Now())  ' 7 روز گذشته
    Me.dtToDate.Value = Now()
    
    ' بارگزاری لاگ‌ها
    Call RefreshSystemLogs
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در بارگزاری فرم: " & Err.Description, vbCritical
End Sub

' =========================================================
' زیربرنامه: RefreshSystemLogs
' =========================================================
'
' وظیفه:
' لاگ‌های سیستمی را بروزرسانی می‌کند
'
' =========================================================

Private Sub RefreshSystemLogs()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fromDate As Date
    Dim toDate As Date
    Dim count As Long
    
    Set db = CurrentDb()
    
    fromDate = Me.dtFromDate.Value
    toDate = Me.dtToDate.Value + 1
    
    ' بررسی جدول وجود دارد
    If Not TableExists("tblSystemLogs") Then
        Me.lblCount.Caption = "لاگی موجود نیست"
        Exit Sub
    End If
    
    ' جستجو
    sql = "SELECT LogID, LogDate, Source, Message FROM tblSystemLogs " & _
          "WHERE LogDate >= " & SqlDateTime(fromDate) & " " & _
          "AND LogDate < " & SqlDateTime(toDate) & " " & _
          "ORDER BY LogDate DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' پاک کردن ListBox
    While Me.lstSystemLogs.ListCount > 0
        Me.lstSystemLogs.RemoveItem 0
    Wend
    
    count = 0
    Do While Not rs.EOF
        Dim line As String
        Dim logDate As String
        
        logDate = Format$(rs!LogDate, "yyyy/mm/dd hh:mm:ss")
        line = "[" & logDate & "] " & Nz(rs!Source, "") & ": " & Nz(rs!Message, "")
        
        Me.lstSystemLogs.AddItem line
        count = count + 1
        rs.MoveNext
    Loop
    
    Me.lblCount.Caption = "کل: " & CStr(count) & " لاگ"
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در بروزرسانی لاگ‌ها: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' =========================================================
' رویداد: cmdSearch_Click
' =========================================================

Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim searchText As String
    Dim fromDate As Date
    Dim toDate As Date
    Dim count As Long
    
    Set db = CurrentDb()
    
    searchText = Trim$(Me.txtSearch.Value)
    fromDate = Me.dtFromDate.Value
    toDate = Me.dtToDate.Value + 1
    
    If Len(searchText) = 0 Then
        MsgBox "لطفاً متن جستجو را وارد کنید", vbExclamation
        Me.txtSearch.SetFocus
        Exit Sub
    End If
    
    If Not TableExists("tblSystemLogs") Then
        MsgBox "لاگی موجود نیست", vbInformation
        Exit Sub
    End If
    
    ' جستجو با LIKE
    sql = "SELECT LogID, LogDate, Source, Message FROM tblSystemLogs " & _
          "WHERE (Source LIKE '%" & Replace(searchText, "'", "''") & "%' " & _
          "OR Message LIKE '%" & Replace(searchText, "'", "''") & "%') " & _
          "AND LogDate >= " & SqlDateTime(fromDate) & " " & _
          "AND LogDate < " & SqlDateTime(toDate) & " " & _
          "ORDER BY LogDate DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' پاک کردن ListBox
    While Me.lstSystemLogs.ListCount > 0
        Me.lstSystemLogs.RemoveItem 0
    Wend
    
    count = 0
    Do While Not rs.EOF
        Dim line As String
        Dim logDate As String
        
        logDate = Format$(rs!LogDate, "yyyy/mm/dd hh:mm:ss")
        line = "[" & logDate & "] " & Nz(rs!Source, "") & ": " & Nz(rs!Message, "")
        
        Me.lstSystemLogs.AddItem line
        count = count + 1
        rs.MoveNext
    Loop
    
    Me.lblCount.Caption = "نتیجهٌ جستجو: " & CStr(count) & " لاگ"
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در جستجو: " & Err.Description, vbCritical
End Sub

' =========================================================
' رویداد: cmdRefresh_Click
' =========================================================

Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    
    Call RefreshSystemLogs
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

' =========================================================
' رویداد: cmdExport_Click
' =========================================================

Private Sub cmdExport_Click()
    On Error GoTo ErrHandler
    
    Dim folderPath As String
    Dim fileName As String
    Dim fso As Object
    Dim file As Object
    Dim count As Long
    
    ' درخواست مسیر
    folderPath = GetFolderPath()
    
    If Len(folderPath) = 0 Then
        Exit Sub
    End If
    
    ' ایجاد نام فایل
    fileName = folderPath & "\SystemLogs_" & Format$(Now(), "yyyymmdd_hhmmss") & ".txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(fileName, True, True)
    
    ' نوشتن هدر
    file.WriteLine "لاگ‌های سیستم"
    file.WriteLine "=" & String(60, "=")
    file.WriteLine ""
    file.WriteLine "تاریخ صادرات: " & Now()
    file.WriteLine ""
    
    ' نوشتن لاگ‌ها
    Dim i As Long
    For i = 0 To Me.lstSystemLogs.ListCount - 1
        file.WriteLine Me.lstSystemLogs.List(i)
        count = count + 1
    Next i
    
    file.Close
    
    MsgBox "لاگ‌ها صادر شدند" & vbCrLf & "مسیر: " & fileName, vbInformation
    
CleanExit:
    Set file = Nothing
    Set fso = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در صادرات: " & Err.Description, vbCritical
End Sub

' =========================================================
' رویداد: cmdClose_Click
' =========================================================

Private Sub cmdClose_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name
End Sub

' =========================================================
' تابع کمکی: GetFolderPath
' =========================================================

Private Function GetFolderPath() As String
    On Error GoTo ErrHandler
    
    Dim shell As Object
    Dim folder As Object
    
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, "مسیری برای ذخیره لاگ‌ها را انتخاب کنید:", 0, 0)
    
    If Not folder Is Nothing Then
        GetFolderPath = folder.Self.Path
    End If
    
    Set folder = Nothing
    Set shell = Nothing
    
    Exit Function
    
ErrHandler:
    GetFolderPath = ""
End Function

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

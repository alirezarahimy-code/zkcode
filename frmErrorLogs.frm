Option Compare Database
Option Explicit

' =========================================================
' فرم: frmErrorLogs
' =========================================================
'
' توضیح فرم:
' این فرم لاگ‌های خطا را نمایش می‌دهد.
' شامل تمام خطاهای برنامه با جزئیات کامل.
'
' کاربرد:
' - نمایش خطاهای برنامه
' - فیلتر بر اساس شماره خطا یا تابع
' - جستجو در خطاها
' - صادرات گزارش خطاها
' - تشخیص مشکلات
'
' کنترل‌ها:
' - lstErrorLogs: لیستی از خطاها
' - txtErrorNumber: فیلتر شماره خطا
' - txtProcedure: فیلتر نام تابع
' - dtFromDate, dtToDate: فیلتر تاریخ
' - cmdFilter, cmdClear, cmdExport, cmdClose
'
' رویدادها:
' - Form_Load: بارگزاری
' - cmdFilter_Click: اعمال فیلتر
' - cmdClear_Click: پاک کردن فیلتر
' - cmdExport_Click: صادرات
'
' =========================================================

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Me.Caption = "لاگ‌های خطا"
    
    ' تنظیم تاریخ‌های پیش‌فرض
    Me.dtFromDate.Value = DateAdd("d", -7, Now())  ' 7 روز گذشته
    Me.dtToDate.Value = Now()
    
    ' بارگزاری خطاها
    Call RefreshErrorLogs
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در بارگزاری فرم: " & Err.Description, vbCritical
End Sub

' =========================================================
' زیربرنامه: RefreshErrorLogs
' =========================================================
'
' وظیفه:
' لاگ‌های خطا را بروزرسانی می‌کند
'
' =========================================================

Private Sub RefreshErrorLogs()
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
    If Not TableExists("tblErrorLogs") Then
        Me.lblCount.Caption = "خطایی ثبت نشده"
        Exit Sub
    End If
    
    ' جستجو
    sql = "SELECT ErrorID, ErrorDate, ErrNumber, ErrDescription, Procedure " & _
          "FROM tblErrorLogs " & _
          "WHERE ErrorDate >= " & SqlDateTime(fromDate) & " " & _
          "AND ErrorDate < " & SqlDateTime(toDate) & " " & _
          "ORDER BY ErrorDate DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' پاک کردن ListBox
    While Me.lstErrorLogs.ListCount > 0
        Me.lstErrorLogs.RemoveItem 0
    Wend
    
    count = 0
    Do While Not rs.EOF
        Dim line As String
        Dim errorDate As String
        Dim errorNum As Long
        
        errorDate = Format$(rs!ErrorDate, "yyyy/mm/dd hh:mm:ss")
        errorNum = Nz(rs!ErrNumber, 0)
        
        line = "[" & errorDate & "] خطای " & CStr(errorNum) & " در " & _
               Nz(rs!Procedure, "نامعلوم") & ": " & Nz(rs!ErrDescription, "")
        
        Me.lstErrorLogs.AddItem line & "|" & rs!ErrorID
        count = count + 1
        rs.MoveNext
    Loop
    
    Me.lblCount.Caption = "کل: " & CStr(count) & " خطا"
    
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
' رویداد: cmdFilter_Click
' =========================================================

Private Sub cmdFilter_Click()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim fromDate As Date
    Dim toDate As Date
    Dim errorNum As String
    Dim procName As String
    Dim count As Long
    Dim whereClause As String
    
    Set db = CurrentDb()
    
    fromDate = Me.dtFromDate.Value
    toDate = Me.dtToDate.Value + 1
    errorNum = Trim$(Me.txtErrorNumber.Value)
    procName = Trim$(Me.txtProcedure.Value)
    
    If Not TableExists("tblErrorLogs") Then
        MsgBox "خطایی ثبت نشده", vbInformation
        Exit Sub
    End If
    
    ' ساخت شرط جستجو
    whereClause = "WHERE ErrorDate >= " & SqlDateTime(fromDate) & " " & _
                  "AND ErrorDate < " & SqlDateTime(toDate)
    
    If Len(errorNum) > 0 Then
        whereClause = whereClause & " AND ErrNumber=" & CLng(errorNum)
    End If
    
    If Len(procName) > 0 Then
        whereClause = whereClause & " AND Procedure LIKE '%" & Replace(procName, "'", "''") & "%'"
    End If
    
    ' جستجو
    sql = "SELECT ErrorID, ErrorDate, ErrNumber, ErrDescription, Procedure " & _
          "FROM tblErrorLogs " & whereClause & " " & _
          "ORDER BY ErrorDate DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' پاک کردن ListBox
    While Me.lstErrorLogs.ListCount > 0
        Me.lstErrorLogs.RemoveItem 0
    Wend
    
    count = 0
    Do While Not rs.EOF
        Dim line As String
        Dim errorDate As String
        Dim errorNum2 As Long
        
        errorDate = Format$(rs!ErrorDate, "yyyy/mm/dd hh:mm:ss")
        errorNum2 = Nz(rs!ErrNumber, 0)
        
        line = "[" & errorDate & "] خطای " & CStr(errorNum2) & " در " & _
               Nz(rs!Procedure, "نامعلوم") & ": " & Nz(rs!ErrDescription, "")
        
        Me.lstErrorLogs.AddItem line & "|" & rs!ErrorID
        count = count + 1
        rs.MoveNext
    Loop
    
    Me.lblCount.Caption = "نتیجهٌ فیلتر: " & CStr(count) & " خطا"
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در اعمال فیلتر: " & Err.Description, vbCritical
End Sub

' =========================================================
' رویداد: cmdClear_Click
' =========================================================

Private Sub cmdClear_Click()
    On Error GoTo ErrHandler
    
    ' پاک کردن فیلترها
    Me.txtErrorNumber.Value = ""
    Me.txtProcedure.Value = ""
    Me.dtFromDate.Value = DateAdd("d", -7, Now())
    Me.dtToDate.Value = Now()
    
    ' بروزرسانی
    Call RefreshErrorLogs
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

' =========================================================
' رویداد: lstErrorLogs_DblClick
' =========================================================

Private Sub lstErrorLogs_DblClick(Cancel As Integer)
    On Error GoTo ErrHandler
    
    If Me.lstErrorLogs.ListIndex = -1 Then Exit Sub
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim errorID As Long
    Dim sql As String
    Dim details As String
    
    ' استخراج ErrorID
    errorID = CLng(Me.lstErrorLogs.Column(1, Me.lstErrorLogs.ListIndex))
    
    Set db = CurrentDb()
    
    sql = "SELECT ErrorDate, ErrNumber, ErrDescription, Procedure, AdditionalInfo " & _
          "FROM tblErrorLogs WHERE ErrorID=" & CStr(errorID)
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        details = "جزئیات خطا:" & vbCrLf & vbCrLf & _
                  "شماره خطا: " & Nz(rs!ErrNumber, 0) & vbCrLf & _
                  "تاریخ: " & Format$(rs!ErrorDate, "yyyy/mm/dd hh:mm:ss") & vbCrLf & _
                  "تابع: " & Nz(rs!Procedure, "نامعلوم") & vbCrLf & vbCrLf & _
                  "شرح خطا:" & vbCrLf & Nz(rs!ErrDescription, "") & vbCrLf & vbCrLf & _
                  "اطلاعات اضافی:" & vbCrLf & Nz(rs!AdditionalInfo, "")
        
        MsgBox details, vbInformation
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در نمایش جزئیات: " & Err.Description, vbCritical
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
    fileName = folderPath & "\ErrorLogs_" & Format$(Now(), "yyyymmdd_hhmmss") & ".txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(fileName, True, True)
    
    ' نوشتن هدر
    file.WriteLine "گزارش خطاهای سیستم"
    file.WriteLine "=" & String(60, "=")
    file.WriteLine ""
    file.WriteLine "تاریخ صادرات: " & Now()
    file.WriteLine ""
    
    ' نوشتن خطاها
    Dim i As Long
    For i = 0 To Me.lstErrorLogs.ListCount - 1
        file.WriteLine Me.lstErrorLogs.List(i)
        count = count + 1
    Next i
    
    file.WriteLine ""
    file.WriteLine "=" & String(60, "=")
    file.WriteLine "کل خطاهای صادرشده: " & count
    
    file.Close
    
    MsgBox "خطاها صادر شدند" & vbCrLf & "مسیر: " & fileName, vbInformation
    
CleanExit:
    Set file = Nothing
    Set fso = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در صادرات: " & Err.Description, vbCritical
End Sub

' =========================================================
' رویداد: cmdDeleteOld_Click
' =========================================================

Private Sub cmdDeleteOld_Click()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim sql As String
    Dim deletedCount As Long
    Dim cutoffDate As Date
    
    ' درخواست تأیید
    If MsgBox("آیا خطاهای قدیمی‌تر از 30 روز حذف شوند؟", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If
    
    Set db = CurrentDb()
    
    cutoffDate = DateAdd("d", -30, Now())
    
    sql = "DELETE FROM tblErrorLogs WHERE ErrorDate < " & SqlDateTime(cutoffDate)
    
    db.Execute sql
    
    deletedCount = db.RecordsAffected
    
    MsgBox "حذف شد: " & CStr(deletedCount) & " خطای قدیم", vbInformation
    
    Call RefreshErrorLogs
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در حذف: " & Err.Description, vbCritical
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
    Set folder = shell.BrowseForFolder(0, "مسیری برای ذخیره خطاها را انتخاب کنید:", 0, 0)
    
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

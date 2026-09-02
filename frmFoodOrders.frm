Option Compare Database
Option Explicit

' =========================================================
' فرم: frmFoodOrders
' =========================================================
'
' توضیح فرم:
' این فرم سفارشات غذای روزانه را نمایش و مدیریت می‌کند.
' کاربران می‌توانند برای روزهای آینده غذا سفارش دهند.
'
' کاربرد:
' - نمایش لیست سفارشات
' - اضافه کردن سفارش جدید
' - ویرایش سفارش موجود
' - حذف سفارش
' - فیلتر بر اساس تاریخ
'
' کنترل‌ها:
' - lstFoodOrders: لیستی از سفارشات
' - cmbEmployee: انتخاب کارمند
' - cmbMealType: نوع غذا
' - txtReserveDate: تاریخ رزرو
' - cmdNew, cmdSave, cmdDelete, cmdClose
'
' رویدادها:
' - Form_Load: بارگزاری فرم
' - cmdNew_Click: سفارش جدید
' - cmdSave_Click: ذخیره
' - cmdDelete_Click: حذف
' - cmdClose_Click: بسته کردن
'
' =========================================================

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    ' تنظیمات فرم
    Me.Caption = "مدیریت سفارشات غذا"
    Me.RecordSource = TABLE_MEAL_ORDERS
    
    ' بارگزاری لیست کارمندان
    Call LoadEmployeeCombo
    
    ' بارگزاری لیست نوع غذا
    Call LoadMealTypeCombo
    
    ' بارگزاری سفارشات امروز
    Call RefreshFoodOrdersList
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در بارگزاری فرم: " & Err.Description, vbCritical
End Sub

' =========================================================
' زیربرنامه: LoadEmployeeCombo
' =========================================================
'
' وظیفه:
' لیست کارمندان فعال را بارگزاری می‌کند
'
' =========================================================

Private Sub LoadEmployeeCombo()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb()
    
    ' پاک کردن لیست قبلی
    If Me.cmbEmployee.ListCount > 0 Then
        Me.cmbEmployee.Clear
    End If
    
    ' اضافه کردن گزینهٌ خالی
    Me.cmbEmployee.AddItem ""
    
    ' بارگزاری کارمندان فعال
    sql = "SELECT EmployeeID, FirstName, LastName FROM " & TABLE_EMPLOYEES & " " & _
          "WHERE IsActive=True ORDER BY FirstName"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    Do While Not rs.EOF
        Dim empName As String
        empName = Trim$(rs!FirstName) & " " & Trim$(rs!LastName)
        Me.cmbEmployee.AddItem empName & " (" & rs!EmployeeID & ")"
        rs.MoveNext
    Loop
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در بارگزاری کارمندان: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' =========================================================
' زیربرنامه: LoadMealTypeCombo
' =========================================================
'
' وظیفه:
' انواع غذا را بارگزاری می‌کند
'
' =========================================================

Private Sub LoadMealTypeCombo()
    On Error GoTo ErrHandler
    
    ' پاک کردن لیست قبلی
    If Me.cmbMealType.ListCount > 0 Then
        Me.cmbMealType.Clear
    End If
    
    ' انواع غذای پیش‌فرض
    Me.cmbMealType.AddItem ""
    Me.cmbMealType.AddItem "ناهار"
    Me.cmbMealType.AddItem "صبحانه"
    Me.cmbMealType.AddItem "شام"
    Me.cmbMealType.AddItem "میان‌وعده"
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا: " & Err.Description, vbCritical
End Sub

' =========================================================
' زیربرنامه: RefreshFoodOrdersList
' =========================================================
'
' وظیفه:
' لیست سفارشات امروز را بروزرسانی می‌کند
'
' =========================================================

Private Sub RefreshFoodOrdersList()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim todayDate As Date
    
    Set db = CurrentDb()
    todayDate = DateValue(Now())
    
    ' جستجو برای سفارشات امروز
    sql = "SELECT OrderID, FirstName, LastName, MealType, ReserveDate " & _
          "FROM " & TABLE_MEAL_ORDERS & " " & _
          "WHERE DateValue(ReserveDate)=DateValue(Now()) " & _
          "ORDER BY FirstName"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' پاک کردن ListBox قبلی
    If Me.lstFoodOrders.ListCount > 0 Then
        While Me.lstFoodOrders.ListCount > 0
            Me.lstFoodOrders.RemoveItem 0
        Wend
    End If
    
    ' اضافه کردن رکوردها
    Do While Not rs.EOF
        Dim line As String
        line = Trim$(rs!FirstName) & " " & Trim$(rs!LastName) & " - " & rs!MealType
        Me.lstFoodOrders.AddItem line & "|" & rs!OrderID
        rs.MoveNext
    Loop
    
    ' نمایش تعداد
    Me.lblCount.Caption = "کل سفارشات: " & Me.lstFoodOrders.ListCount
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در بروزرسانی لیست: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' =========================================================
' رویداد: cmdNew_Click
' =========================================================

Private Sub cmdNew_Click()
    On Error GoTo ErrHandler
    
    Me.cmbEmployee.Value = ""
    Me.cmbMealType.Value = ""
    Me.txtReserveDate.Value = format_jalali_date(Now(), "SHORT")
    
    Me.cmbEmployee.SetFocus
    
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
    Dim empID As String
    
    ' بررسی: کارمند انتخاب شده
    If Len(Trim$(Me.cmbEmployee.Value)) = 0 Then
        MsgBox "لطفاً کارمند را انتخاب کنید", vbExclamation
        Me.cmbEmployee.SetFocus
        Exit Sub
    End If
    
    ' بررسی: نوع غذا انتخاب شده
    If Len(Trim$(Me.cmbMealType.Value)) = 0 Then
        MsgBox "لطفاً نوع غذا را انتخاب کنید", vbExclamation
        Me.cmbMealType.SetFocus
        Exit Sub
    End If
    
    ' استخراج کد کارمند
    empID = Mid$(Me.cmbEmployee.Value, InStr(Me.cmbEmployee.Value, "(") + 1)
    empID = Left$(empID, Len(empID) - 1)
    
    Set db = CurrentDb()
    
    ' درج سفارش جدید
    sql = "INSERT INTO " & TABLE_MEAL_ORDERS & " " & _
          "(EmployeeID, MealType, ReserveDate, CreatedDate) VALUES " & _
          "('" & empID & "', '" & Me.cmbMealType.Value & "', " & _
          "DateValue(Now()), Now())"
    
    db.Execute sql
    
    MsgBox "سفارش ذخیره شد", vbInformation
    
    Call RefreshFoodOrdersList
    Call cmdNew_Click
    
    Exit Sub
    
ErrHandler:
    MsgBox "خطا در ذخیره: " & Err.Description, vbCritical
    Resume Next
End Sub

' =========================================================
' رویداد: cmdDelete_Click
' =========================================================

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim orderID As Long
    Dim sql As String
    
    If Me.lstFoodOrders.ListIndex = -1 Then
        MsgBox "لطفاً سفارشی را انتخاب کنید", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("آیا مطمئن هستید؟", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If
    
    ' استخراج OrderID
    orderID = CLng(Me.lstFoodOrders.Column(1, Me.lstFoodOrders.ListIndex))
    
    Set db = CurrentDb()
    
    sql = "DELETE FROM " & TABLE_MEAL_ORDERS & " WHERE OrderID=" & CStr(orderID)
    db.Execute sql
    
    MsgBox "سفارش حذف شد", vbInformation
    
    Call RefreshFoodOrdersList
    
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

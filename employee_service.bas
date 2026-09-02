Option Compare Database
Option Explicit

' =========================================================
' ماژول: employee_service.bas
' =========================================================
' 
' توضیح ماژول:
' مدیریت اطلاعات کارمندان
' جستجو بر اساس کد ثبتی دستگاه
'
' =========================================================

' =========================================================
' تابع: emp_ResolveByDeviceEnroll
' =========================================================
' 
' وظیفه:
' کد پرسنلی و نام را از کد ثبتی دستگاه پیدا می‌کند
'
' پارامتر:
'   deviceEnroll (String): کد ثبتی از دستگاه
'
' خروجی: Variant (Array)
'   Array(0) = EmployeeID
'   Array(1) = FullName
'   یا Null اگر یافت نشد
'
' =========================================================

Public Function emp_ResolveByDeviceEnroll(ByVal deviceEnroll As String) As Variant
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim enroll As String
    Dim safe As String
    Dim result(0 To 1) As Variant
    
    enroll = Trim$(deviceEnroll)
    
    If Len(enroll) = 0 Then
        emp_ResolveByDeviceEnroll = Null
        Exit Function
    End If
    
    safe = Replace(enroll, "'", "''")
    
    Set db = CurrentDb()
    
    ' جستجو بر اساس DeviceUserID یا EmployeeID
    Set rs = db.OpenRecordset( _
        "SELECT TOP 1 EmployeeID, DeviceUserID, FullName " & _
        "FROM " & TABLE_EMPLOYEES & " " & _
        "WHERE DeviceUserID='" & safe & "' OR EmployeeID='" & safe & "' " & _
        "ORDER BY IIf(DeviceUserID='" & safe & "',0,1), EmployeeID", _
        dbOpenSnapshot)
    
    If rs.EOF Then
        emp_ResolveByDeviceEnroll = Null
    Else
        result(0) = Nz(rs!EmployeeID, Nz(rs!DeviceUserID, ""))
        result(1) = Nz(rs!FullName, "")
        emp_ResolveByDeviceEnroll = result
    End If
    
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    Call LogError("emp_ResolveByDeviceEnroll", Err.Number, Err.Description, _
                  "DeviceEnroll=" & enroll)
    emp_ResolveByDeviceEnroll = Null
    Resume CleanExit
End Function

' =========================================================
' تابع: emp_GetAll
' =========================================================
' 
' وظیفه:
' تمام کارمندان را برمی‌گرداند
'
' خروجی: Recordset
'
' =========================================================

Public Function emp_GetAll() As DAO.Recordset
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    Set emp_GetAll = db.OpenRecordset( _
        "SELECT EmployeeID, FullName, NationalCode, IsActive FROM " & TABLE_EMPLOYEES & " " & _
        "ORDER BY FullName", dbOpenSnapshot)
    
    Exit Function
    
ErrHandler:
    Call LogError("emp_GetAll", Err.Number, Err.Description, "")
    Set emp_GetAll = Nothing
End Function

' =========================================================
' تابع: emp_InsertOrUpdate
' =========================================================
' 
' وظیفه:
' کارمند جدید را اضافه یا موجود را بروزرسانی می‌کند
'
' =========================================================

Public Function emp_InsertOrUpdate(ByVal empID As String, ByVal fullName As String, _
                                    ByVal nationalCode As String, ByVal deviceUserID As String) As Boolean
    On Error GoTo ErrHandler
    
    empID = Trim$(empID)
    If empID = "" Then Exit Function
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    
    ' جستجو برای موجود
    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_EMPLOYEES & " WHERE EmployeeID='" & Replace(empID, "'", "''") & "'", _
        dbOpenDynaset)
    
    If rs.EOF Then
        ' درج جدید
        rs.AddNew
        rs!EmployeeID = empID
    Else
        ' بروزرسانی
        rs.Edit
    End If
    
    rs!FullName = fullName
    rs!NationalCode = nationalCode
    If Len(Trim$(deviceUserID)) > 0 Then
        rs!DeviceUserID = deviceUserID
    End If
    rs!IsActive = True
    
    If rs.EOF Then
        rs!CreatedDate = Now()
    End If
    
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    emp_InsertOrUpdate = True
    Exit Function
    
ErrHandler:
    Call LogError("emp_InsertOrUpdate", Err.Number, Err.Description, "EmpID=" & empID)
    emp_InsertOrUpdate = False
End Function
Option Compare Database
Option Explicit

' =========================================================
' ماژول: receipt_service.bas
' =========================================================
' سرویس مدیریت فیش‌ها: ایجاد، خواندن و به‌روزرسانی وضعیت فیش‌ها' =========================================================

Public Function receipt_CreatePending(ByVal empID As String, ByVal mealListID As Long, ByVal deviceIP As String, ByVal attendanceRecordID As Long) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb()
    Set rs = db.OpenRecordset(TABLE_RECEIPTS, dbOpenDynaset)
    rs.AddNew
    rs!EmployeeID = empID
    rs!MealListID = mealListID
    rs!DeviceIP = deviceIP
    rs!AttendanceRecordID = attendanceRecordID
    rs!PrintStatus = "" ' pending/unknown
    rs!CreatedDate = Now()
    rs.Update
    receipt_CreatePending = Nz(rs!ReceiptID, 0)
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    Call LogError("receipt_CreatePending", Err.Number, Err.Description, CStr(attendanceRecordID))
    receipt_CreatePending = 0
    Resume CleanExit
End Function

Public Function receipt_GetByAttendance(ByVal attendanceRecordID As Long) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT ReceiptID FROM " & TABLE_RECEIPTS & " WHERE AttendanceRecordID=" & CStr(attendanceRecordID), dbOpenSnapshot)
    If rs.EOF Then
        receipt_GetByAttendance = 0
    Else
        receipt_GetByAttendance = Nz(rs!ReceiptID, 0)
    End If
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    Call LogError("receipt_GetByAttendance", Err.Number, Err.Description, CStr(attendanceRecordID))
    receipt_GetByAttendance = 0
    Resume CleanExit
End Function

Public Function receipt_GetPrintStatus(ByVal receiptID As Long) As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    If receiptID <= 0 Then Exit Function
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT PrintStatus FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), dbOpenSnapshot)
    If rs.EOF Then
        receipt_GetPrintStatus = ""
    Else
        receipt_GetPrintStatus = Nz(rs!PrintStatus, "")
    End If
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    Call LogError("receipt_GetPrintStatus", Err.Number, Err.Description, CStr(receiptID))
    receipt_GetPrintStatus = ""
    Resume CleanExit
End Function

Public Function receipt_SetPrintStatus(ByVal receiptID As Long, ByVal status As String, ByVal note As String) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    If receiptID <= 0 Then Exit Function
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT PrintStatus, Note FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), dbOpenDynaset)
    If rs.EOF Then GoTo CleanExit
    rs.Edit
    rs!PrintStatus = status
    If Len(Trim$(note)) > 0 Then rs!Note = note
    rs!LastStatusChange = Now()
    rs.Update
    receipt_SetPrintStatus = True
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    Call LogError("receipt_SetPrintStatus", Err.Number, Err.Description, CStr(receiptID))
    receipt_SetPrintStatus = False
    Resume CleanExit
End Function

Public Function receipt_GetMealListID(ByVal receiptID As Long) As Long
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    If receiptID <= 0 Then Exit Function
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT MealListID FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), dbOpenSnapshot)
    If rs.EOF Then
        receipt_GetMealListID = 0
    Else
        receipt_GetMealListID = Nz(rs!MealListID, 0)
    End If
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    Call LogError("receipt_GetMealListID", Err.Number, Err.Description, CStr(receiptID))
    receipt_GetMealListID = 0
    Resume CleanExit
End Function

Public Function receipt_AssignMeal(ByVal receiptID As Long, ByVal mealListID As Long) As Boolean
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    If receiptID <= 0 Or mealListID <= 0 Then Exit Function
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT MealListID FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), dbOpenDynaset)
    If rs.EOF Then GoTo CleanExit
    rs.Edit
    rs!MealListID = mealListID
    rs.Update
    receipt_AssignMeal = True
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
ErrHandler:
    Call LogError("receipt_AssignMeal", Err.Number, Err.Description, CStr(receiptID))
    receipt_AssignMeal = False
    Resume CleanExit
End Function

Public Function receipt_UpdatePrintDateTime(ByVal receiptID As Long, ByVal printDT As Date) As Boolean
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    If receiptID <= 0 Then Exit Function
    Set rs = CurrentDb().OpenRecordset("SELECT PrintDateTime FROM " & TABLE_RECEIPTS & " WHERE ReceiptID=" & CStr(receiptID), dbOpenDynaset)
    If rs.EOF Then GoTo CleanExit
    rs.Edit
    rs!PrintDateTime = printDT
    rs.Update
    receipt_UpdatePrintDateTime = True
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function
ErrHandler:
    Call LogError("receipt_UpdatePrintDateTime", Err.Number, Err.Description, CStr(receiptID))
    receipt_UpdatePrintDateTime = False
    Resume CleanExit
End Function

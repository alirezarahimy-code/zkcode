' =========================================================
' File: tests/e2e_test.bas
' Purpose: simple end-to-end test helper to create a sample attendance
'          and run processing to exercise the pipeline (migration -> process -> print)
' =========================================================
Option Compare Database
Option Explicit

Public Sub run_e2e_test()
    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim recID As Long
    Dim tempReceipt As Long

    Set db = CurrentDb()

    ' Ensure schema
    If Not db_migration_CreateOrUpdateSchema() Then
        MsgBox "db_migration failed", vbCritical
        Exit Sub
    End If

    ' Insert a sample employee if not exists
    Set rs = db.OpenRecordset("SELECT TOP 1 EmployeeID FROM " & TABLE_EMPLOYEES & " WHERE EmployeeID='E_TEST'", dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Set rs = db.OpenRecordset(TABLE_EMPLOYEES, dbOpenDynaset)
        rs.AddNew
        rs!EmployeeID = "E_TEST"
        rs!DeviceUserID = "1001"
        rs!FullName = "آزمون تست"
        rs!NationalCode = "0000000000"
        rs!IsActive = True
        rs.Update
        rs.Close
    Else
        rs.Close
    End If

    ' Insert a meal order for today
    Set rs = db.OpenRecordset(TABLE_DAILY_MEALS, dbOpenDynaset)
    rs.AddNew
    rs!MealDate = Date
    rs!EmployeeID = "E_TEST"
    rs!FullName = "آزمون تست"
    rs!MealType = "ناهار"
    rs!Quantity = 1
    rs.Update
    rs.Close

    ' Insert attendance raw
    Set rs = db.OpenRecordset(TABLE_ATTENDANCE, dbOpenDynaset)
    rs.AddNew
    rs!DeviceEnrollID = "1001"
    rs!AttendanceDateTime = Now()
    rs!DeviceIP = "127.0.0.1"
    rs!DevicePort = DEFAULT_ZK_PORT
    rs!DeviceMachineNumber = 1
    rs!RawKey = "TEST|1001|" & Format$(Now(), "yyyy-mm-dd HH:nn:ss")
    rs!ProcessingResult = "NEW"
    rs!PrintAttempts = 0
    rs!CreatedDate = Now()
    rs.Update
    rs.Bookmark = rs.LastModified
    recID = Nz(rs!RecordID, 0)
    rs.Close

    If recID <= 0 Then
        MsgBox "Failed to insert sample attendance", vbCritical
        Exit Sub
    End If

    ' Force processing
    Call proc_ProcessPendingBatch()

    MsgBox "E2E test completed. Check receipts and logs.", vbInformation

CleanExit:
    On Error Resume Next
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    MsgBox "E2E test error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub

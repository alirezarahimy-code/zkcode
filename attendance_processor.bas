Option Compare Database
Option Explicit

' =========================================================
' ماژول: attendance_processor.bas
' =========================================================
' پردازش صف ترددها: گرفتن رکوردهای NEW یا PENDING و پردازش آنها به صورت atomic' =========================================================

Public Sub proc_ProcessPendingBatch()
    On Error GoTo ErrHandler
    Dim db As DAO.Database, rs As DAO.Recordset
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT RecordID FROM " & TABLE_ATTENDANCE & " WHERE Nz(ProcessingResult,'') IN ('NEW','PENDING','PROCESSING') ORDER BY CreatedDate", dbOpenSnapshot)
    Do While Not rs.EOF
        Dim recID As Long
        recID = Nz(rs!RecordID, 0)
        If recID > 0 Then
            ' سعی می‌کنیم پردازش اتمیک برای هر رکورد انجام دهیم
            On Error Resume Next
            Call proc_ProcessRecordAtomic(recID)
            On Error GoTo ErrHandler
        End If
        rs.MoveNext
    Loop
CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
ErrHandler:
    Call LogError("proc_ProcessPendingBatch", Err.Number, Err.Description, "")
    Resume CleanExit
End Sub

' proc_ProcessRecordAtomic already implemented in earlier commit; ensure it exists in module scope.

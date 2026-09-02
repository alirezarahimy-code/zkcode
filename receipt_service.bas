Option Compare Database
Option Explicit

' =========================================================
' ماژول: receipt_service.bas
' =========================================================
' 
' توضیح ماژول:
' مدیریت فیش‌های چاپ‌شده
' ایجاد، بروزرسانی، پرس‌و‌جو فیش‌ها
'
' =========================================================

' (existing functions kept above...)

' =========================================================
' تابع: receipt_UpdatePrintDateTime
' =========================================================
' 
' وظیفه:
' بروزرسانی تاریخ و ساعت چاپ فیش (متمرکزکردن در receipt_service)' 
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

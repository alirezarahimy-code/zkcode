Option Compare Database
Option Explicit

' =========================================================
' ماژول: attendance_service.bas
' =========================================================
' 
' توضیح ماژول:
' این ماژول درج تردد‌های جدید را مدیریت می‌کند.
' هر تردد دریافت شده از دستگاه درج می‌شود.
'
' کاربرد:
' - ایجاد رکورد تردد جدید
' - تولید کلید یکتای RawKey
' - جلوگیری از تاثیرات دوگانه
'
' =========================================================

' =========================================================
' تابع: att_InsertRaw
' =========================================================
' 
' وظیفه:
' یک تردد خام را درج می‌کند
' RawKey برای جلوگیری از Duplicate استفاده می‌شود
'
' پارامترها:
'   deviceEnroll (String): کد ثبتی
'   deviceIP (String): IP دستگاه
'   devicePort (Long): پورت
'   deviceMachineNumber (Long): شماره دستگاه
'   attDT (Date): تاریخ/زمان تردد
'   inOut (String): نوع تردد (IN/OUT/etc)
'   rawKey (String): کلید یکتای (ماشین|ثبتی|زمان|نوع|کد)
'   processingState (String): وضعیت اولیه (NEW/HISTORICAL)
'
' خروجی: Long
'   شماره رکورد یا 0 اگر ناموفق
'
' =========================================================

Public Function att_InsertRaw(ByVal deviceEnroll As String, ByVal deviceIP As String, _
                              ByVal devicePort As Long, ByVal deviceMachineNumber As Long, _
                              ByVal attDT As Date, ByVal inOut As String, ByVal rawKey As String, _
                              Optional ByVal processingState As String = "NEW") As Long
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim safeEnroll As String, safeIP As String, safeRawKey As String
    Dim existingID As Long, transactionStarted As Boolean
    
    If devicePort <= 0 Then devicePort = DEFAULT_ZK_PORT
    If deviceMachineNumber <= 0 Then deviceMachineNumber = 1
    
    safeEnroll = Replace(Trim$(deviceEnroll), "'", "''")
    safeIP = Replace(Trim$(deviceIP), "'", "''")
    safeRawKey = Replace(Trim$(rawKey), "'", "''")
    
    Set db = CurrentDb()
    
    ' بررسی اگر RawKey قبلاً موجود است
    If Len(Trim$(rawKey)) > 0 Then
        Set rs = db.OpenRecordset("SELECT TOP 1 RecordID FROM " & TABLE_ATTENDANCE & _
                                  " WHERE RawKey='" & safeRawKey & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            existingID = Nz(rs!RecordID, 0)
            rs.Close
            Set rs = Nothing
            Set db = Nothing
            att_InsertRaw = existingID
            Call LogSystemEvent("att_InsertRaw", "رکورد موجود: RawKey=" & rawKey)
            Exit Function
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    ' بررسی اگر این تردد قبلاً درج شد
    Dim sql As String
    sql = "SELECT TOP 1 RecordID FROM " & TABLE_ATTENDANCE & " " & _
          "WHERE DeviceEnrollID='" & safeEnroll & "' " & _
          "AND AttendanceDateTime=" & SqlDateTime(attDT) & " " & _
          "AND AttendanceType='" & Replace(Trim$(inOut), "'", "''") & "' " & _
          "AND DeviceIP='" & safeIP & "' " & _
          "AND DevicePort=" & CStr(devicePort) & " " & _
          "AND DeviceMachineNumber=" & CStr(deviceMachineNumber)
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then
        existingID = Nz(rs!RecordID, 0)
        rs.Close
        Set rs = Nothing
        Set db = Nothing
        att_InsertRaw = existingID
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
    
    ' درج رکورد جدید (با Transaction)
    DBEngine(0).BeginTrans
    transactionStarted = True
    
    Set rs = db.OpenRecordset(TABLE_ATTENDANCE, dbOpenDynaset)
    
    rs.AddNew
    rs!DeviceEnrollID = Trim$(deviceEnroll)
    rs!EmployeeID = ""
    rs!FullName = ""
    rs!AttendanceDateTime = attDT
    rs!AttendanceType = inOut
    rs!DeviceIP = deviceIP
    rs!DevicePort = devicePort
    rs!DeviceMachineNumber = deviceMachineNumber
    rs!DeviceKey = MakeDeviceKey(deviceIP, devicePort, deviceMachineNumber)
    
    If Len(Trim$(rawKey)) > 0 Then
        rs!RawKey = rawKey
    Else
        rs!RawKey = Null
    End If
    
    rs!IsProcessed = False
    processingState = UCase$(Trim$(processingState))
    If processingState = "" Then processingState = "NEW"
    rs!ProcessingResult = processingState
    rs!PrintAttempts = 0
    rs!CreatedDate = Now()
    
    rs.Update
    rs.Bookmark = rs.LastModified
    existingID = Nz(rs!RecordID, 0)
    
    rs.Close
    Set rs = Nothing
    
    DBEngine(0).CommitTrans
    transactionStarted = False
    
    Set db = Nothing
    att_InsertRaw = existingID
    
    If existingID > 0 Then
        Call LogSystemEvent("att_InsertRaw", "رکورد جدید: ID=" & CStr(existingID))
    End If
    
    Exit Function
    
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If transactionStarted Then DBEngine(0).Rollback
    Set db = Nothing
    
    ' تلاش مجدد برای بررسی موجودیت
    On Error GoTo RequeryFailed
    Set db = CurrentDb()
    If Len(Trim$(rawKey)) > 0 Then
        Set rs = db.OpenRecordset("SELECT TOP 1 RecordID FROM " & TABLE_ATTENDANCE & _
                                  " WHERE RawKey='" & safeRawKey & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            existingID = Nz(rs!RecordID, 0)
            rs.Close
            Set rs = Nothing
            Set db = Nothing
            att_InsertRaw = existingID
            Exit Function
        End If
        rs.Close
        Set rs = Nothing
    End If
    
RequeryFailed:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Call LogError("att_InsertRaw", en, ed, "Enroll=" & deviceEnroll & " | RawKey=" & rawKey)
    att_InsertRaw = 0
End Function
Option Compare Database
Option Explicit

' =========================================================
' ماژول: globals.bas (ensure no IsNumeric override)' =========================================================
Option Compare Database
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public DeviceSessions As Object
Public ZKRealtimeSessions As Object
Public MonitorRunning As Boolean
Public MonitorBusy As Boolean

Public Function MakeDeviceKey(ByVal ip As String, ByVal port As Long, Optional ByVal machineNumber As Variant) As String
    Dim p As Long, m As Long
    p = port: If p <= 0 Then p = DEFAULT_ZK_PORT
    m = 1
    If Not IsMissing(machineNumber) Then
        If Not IsNull(machineNumber) Then
            If VBA.IsNumeric(machineNumber) Then If CLng(machineNumber) > 0 Then m = CLng(machineNumber)
        End If
    End If
    MakeDeviceKey = Trim$(ip) & ":" & CStr(p) & ":" & CStr(m)
End Function

Public Function SqlDateTime(ByVal d As Date) As String
    SqlDateTime = "#" & Format$(d, "yyyy\/mm\/dd HH:nn:ss") & "#"
End Function

Public Function ShortGuid() As String
    Static initialized As Boolean
    If Not initialized Then Randomize Timer: initialized = True
    ShortGuid = Format$(CLng(Int(Rnd() * 1000000)), "000000")
End Function

Public Sub EnsureDeviceSessions()
    On Error GoTo EH
    If DeviceSessions Is Nothing Then
        Set DeviceSessions = CreateObject("Scripting.Dictionary")
        DeviceSessions.CompareMode = vbBinaryCompare
    End If
    Exit Sub
EH:
    Set DeviceSessions = Nothing
    Call LogError("EnsureDeviceSessions", Err.Number, Err.Description, "")
    Err.Clear
End Sub

Public Sub EnsureZKRealtimeSessions()
    On Error GoTo EH
    If ZKRealtimeSessions Is Nothing Then Set ZKRealtimeSessions = CreateObject("Scripting.Dictionary")
    Exit Sub
EH:
    Set ZKRealtimeSessions = Nothing
    Call LogError("EnsureZKRealtimeSessions", Err.Number, Err.Description, "")
    Err.Clear
End Sub

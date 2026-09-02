Option Compare Database
Option Explicit

' =========================================================
' ماژول: modForms_Helper.bas
' =========================================================
' 
' توضیح:
' توابع کمکی برای فرم‌ها
' مثل: SetTimer, KillTimer, Colors, Fonts
'
' =========================================================

#If VBA7 Then
    Public Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    
    Public Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
    Public Declare Function SetTimer Lib "user32" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    
    Public Declare Function KillTimer Lib "user32" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
#End If

Public Function IsDate(ByVal value As Variant) As Boolean
    On Error GoTo EH
    IsDate = Not IsNull(CDate(value))
    Exit Function
EH:
    IsDate = False
End Function

Public Function NZ(ByVal value As Variant, ByVal replacement As Variant) As Variant
    If IsNull(value) Or value = "" Then
        NZ = replacement
    Else
        NZ = value
    End If
End Function
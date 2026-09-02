Option Compare Database
Option Explicit

' =========================================================
' ماژول: app_setup.bas
' =========================================================
' 
' توضیح:
' راه‌اندازی اولیه برنامه
' فراخوانی شود هنگام شروع برنامه
'
' =========================================================

Public Sub Application_Start()
    On Error GoTo EH
    
    Application.Echo False, "درحال راه‌اندازی..."
    
    ' 1. مقداردهی متغیرهای سراسری
    Call EnsureDeviceSessions()
    Call EnsureZKRealtimeSessions()
    
    MonitorRunning = False
    MonitorBusy = False
    
    ' 2. ایجاد جداول
    If Not db_migration_CreateOrUpdateSchema() Then
        MsgBox "خطا در ایجاد پایگاه‌داده", vbCritical
        Exit Sub
    End If
    
    ' 3. باز کردن فرم اصلی
    Application.Echo True
    DoCmd.OpenForm "frmMain", acNormal
    
    Call LogSystemEvent("app_startup", "برنامه شروع شد")
    
    Exit Sub
    
EH:
    Application.Echo True
    MsgBox "خطا در راه‌اندازی: " & Err.Description, vbCritical
    Call LogError("Application_Start", Err.Number, Err.Description, "")
End Sub

Public Sub Application_Stop()
    On Error Resume Next
    
    Call monitor_Stop()
    Call LogSystemEvent("app_shutdown", "برنامه بسته شد")
End Sub
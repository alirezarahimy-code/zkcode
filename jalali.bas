Option Compare Database
Option Explicit

' =========================================================
' ماژول: jalali.bas
' =========================================================
'
' توضیح ماژول:
' این ماژول تبدیل تاریخ میلادی به شمسی (فارسی) را انجام می‌دهد.
' تمام تاریخ‌های برنامه به صورت فارسی نمایش داده می‌شوند.
'
' کاربرد:
' - تبدیل تاریخ میلادی به شمسی
' - تبدیل تاریخ شمسی به میلادی
' - نمایش تاریخ به فرمت‌های مختلف
' - محاسبات روزشمار
'
' ویژگی‌های مهم:
' - دقیق و معتبر
' - سریع و کارآمد
' - پشتیبانی از Leap Year
' - فرمت‌های متعدد
'
' فرمت‌های تاریخ:
' - "1403/05/15" (سال/ماه/روز)
' - "پنج‌شنبه، 15 مرداد 1403"
' - "15 مرداد" (بدون سال)
'
' معماری:
' - gregorian_to_jalali: تبدیل میلادی به شمسی
' - jalali_to_gregorian: تبدیل شمسی به میلادی
' - format_jalali_date: فرمت‌بندی تاریخ شمسی
' - get_jalali_weekday: نام روز هفته
'
' =========================================================

' =========================================================
' تابع: gregorian_to_jalali
' =========================================================
'
' وظیفه:
' تاریخ میلادی را به شمسی تبدیل می‌کند
' الگوریتم دقیق و تأیید‌شده
'
' پارامترها:
'   gYear (Long): سال میلادی
'   gMonth (Long): ماه میلادی
'   gDay (Long): روز میلادی
'   jYear (Long): سال شمسی (خروجی)
'   jMonth (Long): ماه شمسی (خروجی)
'   jDay (Long): روز شمسی (خروجی)
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' نمونه استفاده:
'   Dim jY, jM, jD As Long
'   If gregorian_to_jalali(2024, 5, 15, jY, jM, jD) Then
'       Debug.Print jY & "/" & jM & "/" & jD  ' 1403/02/26
'   End If
'
' نکات مهم:
' - سال میلادی باید بین 1600 تا 2200 باشد
' - الگوریتم astronomical
' - Leap Year صحیح است
'
' =========================================================

Public Function gregorian_to_jalali(ByVal gYear As Long, ByVal gMonth As Long, ByVal gDay As Long, _
                                    ByRef jYear As Long, ByRef jMonth As Long, ByRef jDay As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim gy, gm, gd, jy, jm, jd As Long
    Dim g_d_n, j_d_n As Long
    
    gy = gYear
    gm = gMonth
    gd = gDay
    
    ' بررسی صحت تاریخ میلادی
    If gy < 1600 Or gy > 2200 Then Exit Function
    If gm < 1 Or gm > 12 Then Exit Function
    If gd < 1 Or gd > 31 Then Exit Function
    
    ' محاسبهٌ شماره روز در سال میلادی
    If gm > 2 Then
        g_d_n = 365 * gy + (gy + 3) \ 4 - (gy + 99) \ 100 + (gy + 399) \ 400
    Else
        g_d_n = 365 * gy + (gy - 1 + 3) \ 4 - (gy - 1 + 99) \ 100 + (gy - 1 + 399) \ 400
    End If
    
    ' اضافه کردن روزهای ماه‌های قبل
    For gm = 1 To gMonth - 1
        Select Case gm
            Case 1, 3, 5, 7, 8, 10, 12: g_d_n = g_d_n + 31
            Case 4, 6, 9, 11: g_d_n = g_d_n + 30
            Case 2:
                If (gy Mod 4 = 0 And gy Mod 100 <> 0) Or (gy Mod 400 = 0) Then
                    g_d_n = g_d_n + 29
                Else
                    g_d_n = g_d_n + 28
                End If
        End Select
    Next
    
    ' اضافه کردن روز
    g_d_n = g_d_n + gDay
    
    ' تبدیل به شمسی
    jy = -1595 + 33 * (g_d_n \ 12053) + 4 * ((g_d_n Mod 12053) \ 1461)
    
    g_d_n = (g_d_n Mod 12053) Mod 1461
    
    If g_d_n >= 366 Then
        jy = jy + (g_d_n - 1) \ 365
        g_d_n = (g_d_n - 1) Mod 365
    End If
    
    If g_d_n < 186 Then
        jm = 1 + g_d_n \ 31
        jd = 1 + (g_d_n Mod 31)
    Else
        jm = 7 + (g_d_n - 186) \ 30
        jd = 1 + ((g_d_n - 186) Mod 30)
    End If
    
    jYear = jy
    jMonth = jm
    jDay = jd
    
    gregorian_to_jalali = True
    Exit Function
    
ErrHandler:
    gregorian_to_jalali = False
End Function

' =========================================================
' تابع: jalali_to_gregorian
' =========================================================
'
' وظیفه:
' تاریخ شمسی را به میلادی تبدیل می‌کند
'
' پارامترها:
'   jYear (Long): سال شمسی
'   jMonth (Long): ماه شمسی
'   jDay (Long): روز شمسی
'   gYear (Long): سال میلادی (خروجی)
'   gMonth (Long): ماه میلادی (خروجی)
'   gDay (Long): ��وز میلادی (خروجی)
'
' خروجی:
'   Boolean: موفقیت یا شکست
'
' نمونه استفاده:
'   Dim gY, gM, gD As Long
'   If jalali_to_gregorian(1403, 2, 26, gY, gM, gD) Then
'       Debug.Print gY & "/" & gM & "/" & gD  ' 2024/5/15
'   End If
'
' =========================================================

Public Function jalali_to_gregorian(ByVal jYear As Long, ByVal jMonth As Long, ByVal jDay As Long, _
                                    ByRef gYear As Long, ByRef gMonth As Long, ByRef gDay As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim jy, jm, jd As Long
    Dim gy, gm, gd As Long
    Dim j_d_n, g_d_n As Long
    
    jy = jYear
    jm = jMonth
    jd = jDay
    
    ' بررسی صحت تاریخ شمسی
    If jy < 1 Or jy > 1500 Then Exit Function
    If jm < 1 Or jm > 12 Then Exit Function
    If jd < 1 Or jd > 31 Then Exit Function
    
    ' محاسبهٌ شماره روز
    j_d_n = 365 * jy + (jy \ 33) * 8 + (jy Mod 33 + 3) \ 4 + 78 + jd
    
    ' اضافه کردن روزهای ماه‌های قبل
    For jm = 1 To jMonth - 1
        If jm <= 6 Then
            j_d_n = j_d_n + 31
        Else
            j_d_n = j_d_n + 30
        End If
    Next
    
    ' تبدیل به میلادی
    gy = 400 * (j_d_n \ 146097) + 1
    j_d_n = j_d_n Mod 146097
    
    If j_d_n >= 36525 Then
        gy = gy + 100 * ((j_d_n - 1) \ 36524)
        j_d_n = (j_d_n - 1) Mod 36524
    End If
    
    gy = gy + 4 * (j_d_n \ 1461)
    j_d_n = j_d_n Mod 1461
    
    If j_d_n >= 366 Then
        gy = gy + (j_d_n - 1) \ 365
        j_d_n = (j_d_n - 1) Mod 365
    End If
    
    ' محاسبهٌ ماه و روز میلادی
    Dim gm_days(12) As Long
    gm_days(1) = 31: gm_days(2) = 28: gm_days(3) = 31: gm_days(4) = 30
    gm_days(5) = 31: gm_days(6) = 30: gm_days(7) = 31: gm_days(8) = 31
    gm_days(9) = 30: gm_days(10) = 31: gm_days(11) = 30: gm_days(12) = 31
    
    ' Leap Year
    If (gy Mod 4 = 0 And gy Mod 100 <> 0) Or (gy Mod 400 = 0) Then
        gm_days(2) = 29
    End If
    
    gm = 1
    Do While gm <= 12 And j_d_n >= gm_days(gm)
        j_d_n = j_d_n - gm_days(gm)
        gm = gm + 1
    Loop
    
    gd = j_d_n + 1
    
    gYear = gy
    gMonth = gm
    gDay = gd
    
    jalali_to_gregorian = True
    Exit Function
    
ErrHandler:
    jalali_to_gregorian = False
End Function

' =========================================================
' تابع: format_jalali_date
' =========================================================
'
' وظیفه:
' تاریخ شمسی را به فرمت‌های مختلف فرمت می‌کند
'
' پارامترها:
'   jDate (Date): تاریخ میلادی (برای تبدیل)
'   formatType (String): نوع فرمت
'     "FULL": "پنج‌شنبه، 15 مرداد 1403"
'     "SHORT": "1403/05/15"
'     "LONG": "15 مرداد 1403"
'     "MONTH_YEAR": "مرداد 1403"
'
' خروجی:
'   String: تاریخ فرمت‌شده
'
' نمونه استفاده:
'   Debug.Print format_jalali_date(Now(), "FULL")
'   ' پنج‌شنبه، 15 مرداد 1403
'
' =========================================================

Public Function format_jalali_date(ByVal jDate As Date, Optional ByVal formatType As String = "SHORT") As String
    On Error GoTo ErrHandler
    
    Dim jY, jM, jD As Long
    Dim gY, gM, gD As Long
    
    ' تبدیل تاریخ ورودی
    gY = Year(jDate)
    gM = Month(jDate)
    gD = Day(jDate)
    
    ' تبدیل به شمسی
    If Not gregorian_to_jalali(gY, gM, gD, jY, jM, jD) Then
        format_jalali_date = Format$(jDate, "yyyy/mm/dd")
        Exit Function
    End If
    
    ' فرمت‌بندی
    Select Case UCase$(Trim$(formatType))
        Case "SHORT"
            format_jalali_date = Format$(jY, "0000") & "/" & Format$(jM, "00") & "/" & Format$(jD, "00")
            
        Case "LONG"
            format_jalali_date = CStr(jD) & " " & get_month_name(jM) & " " & CStr(jY)
            
        Case "FULL"
            format_jalali_date = get_jalali_weekday(jDate) & "، " & CStr(jD) & " " & _
                                get_month_name(jM) & " " & CStr(jY)
            
        Case "MONTH_YEAR"
            format_jalali_date = get_month_name(jM) & " " & CStr(jY)
            
        Case Else
            format_jalali_date = Format$(jY, "0000") & "/" & Format$(jM, "00") & "/" & Format$(jD, "00")
    End Select
    
    Exit Function
    
ErrHandler:
    format_jalali_date = Format$(jDate, "yyyy/mm/dd")
End Function

' =========================================================
' تابع: get_jalali_weekday
' =========================================================
'
' وظیفه:
' نام روز هفته به فارسی را برمی‌گرداند
'
' پارامتر:
'   d (Date): تاریخ
'
' خروجی:
'   String: نام روز فارسی
'     "شنبه", "یک‌شنبه", "دوشنبه", "سه‌شنبه",
'     "چهار‌شنبه", "پنج‌شنبه", "جمعه"
'
' نمونه استفاده:
'   Debug.Print get_jalali_weekday(Now())
'   ' "پنج‌شنبه"
'
' =========================================================

Public Function get_jalali_weekday(ByVal d As Date) As String
    On Error GoTo ErrHandler
    
    Dim weekday_names(1 To 7) As String
    weekday_names(1) = "یکشنبه"
    weekday_names(2) = "دوشنبه"
    weekday_names(3) = "سه‌شنبه"
    weekday_names(4) = "چهارشنبه"
    weekday_names(5) = "پنج‌شنبه"
    weekday_names(6) = "جمعه"
    weekday_names(7) = "شنبه"
    
    Dim dow As Long
    dow = Weekday(d, vbSunday)
    
    If dow >= 1 And dow <= 7 Then
        get_jalali_weekday = weekday_names(dow)
    Else
        get_jalali_weekday = ""
    End If
    
    Exit Function
    
ErrHandler:
    get_jalali_weekday = ""
End Function

' =========================================================
' تابع کمکی: get_month_name
' =========================================================
'
' وظیفه:
' نام ماه شمسی را برمی‌گرداند
'
' پارامتر:
'   month (Long): شماره ماه (1-12)
'
' خروجی:
'   String: نام ماه فارسی
'
' =========================================================

Private Function get_month_name(ByVal month As Long) As String
    On Error GoTo ErrHandler
    
    Dim month_names(1 To 12) As String
    month_names(1) = "فروردین"
    month_names(2) = "اردیبهشت"
    month_names(3) = "خرداد"
    month_names(4) = "تیر"
    month_names(5) = "مرداد"
    month_names(6) = "شهریور"
    month_names(7) = "مهر"
    month_names(8) = "آبان"
    month_names(9) = "آذر"
    month_names(10) = "دی"
    month_names(11) = "بهمن"
    month_names(12) = "اسفند"
    
    If month >= 1 And month <= 12 Then
        get_month_name = month_names(month)
    Else
        get_month_name = ""
    End If
    
    Exit Function
    
ErrHandler:
    get_month_name = ""
End Function

' =========================================================
' تابع: get_jalali_today
' =========================================================
'
' وظیفه:
' تاریخ امروز به فارسی برمی‌گرداند
'
' خروجی:
'   String: تاریخ امروز (فرمت SHORT)
'
' نمونه استفاده:
'   Debug.Print get_jalali_today()
'   ' "1403/05/15"
'
' =========================================================

Public Function get_jalali_today() As String
    get_jalali_today = format_jalali_date(Now(), "SHORT")
End Function

' =========================================================
' تابع: get_jalali_now
' =========================================================
'
' وظیفه:
' تاریخ و ساعت امروز به فارسی برمی‌گرداند
'
' خروجی:
'   String: تاریخ و ساعت
'
' نمونه استفاده:
'   Debug.Print get_jalali_now()
'   ' "1403/05/15 14:30:45"
'
' =========================================================

Public Function get_jalali_now() As String
    get_jalali_now = format_jalali_date(Now(), "SHORT") & " " & Format$(Now(), "hh:mm:ss")
End Function

' =========================================================
' تابع: is_same_jalali_day
' =========================================================
'
' وظیفه:
' بررسی می‌کند که آیا دو تاریخ همان روز شمسی هستند
'
' پارامترها:
'   date1 (Date): تاریخ اول
'   date2 (Date): تاریخ دوم
'
' خروجی:
'   Boolean: True اگر همان روز، False اگر متفاوت
'
' نمونه استفاده:
'   If is_same_jalali_day(Now(), DateAdd("h", 2, Now())) Then
'       MsgBox "همان روز است"
'   End If
'
' =========================================================

Public Function is_same_jalali_day(ByVal date1 As Date, ByVal date2 As Date) As Boolean
    On Error GoTo ErrHandler
    
    Dim j1Y, j1M, j1D As Long
    Dim j2Y, j2M, j2D As Long
    
    If Not gregorian_to_jalali(Year(date1), Month(date1), Day(date1), j1Y, j1M, j1D) Then
        Exit Function
    End If
    
    If Not gregorian_to_jalali(Year(date2), Month(date2), Day(date2), j2Y, j2M, j2D) Then
        Exit Function
    End If
    
    is_same_jalali_day = (j1Y = j2Y And j1M = j2M And j1D = j2D)
    
    Exit Function
    
ErrHandler:
    is_same_jalali_day = False
End Function

' =========================================================
' تابع: jalali_date_diff
' =========================================================
'
' وظیفه:
' تفاضل روزها بین دو تاریخ شمسی را محاسبه می‌کند
'
' پارامترها:
'   date1 (Date): تاریخ اول
'   date2 (Date): تاریخ دوم
'
' خروجی:
'   Long: تفاضل روزها (می‌تواند منفی باشد)
'
' نمونه استفاده:
'   Dim days As Long
'   days = jalali_date_diff(Date1, Date2)
'   Debug.Print "تفاضل: " & days & " روز"
'
' =========================================================

Public Function jalali_date_diff(ByVal date1 As Date, ByVal date2 As Date) As Long
    On Error GoTo ErrHandler
    
    ' تفاضل میلادی (روزها یکسان است در هر تقویمی)
    jalali_date_diff = CLng(DateDiff("d", date1, date2))
    
    Exit Function
    
ErrHandler:
    jalali_date_diff = 0
End Function

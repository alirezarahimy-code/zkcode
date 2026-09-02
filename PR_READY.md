# Pull Request Ready: refactor(core)

Title: refactor(core): consolidate DB/printing state, avoid shadowing builtins, add fallback printer helper

Body:
این PR بازسازی هسته‌ای کد را انجام می‌دهد تا برنامه پایدارتر، قابل‌نگهداری‌تر و آمادهٔ اجرا در محیط‌های ویندوزی ساده باشد.

خلاصهٔ تغییرات:
- حذف بازتعریف توابع توکار VB (مثلاً IsNumeric) و استفاده از توابع توکار.
- ShortGuid تولیدکنندهٔ 6 رقم با leading zeros شد.
- جداسازی لایهٔ چاپ: printer_service به عنوان façade، و print_helper_client مسئول ارسال low-level (fallback فعلی با فایل موقت + Notepad).
- یکپارچه‌سازی وضعیت‌های چاپ/پردازش: receipt_service مالک وضعیت فیش؛ db_service مالک state-related DB ops؛ attendance_processor اکنون از db_service برای به‌روزرسانی وضعیت‌ها استفاده می‌کند.
- افزودن db_SetPrintAttemptStarted برای ثبت تلاش چاپ.
- اصلاح monitor_service برای اعتبارسنجی ورودی‌ها.

فایل‌های مهم تغییر یافته:
- globals.bas
- monitor_service.bas
- printer_service.bas
- db_service.bas
- attendance_processor.bas
- print_helper_client.bas
- receipt_service.bas

راهنمای تست سریع (گام‌به‌گام):
1. اجرای `db_migration_CreateOrUpdateSchema()` در محیط Access برای ایجاد/به‌روزرسانی جداول.
2. تنظیم `tblPrinterSettings` (برای fallback نیازی به Winsock نیست؛ چاپ با Notepad انجام می‌شود).
3. اضافه کردن یک رکورد نمونه در `TABLE_ATTENDANCE` یا شبیه‌سازی رخداد realtime.
4. اجرای `proc_ProcessPendingBatch` یا `monitor_Tick` و بررسی:
   - ایجاد/بروزرسانی رکورد در `TABLE_RECEIPTS`,
   - ساخت فایل موقت فیش در پوشهٔ TEMP و فراخوانی Notepad برای چاپ،
   - ثبت وضعیت‌های `ProcessingResult` و `PrintStatus` در DB.
5. بررسی لاگ‌ها (`tblSystemLogs`, `tblErrorLogs`) برای خطاها.

Notes:
- Low-level printing: fallback implemented (file -> notepad). If you need direct socket printing, I'll add Winsock implementation on request.


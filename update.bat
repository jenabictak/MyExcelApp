@echo off
chcp 65001 > nul
echo ========================================
echo به‌روزرساني و همگام‌سازي پروژه با GitHub
echo ========================================

cd /d C:\inetpub\MyExcelApp

echo 1. دريافت آخرين تغييرات از GitHub (Pull)...
git pull origin main

echo 2. افزودن تغييرات جديد به Git...
git add .

echo 3. ثبت تغييرات با پيام...
git commit -m "به‌روزرساني خودکار - %date% %time%"

echo 4. ارسال تغييرات به GitHub (Push)...
git push origin main

echo 5. توقف سايت در IIS...
%windir%\system32\inetsrv\appcmd stop site "MyExcelApp"

echo 6. پابليش پروژه...
dotnet publish -c Release -o C:\inetpub\MyExcelApp\publish /p:MvcRazorCompileOnPublish=false

echo 7. شروع مجدد سايت...
%windir%\system32\inetsrv\appcmd start site "MyExcelApp"

echo 8. ريستارت IIS...
iisreset

echo ========================================
echo همگام‌سازي و به‌روزرساني با موفقيت انجام شد!
echo ========================================
pause
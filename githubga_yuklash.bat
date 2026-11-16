@echo off
chcp 65001 >nul
title Academic Bot - GitHub'ga yuklash
cd /d "%~dp0"

echo ============================================
echo   Academic Bot - GitHub'ga yuklash
echo   https://github.com/Eldor24/Academic_ai_bot
echo ============================================
echo.

where git >nul 2>&1
if errorlevel 1 (
    echo [XATO] Git o'rnatilmagan. https://git-scm.com/download/win dan o'rnating, keyin qayta ishga tushiring.
    pause
    exit /b 1
)

if not exist ".gitignore" (
    echo [XATO] .gitignore topilmadi. Maxfiy fayllar yuklanib ketmasligi uchun to'xtatildi.
    pause
    exit /b 1
)

if not exist ".git" (
    git init
    git checkout -b main
)

rem Muallif nomi (email o'rniga GitHub'ning yashirin manzili ishlatiladi)
git config user.name >nul 2>&1 || git config user.name "Eldor24"
git config user.email >nul 2>&1 || git config user.email "Eldor24@users.noreply.github.com"

git add -A

rem ── Xavfsizlik tekshiruvi: maxfiy fayllar ro'yxatda bo'lmasligi kerak ──
git diff --cached --name-only | findstr /I /R /C:"^\.env$" /C:"users\.json" /C:"\.log$" /C:"serverga yuklash" /C:"Baza_Zaxira" >nul
if not errorlevel 1 (
    echo [XATO] Maxfiy fayl yuklanish ro'yxatiga tushib qoldi:
    git diff --cached --name-only | findstr /I /R /C:"^\.env$" /C:"users\.json" /C:"\.log$" /C:"serverga yuklash" /C:"Baza_Zaxira"
    echo Yuklash to'xtatildi. Hech narsa yuborilmadi.
    git reset -q
    pause
    exit /b 1
)

echo GitHub'ga yuklanadigan fayllar:
git diff --cached --name-only
echo.

git commit -q -m "Academic AI Bot v11: 15 ta taqdimot shabloni, yangi xizmatlar, xavfsizlik qatlami"
if errorlevel 1 echo (Yangi o'zgarish yo'q yoki commit allaqachon qilingan)

git remote get-url origin >nul 2>&1 || git remote add origin https://github.com/Eldor24/Academic_ai_bot.git

echo.
echo GitHub'ga yuborilmoqda... (birinchi marta brauzerda GitHub'ga kirish oynasi ochiladi)
git push -u origin main
if errorlevel 1 (
    echo.
    echo GitHub'dagi repozitoriyada boshqa fayllar bor ekan, birlashtirilmoqda...
    git pull origin main --allow-unrelated-histories --no-edit -X ours
    git push -u origin main
)

if errorlevel 1 (
    echo.
    echo [XATO] Yuklab bo'lmadi. Yuqoridagi xabarni Claude'ga yuboring.
) else (
    echo.
    echo ✅ Tayyor! https://github.com/Eldor24/Academic_ai_bot
)
pause

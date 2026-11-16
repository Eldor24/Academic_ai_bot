@echo off
chcp 65001 >nul
title Academic Bot (lokal)
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo ============================================
echo   Academic Bot - lokal ishga tushirish
echo ============================================

set PY=
where py >nul 2>&1 && set PY=py -3
if not defined PY ( where python >nul 2>&1 && set PY=python )
if not defined PY (
    echo [XATO] Python topilmadi. https://www.python.org/downloads/ dan Python 3.11 yoki 3.12 ni o'rnating
    echo        O'rnatishda "Add python.exe to PATH" belgisini qo'ying.
    pause
    exit /b 1
)

if not exist ".env" (
    echo [XATO] .env fayli topilmadi. U bot.py bilan bir papkada bo'lishi kerak.
    pause
    exit /b 1
)

if not exist "venv\Scripts\python.exe" (
    echo [1/3] Virtual muhit yaratilmoqda...
    %PY% -m venv venv || ( echo [XATO] venv yaratib bo'lmadi & pause & exit /b 1 )
)

echo [2/3] Kutubxonalar o'rnatilmoqda (birinchi marta 2-5 daqiqa)...
"venv\Scripts\python.exe" -m pip install --upgrade pip -q
"venv\Scripts\python.exe" -m pip install -r requirements.txt -q || ( echo [XATO] Kutubxonalarni o'rnatib bo'lmadi & pause & exit /b 1 )

echo [3/3] Bot ishga tushdi. To'xtatish uchun shu oynada Ctrl+C bosing.
echo.
"venv\Scripts\python.exe" bot.py

echo.
echo Bot to'xtadi.
pause

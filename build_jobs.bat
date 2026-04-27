@echo off
chcp 65001 >nul
title RPA Job Build

echo ============================================
echo   RPA Job EXE 빌드 시작
echo ============================================
cd /d "%~dp0"

echo.
echo [1/2] daily_printout_automail.exe 빌드 중...
pyinstaller rpa_tasks/dailyprintout/daily_printout_automail.spec --distpath rpa_tasks/dailyprintout --noconfirm
if %errorlevel% neq 0 (
    echo [ERROR] daily_printout_automail.exe 빌드 실패
    pause & exit /b 1
)
echo [OK] daily_printout_automail.exe 빌드 완료

echo.
echo [2/2] db_sink_prod_to_dev.exe 빌드 중...
pyinstaller rpa_tasks/dailyprintout/db_sink_prod_to_dev.spec --distpath rpa_tasks/dailyprintout --noconfirm
if %errorlevel% neq 0 (
    echo [ERROR] db_sink_prod_to_dev.exe 빌드 실패
    pause & exit /b 1
)
echo [OK] db_sink_prod_to_dev.exe 빌드 완료

echo.
echo ============================================
echo   빌드 완료
echo   결과물: rpa_tasks/dailyprintout/*.exe
echo ============================================
pause

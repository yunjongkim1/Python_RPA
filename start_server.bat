@echo off
chcp 65001 >nul
title Mando RPA API Server

:: 현재 bat 파일 위치(프로젝트 루트)로 이동
cd /d "%~dp0"

:: PYTHONPATH를 현재 폴더로 설정 (core/ 등 모듈 인식)
set PYTHONPATH=%~dp0

:RESTART_LOOP
echo.
echo ============================================
echo   Starting Mando RPA API Portal Server...
echo ============================================

"%~dp0.venv\Scripts\python.exe" app/web/rpa_server.py

set EXIT_CODE=%errorlevel%

if %EXIT_CODE% neq 0 (
    echo.
    echo [ERROR] 서버가 비정상 종료되었습니다. (exit code: %EXIT_CODE%)
    echo.
)
echo.
echo   재시작하려면 "Terminate batch job?" 에서 [N] 입력
echo ============================================
echo   서버가 종료되었습니다.
echo   [R] 재시작   [아무 키] 종료
echo ============================================
choice /c RC /n /m "선택: "
if %errorlevel% equ 1 goto RESTART_LOOP

echo.
echo 서버를 종료합니다.
pause
@echo off
title Mando RPA API Server
echo ============================================
echo   Starting Mando RPA API Portal Server...
echo ============================================

:: 1. 현재 bat 파일 위치(프로젝트 루트)로 이동
cd /d "%~dp0"

:: 2. PYTHONPATH를 현재 폴더로 설정 (core/ 등 모듈 인식)
set PYTHONPATH=%~dp0

:: 3. 서버 실행
python app/rest_api_server.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 서버 실행에 실패했습니다.
    echo python --version 으로 Python이 설치되어 있는지 확인하세요.
)

pause
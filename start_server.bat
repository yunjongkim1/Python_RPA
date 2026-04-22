@echo off
title Mando RPA API Server
echo ============================================
echo   Starting Mando RPA API Portal Server...
echo ============================================

:: 1. .env 파일에서 PROJECT_ROOT 경로 추출 (따옴표 제거 처리)
for /f "tokens=2 delims==" %%a in ('findstr "PROJECT_ROOT" .env') do set "RAW_PATH=%%a"
set "ROOT_PATH=%RAW_PATH:"=%"

:: 2. 프로젝트 루트 폴더로 이동
cd /d "%ROOT_PATH%"

:: 3. 파이썬이 현재 폴더(PYTHON_RPA)를 모듈로 인식하도록 경로 강제 추가
set PYTHONPATH=%ROOT_PATH%

:: 4. 서버 실행 (파일명과 경로를 점(.)으로 연결)
:: 주의: app/web/rest_api_server.py가 실제로 있는지 확인하세요.
python -m uvicorn app.rest_api_server:app --host 0.0.0.0 --port 8000

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 서버 실행에 실패했습니다. 
    echo 1. app/web/__init__.py 파일이 있는지 확인하세요.
    echo 2. rest_api_server.py 파일명이 정확한지 확인하세요.
)

pause
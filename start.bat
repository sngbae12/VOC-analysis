@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo Python을 찾지 못했습니다. Python 3.10 이상을 설치하거나 프로젝트에 .venv를 만들어 주세요.
    pause
    exit /b 1
  )
  set "PY=python"
)

echo [고객사 VOC 분석] %PY% 로 실행합니다.
echo 브라우저에서 http://127.0.0.1:7860 을 엽니다.
"%PY%" voc_analysis_app.py
if errorlevel 1 (
  echo.
  echo 실행에 실패했습니다. README의 설치 방법을 확인하세요.
  pause
)

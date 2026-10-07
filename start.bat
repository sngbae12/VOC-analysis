@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  set "PY=.venv\Scripts\python.exe"
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo Python을 찾지 못했습니다. Python 3.11을 설치하거나 setup.bat으로 .venv를 만드세요.
    pause
    exit /b 1
  )
  set "PY=python"
)

"%PY%" -c "import sys; v=sys.version_info; raise SystemExit(0 if (v.major, v.minor) in {(3,11),(3,12),(3,13)} else 1)"
if errorlevel 1 (
  echo 이 앱은 Python 3.11, 3.12, 3.13만 지원합니다. Python 3.11을 권장합니다.
  "%PY%" -c "import sys; print('현재 실행 파일:', sys.executable); print(sys.version)"
  echo setup.bat을 해당 버전으로 다시 실행하세요.
  pause
  exit /b 1
)

"%PY%" -c "import gradio,pandas,matplotlib,wordcloud,docx,crewai,openai,plotly" >nul 2>&1
if errorlevel 1 (
  echo 필요한 패키지가 없습니다. 프로젝트 폴더에서 setup.bat을 먼저 실행하세요.
  echo 현재 Python: %PY%
  pause
  exit /b 1
)

echo [고객사 VOC 분석] %PY% 로 실행합니다.
echo 브라우저에서 http://127.0.0.1:7860 을 엽니다.
echo 포트가 사용 중이면 실행이 중단됩니다. 기존 앱을 종료한 뒤 다시 시도하세요.
"%PY%" voc_analysis_app.py
if errorlevel 1 (
  echo.
  echo 실행에 실패했습니다.
  echo - 한글 폰트가 없으면 워드클라우드가 실패할 수 있습니다. VOC_FONT_PATH에 TTF를 지정하세요.
  echo - 패키지가 없으면 setup.bat을 실행하세요.
  echo - 포트 7860이 사용 중이면 해당 프로그램을 종료하세요.
  pause
)

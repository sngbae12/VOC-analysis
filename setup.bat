@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo [고객사 VOC 분석] 설치를 시작합니다.
echo 프로젝트 폴더: %CD%

where python >nul 2>&1
if errorlevel 1 (
  echo Python을 찾지 못했습니다. Python 3.11을 설치한 뒤 PATH에 추가하고 다시 실행하세요.
  echo https://www.python.org/downloads/
  pause
  exit /b 1
)

python -c "import sys; v=sys.version_info; raise SystemExit(0 if (v.major, v.minor) in {(3,11),(3,12),(3,13)} else 1)"
if errorlevel 1 (
  echo.
  python -c "import sys; print('현재 Python', sys.version)"
  echo 이 앱은 Python 3.11, 3.12, 3.13에서만 설치할 수 있습니다. Python 3.11을 권장합니다.
  echo 고정된 CrewAI 1.15.1은 Python 3.14를 허용하지 않고, pandas 3.0은 3.11 이상이 필요합니다.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo 가상환경 .venv 를 만듭니다.
  python -m venv .venv
  if errorlevel 1 (
    echo 가상환경을 만들지 못했습니다. 폴더 쓰기 권한과 Python 설치를 확인하세요.
    pause
    exit /b 1
  )
)

set "VPY=.venv\Scripts\python.exe"
echo pip을 준비합니다.
"%VPY%" -m pip install -U pip
if errorlevel 1 (
  echo pip 업그레이드에 실패했습니다. 인터넷 연결과 방화벽을 확인하세요.
  pause
  exit /b 1
)

echo requirements.txt 패키지를 설치합니다.
"%VPY%" -m pip install -r requirements.txt
if errorlevel 1 (
  echo 패키지 설치에 실패했습니다.
  echo - Python 버전이 3.11~3.13인지 확인하세요.
  echo - 한글·공백이 있는 경로에서도 이 스크립트는 프로젝트 폴더 기준으로 동작합니다.
  echo - 인터넷이 되거나 사내 파이 미러를 사용할 수 있는지 확인하세요.
  pause
  exit /b 1
)

"%VPY%" -m pip check
if errorlevel 1 (
  echo 설치된 패키지 의존성에 문제가 있습니다. 위 pip check 출력을 확인하세요.
  pause
  exit /b 1
)

echo.
echo 설치가 끝났습니다. start.bat 으로 실행하세요.
pause

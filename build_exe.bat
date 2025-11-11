@echo off
chcp 65001 >nul
echo ========================================
echo FileRenamerX 실행 파일 빌드
echo ========================================
echo.

REM PyInstaller 설치 확인
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller가 설치되어 있지 않습니다. 설치 중...
    pip install pyinstaller
    if errorlevel 1 (
        echo PyInstaller 설치 실패!
        pause
        exit /b 1
    )
)

echo.
echo 실행 파일 빌드 중...
echo.

REM 기존 빌드 폴더 삭제
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist FileRenamerX.spec del /q FileRenamerX.spec

REM PyInstaller로 실행 파일 생성
pyinstaller --name=FileRenamerX ^
    --onefile ^
    --windowed ^
    --icon=NONE ^
    --add-data "prompt.txt;." ^
    --hidden-import=tkinter ^
    --hidden-import=PIL ^
    --hidden-import=cv2 ^
    --hidden-import=numpy ^
    --hidden-import=google.cloud.vision ^
    --hidden-import=google.oauth2 ^
    --hidden-import=openai ^
    --hidden-import=google.generativeai ^
    --hidden-import=psutil ^
    --hidden-import=openpyxl ^
    --collect-all=google.cloud.vision ^
    --collect-all=google.oauth2 ^
    --collect-all=openai ^
    --collect-all=google.generativeai ^
    --exclude-module=matplotlib ^
    --exclude-module=scipy ^
    --exclude-module=pandas ^
    run.py

if errorlevel 1 (
    echo.
    echo 빌드 실패!
    pause
    exit /b 1
)

echo.
echo ========================================
echo 빌드 완료!
echo ========================================
echo.
echo 실행 파일 위치: dist\FileRenamerX.exe
echo.
echo 배포 시 다음 파일들을 함께 배포해야 합니다:
echo   - dist\FileRenamerX.exe
echo   - vision-api-key\ 폴더 (API 키 파일 포함)
echo   - chatgpt-api-key\ 폴더 (API 키 파일 포함)
echo   - prompt.txt
echo.
pause



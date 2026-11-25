@echo off
chcp 65001 >nul
setlocal

where python >nul 2>&1
if %errorlevel%==0 (
    python "%~dp0video_prep.py" app
    goto :done
)

where py >nul 2>&1
if %errorlevel%==0 (
    py -3 "%~dp0video_prep.py" app
    goto :done
)

echo.
echo [오류] Python 3가 설치되어 있지 않거나 PATH에 등록되지 않았습니다.
echo PowerShell에서 아래 명령어를 실행한 뒤 이 파일을 다시 실행하세요:
echo.
echo   winget install -e --id Python.Python.3.12
echo   winget install -e --id Gyan.FFmpeg
echo   winget install -e --id Git.Git
echo.
pause
exit /b 1

:done
if %errorlevel% neq 0 (
    echo.
    echo [안내] 오류가 발생했습니다. 위 메시지를 확인해 주세요.
    pause
)

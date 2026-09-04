@echo off
setlocal
rem Упаковка dist\TurNalog в zip для GitHub Releases.
rem Использование: make_release.bat   (после build.bat)
set ZIP=TurNalog-win64.zip
if not exist dist\TurNalog\TurNalog.exe (
    echo [ERROR] dist\TurNalog\TurNalog.exe not found. Run build.bat first.
    pause
    exit /b 1
)
powershell -NoProfile -Command "if (Test-Path '%ZIP%') { Remove-Item '%ZIP%' }; Compress-Archive -Path 'dist\TurNalog' -DestinationPath '%ZIP%'"
echo.
echo ============================================
echo   DONE: %ZIP%
echo   GitHub -^> Releases -^> Draft new release -^>
echo   tag (v2.3) + attach %ZIP% -^> Publish
echo ============================================
pause

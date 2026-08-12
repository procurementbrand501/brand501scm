@echo off
chcp 65001 >nul
echo Brand501 SCM 대시보드 업데이트 시작...
echo.

cd /d C:\scm
python app.py
if errorlevel 1 (
    echo.
    echo ❌ app.py 실행 중 오류가 발생했습니다. dashboard.html을 재배포하지 않고 중단합니다.
    pause
    exit /b 1
)

if not exist dashboard.html (
    echo.
    echo ❌ dashboard.html이 생성되지 않았습니다. 중단합니다.
    pause
    exit /b 1
)

echo.
echo GitHub 업로드 중...
git add dashboard.html
git commit -m "대시보드 업데이트 %date% %time%"
if errorlevel 1 (
    echo.
    echo ⚠️ 커밋할 변경사항이 없거나 커밋에 실패했습니다. push를 건너뜁니다.
    pause
    exit /b 0
)

git push
if errorlevel 1 (
    echo.
    echo ❌ git push에 실패했습니다. 네트워크/인증 상태를 확인하세요.
    pause
    exit /b 1
)

echo.
echo ✅ 완료! 1~2분 후 새로고침하면 반영됩니다.
pause
@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ============================================
echo   모닝프레임 웹 프로그램 시작
echo ============================================
echo.
echo   브라우저에서 http://localhost:8080 접속
echo   종료: 이 창에서 Ctrl+C
echo.

pip install -q -r requirements.txt
start http://localhost:8080
python -m uvicorn web.app:app --host 0.0.0.0 --port 8080 --reload

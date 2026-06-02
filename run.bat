@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo ============================================
echo   모닝프레임 주문/생산/월결산 자동 처리
echo ============================================
echo.
echo [1/2] Python 처리 중... (약 10~20초)
echo.

python -m src.main all
if errorlevel 1 (
    echo.
    echo [오류] 처리 실패. 아래를 확인하세요:
    echo   - Python 설치 여부
    echo   - pip install -r requirements.txt 실행 여부
    echo   - output\logs 폴더의 로그 파일
    echo.
    pause
    exit /b 1
)

echo.
echo [2/2] 완료!
echo.
echo   결과 위치:
echo     생산지시서  - output\production\
echo     월결산표    - output\settlement\
echo     단가검증    - output\validation\
echo     주문 이미지 - output\images\
echo.
echo   (생산지시서, 월결산 엑셀이 자동으로 열립니다)
echo.
pause

@echo off
REM A-1 조서 자동 생성 - 더블클릭 실행
REM 사용 전: [입력자료] 폴더에 파일 3개를 넣어두세요
REM   - 발송 Control Sheet
REM   - 거래처원장(.xls/.xlsx) 또는 결산보고서(.xlsx)
REM   - 조회서 취합엑셀
REM 템플릿은 [양식자료] 폴더에 두면 자동으로 찾습니다.
REM 결과는 [출력조서] 폴더에 생성됩니다. ([변환자료]는 자동 - 건드리지 마세요)

chcp 65001 >nul
cd /d "%~dp0"

python "src\easy_run.py"
if errorlevel 1 (
  echo.
  echo [문제가 발생했습니다. 위 메시지를 확인하세요.]
)

echo.
pause

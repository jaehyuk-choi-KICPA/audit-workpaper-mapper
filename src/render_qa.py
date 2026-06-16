"""개발용 시각 QA 헬퍼 — 생성된 시트를 PDF로 렌더해 사람(또는 멀티모달 LLM)이 눈으로 확인.

용도: 생성기/템플릿/config를 **신규 작성·수정**할 때, 해당 시트 1장을 PDF로 뽑아
양식 깨짐(격자·줄·너비·순서·제목·정렬 등 프로그래밍 검사가 못 잡는 것)을 시각 확인한다.
안정화된 생성기를 새 데이터로 돌리는 production 단계에서는 불필요(프로그래밍 검사로 충분).

배포 EXE에는 포함되지 않는 **개발 도구**다 (Excel 설치 필요; 사용자는 Excel로 직접 봄).

사용 예:
    python src/render_qa.py "_internal/output/A-1.xlsx" "Control Sheet"
    → 같은 폴더에 "<파일명>__<시트>.pdf" 생성. 그 PDF를 열어/읽어 시각 확인.
"""

import subprocess
import sys
from pathlib import Path

# Excel COM(PowerShell)로 단일 시트를 한 페이지 폭에 맞춰 PDF로 내보낸다.
_PS = r'''
$ErrorActionPreference = "Stop"
$xlsx = "{xlsx}"
$sheet = "{sheet}"
$pdf = "{pdf}"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
try {{
  $wb = $excel.Workbooks.Open($xlsx)
  $ws = $wb.Sheets.Item($sheet)
  $ws.PageSetup.Orientation = 2          # 가로
  $ws.PageSetup.Zoom = $false
  $ws.PageSetup.FitToPagesWide = 1
  $ws.PageSetup.FitToPagesTall = $false
  $ws.ExportAsFixedFormat(0, $pdf)       # 0 = xlTypePDF (활성/해당 시트)
  $wb.Close($false)
}} finally {{
  $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}}
'''


def render_sheet_to_pdf(xlsx_path: str, sheet_name: str,
                        pdf_path: "str | None" = None) -> str:
    """xlsx의 한 시트를 PDF로 렌더해 경로를 반환한다 (Excel COM 사용).

    Raises:
        RuntimeError: Excel 렌더 실패(미설치 등)
    """
    xlsx = Path(xlsx_path).resolve()
    pdf = Path(pdf_path).resolve() if pdf_path else \
        xlsx.with_name(f"{xlsx.stem}__{sheet_name.replace(' ', '_')}.pdf")
    script = _PS.format(xlsx=str(xlsx), sheet=sheet_name, pdf=str(pdf))
    r = subprocess.run(["powershell", "-NoProfile", "-Command", script],
                       capture_output=True, text=True)
    if not pdf.exists():
        raise RuntimeError(f"PDF 렌더 실패(Excel 필요): {r.stderr[-400:]}")
    return str(pdf)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("사용법: python src/render_qa.py <xlsx> <시트명> [pdf경로]")
        raise SystemExit(2)
    out = render_sheet_to_pdf(sys.argv[1], sys.argv[2],
                              sys.argv[3] if len(sys.argv) > 3 else None)
    print(f"렌더 완료 → {out}")

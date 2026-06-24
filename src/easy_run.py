# -*- coding: utf-8 -*-
"""A-1 조서 생성 — 쉬운 실행기 (비전문가용 대화형).

사용법:
  1) `_internal/입력자료/` 폴더에 파일 3개를 넣는다:
       · 발송 Control Sheet
       · 거래처원장(.xls/.xlsx) 또는 결산보고서(.xlsx)
       · 조회서 취합엑셀 (금융기관조회서_..기말감사..xlsx)
  2) 루트의 `A-1생성.bat` 더블클릭 (또는 python src/easy_run.py)
  3) 화면 안내대로 회사명·날짜·작성자를 입력
  4) `_internal/output/` 에 결과 파일 생성

경로·플래그를 몰라도 된다. 파일 위치만 맞춰주면 나머지는 자동.
"""

import glob
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")
except Exception:
    pass

import openpyxl

from run_a1 import discover_inputs
from pipeline import build_a1

def _base_dir() -> Path:
    """프로그램 루트. .exe(frozen)면 실행파일 폴더, 스크립트면 프로젝트 루트."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE = _base_dir()

# 사용자-대면 5폴더 = 배포 루트 최상위. 환경변수(A1_*_DIR)로 재정의 가능.
INPUT_DIR    = Path(os.environ.get("A1_INPUT_DIR",    BASE / "입력자료"))
PARSED_DIR   = Path(os.environ.get("A1_PARSED_DIR",   BASE / "변환자료"))
TEMPLATE_DIR = Path(os.environ.get("A1_TEMPLATE_DIR", BASE / "양식자료"))
REF_DIR      = Path(os.environ.get("A1_REF_DIR",      BASE / "참고자료"))
OUTPUT_DIR   = Path(os.environ.get("A1_OUTPUT_DIR",   BASE / "출력조서"))

# config(셀 매핑) = 프로그램 내부 자산. 스크립트=_internal/config, .exe=내장(_MEIPASS).
CONFIG_DIR = Path(getattr(sys, "_MEIPASS", str(BASE))) / "_internal" / "config"

# A-1 양식이 가져야 할 시트(이 시트들이 있으면 A-1 템플릿으로 인식 — 파일명 무관)
_A1_SHEET_MARKERS = ("A100_금융기관 조회서 완전성체크",
                     "A200_조회서 대사",
                     "A300_조회서_기타조회내역")


def _line(msg=""):
    print(msg, flush=True)


def _ask(label: str, required: bool = True, example: str = "") -> str:
    hint = f" (예: {example})" if example else ""
    while True:
        val = input(f"  {label}{hint}: ").strip()
        if val or not required:
            return val
        _line("  ▷ 값을 입력해 주세요.")


def _find_template() -> "str | None":
    """템플릿 폴더에서 A-1 시트 구조를 가진 엑셀(.xlsx/.xlsm/.xls)을 동적으로 찾는다(파일명 무관)."""
    candidates = [h for h in glob.glob(str(TEMPLATE_DIR / "**" / "*.xls*"), recursive=True)
                  if "~$" not in h]
    for path in sorted(candidates):
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            names = set(wb.sheetnames)
            wb.close()
        except Exception:
            continue
        if all(m in names for m in _A1_SHEET_MARKERS):
            return path
    return None


def main():
    _line("=" * 56)
    _line("  A-1 (은행·금융기관 조회) 조서 자동 생성")
    _line("=" * 56)

    for d in (INPUT_DIR, PARSED_DIR, TEMPLATE_DIR, REF_DIR, OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)

    # 1) 입력 파일 자동 탐색
    found = discover_inputs(str(INPUT_DIR))
    control, confirm, ledger = found["control_sheet"], found["confirm_xlsx"], found["ledger_src"]

    labels = [("발송 Control Sheet", control),
              ("조회서 취합엑셀", confirm),
              ("거래처원장/결산보고서", ledger)]
    missing = [n for n, v in labels if not v]
    if missing:
        _line("\n[안내] 다음 파일을 찾지 못했습니다: " + ", ".join(missing))
        _line(f"       아래 폴더에 파일을 넣고 다시 실행해 주세요:")
        _line(f"       {INPUT_DIR}")
        return 2

    _line("\n[입력 파일 확인]")
    for n, v in labels:
        kind = ""
        if n.startswith("거래처원장"):
            kind = " (결산보고서)" if "결산보고서" in Path(v).name else " (거래처원장)"
        _line(f"  · {n}: {Path(v).name}{kind}")

    aux_a1 = found.get("aux_a1")
    if aux_a1:
        _line(f"  · 보조자료_A1(개별 잔액명세서): {Path(aux_a1).name} (확인됨)")

    # 분개장: 있으면 인식만(A-100 완전성 검증용 — 추후 단계). 현재 A-1 생성엔 미사용.
    journal = next((h for h in glob.glob(str(INPUT_DIR / "**" / "*분개장*.xls*"), recursive=True)
                    if "~$" not in h), None)
    if journal:
        _line(f"  · 분개장: {Path(journal).name} (확인됨 — 완전성 검증용, 현재 A-1 생성엔 미사용)")

    template = _find_template()
    if not template:
        _line(f"\n[안내] 양식 템플릿(A-1 시트 포함 .xlsx)을 찾지 못했습니다. 다음 폴더에 템플릿을 두세요:")
        _line(f"       {TEMPLATE_DIR}")
        return 2
    _line(f"\n[양식] {Path(template).name}")

    # 2) 회사 정보 입력
    _line("\n[회사 정보 입력]")
    company = _ask("회사명", example="주식회사 OO")
    date = _ask("기준일", example="2025-12-31")
    preparer = _ask("작성자(Preparer)", example="CJH")
    reviewer = _ask("검토자(Reviewer)", required=False, example="KHK")

    output = OUTPUT_DIR / f"{company}_A-1.xlsx"

    # 3) 생성
    _line("\n생성 중...")
    try:
        report = build_a1(
            control_sheet=control, ledger_src=ledger, confirm_xlsx=confirm,
            template=template, config_dir=str(CONFIG_DIR), output=str(output),
            parsed_dir=str(PARSED_DIR), ref_dir=str(REF_DIR),
            params={"회사명": company, "날짜": date,
                    "preparer": preparer, "reviewer": reviewer},
        )
    except PermissionError:
        _line(f"\n[오류] 결과 파일이 열려 있어 저장할 수 없습니다. 닫고 다시 실행해 주세요:")
        _line(f"       {output}")
        return 1
    except Exception as e:
        _line(f"\n[오류] 예기치 못한 문제로 생성이 중단됐습니다: {type(e).__name__}: {e}")
        return 1

    # 진단 리포트: 화면 출력 + 파일 저장(엑셀을 못 보는 사용자가 무엇이 문제인지 알도록)
    _line("\n" + report.render())
    report_path = OUTPUT_DIR / f"{company}_실행리포트.txt"
    try:
        report_path.write_text(report.render(), encoding="utf-8")
    except Exception:
        report_path = None

    _line("\n" + "=" * 60)
    _line(f"  결과 파일 : {output}")
    if report_path:
        _line(f"  실행 리포트: {report_path}")
    _line("=" * 60)
    _line("  ※ 초안입니다. 위 [확인필요]/[출력] 항목과, 은행연합회·우편회신·일치여부 열은 감사인이 직접 확인.")
    return 0


if __name__ == "__main__":
    rc = main()
    # 결과/리포트를 볼 수 있도록 창을 항상 유지 (exe 직접 실행해도 즉시 안 닫힘)
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)

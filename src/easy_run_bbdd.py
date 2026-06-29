# -*- coding: utf-8 -*-
"""BBDD(장단기차입금) 조서 단독 생성 — 쉬운 실행기 (비전문가용 대화형, 전용 EXE).

전체조서 EXE(easy_run_lead)는 '정산표→총괄표' 서두 매핑만 한다. 이 실행기는 BBDD 한 조서를
**1.총괄표 + 상세(조회 대사·미지급이자·주석·RECAP·이자비용·조회서기타)까지 완성본으로** 생성한다.
공유 파이프라인(build_lead_all)을 건드리지 않으므로 타 조서 회귀 위험이 없다.

입력자료(파일명 접두/포함으로 자동 인식):
  · *정산표*           → 1.총괄표 leaf (필수)
  · 필수자료2_조회서    → 조회 대사·명세·담보·조회서기타 (BANK 2-2/9/1)
  · 필수자료3_잔액      → 200 좌측 split·100 RECAP (거래처원장 이상양식)
  · *분개장*           → 100 이자비용 (이상적 분개장)
조회서/잔액/분개장이 없으면 그 부분만 비우고 나머지는 생성(부분 실패 허용).

자체 검증: 생성본의 check 셀(200 합계=100!E16, RECAP 기초합=leaf, 이자비용 합=leaf E20)로 타이아웃 확인.
"""

import glob
import os
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", message=".*[Hh]eader or footer.*")
sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")
except Exception:
    pass

from bbdd_pipeline import build_bbdd


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE = _base_dir()
INPUT_DIR    = Path(os.environ.get("A1_INPUT_DIR",    BASE / "입력자료"))
PARSED_DIR   = Path(os.environ.get("A1_PARSED_DIR",   BASE / "변환자료"))
TEMPLATE_DIR = Path(os.environ.get("A1_TEMPLATE_DIR", BASE / "양식자료"))
OUTPUT_DIR   = Path(os.environ.get("A1_OUTPUT_DIR",   BASE / "출력조서"))
CONFIG_DIR = Path(getattr(sys, "_MEIPASS", str(BASE))) / "_internal" / "config"
if not getattr(sys, "frozen", False) and not list(TEMPLATE_DIR.glob("*.xls*")):
    _dev = BASE / "_internal" / "양식"
    if _dev.exists():
        TEMPLATE_DIR = _dev


def _line(msg=""):
    print(msg, flush=True)


def _ask(label, required=True, example=""):
    hint = f" (예: {example})" if example else ""
    while True:
        val = input(f"  {label}{hint}: ").strip()
        if val or not required:
            return val
        _line("  ▷ 값을 입력해 주세요.")


def _find(*patterns):
    """입력자료에서 패턴 중 처음 매칭되는 파일(임시파일 제외)."""
    for pat in patterns:
        cands = [h for h in glob.glob(str(INPUT_DIR / "**" / pat), recursive=True)
                 if "~$" not in h]
        if cands:
            return sorted(cands)[0]
    return None


def _find_template():
    """양식자료(or _internal/양식)에서 BBDD 완성본 템플릿."""
    cands = [h for h in glob.glob(str(TEMPLATE_DIR / "**" / "*BBDD*.xls*"), recursive=True)
             if "~$" not in h]
    return sorted(cands)[0] if cands else None


def main():
    _line("=" * 56)
    _line("  BBDD 장단기차입금 조서 단독 생성")
    _line("=" * 56)
    for d in (INPUT_DIR, PARSED_DIR, TEMPLATE_DIR, OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)

    settlement = _find("*정산표*.xls*")
    template = _find_template()
    if not settlement:
        _line(f"\n[안내] 정산표를 찾지 못했습니다(파일명에 '정산표' 포함). 폴더: {INPUT_DIR}")
        return 2
    if not template:
        _line(f"\n[안내] BBDD 템플릿을 찾지 못했습니다(파일명에 'BBDD' 포함). 폴더: {TEMPLATE_DIR}")
        return 2

    confirm = _find("필수자료2*", "*조회서*.xls*")
    ledger  = _find("필수자료3*", "*잔액*.xls*")
    journal = _find("*분개장*.xls*")

    _line(f"\n[정산표] {Path(settlement).name}")
    _line(f"[템플릿] {Path(template).name}")
    _line(f"[조회서] {Path(confirm).name if confirm else '— (조회 대사·명세·담보·조회서기타 비움)'}")
    _line(f"[잔액]   {Path(ledger).name if ledger else '— (200 좌측·RECAP 비움)'}")
    _line(f"[분개장] {Path(journal).name if journal else '— (이자비용 비움)'}")

    _line("\n[회사 정보 입력]")
    company  = _ask("회사명", example="주식회사 OO")
    date     = _ask("기준일", example="2025-12-31")
    preparer = _ask("작성자(Preparer)", example="CJH")
    reviewer = _ask("검토자(Reviewer)", required=False, example="KHK")

    out_dir = OUTPUT_DIR / company
    out_dir.mkdir(parents=True, exist_ok=True)
    output = out_dir / f"BBDD_장단기차입금_{company}{Path(template).suffix}"

    _line("\n생성 중... (조회서·잔액·분개장 파싱 + 5시트 이식, 수십 초 소요)\n")
    try:
        r = build_bbdd(
            template, confirm=confirm, ledger=ledger, settlement=settlement, journal=journal,
            config_dir=str(CONFIG_DIR), parsed_dir=str(PARSED_DIR / company), output=str(output),
            params={"회사명": company, "날짜": date, "preparer": preparer, "reviewer": reviewer},
        )
    except PermissionError:
        _line(f"\n[오류] 출력 파일이 열려 있어 저장할 수 없습니다. 닫고 다시 실행해 주세요:\n       {output}")
        return 1
    except Exception as e:
        _line(f"\n[오류] 생성이 중단됐습니다: {type(e).__name__}: {e}")
        return 1

    # 진단 요약
    s = r.get("sheets", {})
    def _g(k, *path):
        v = s.get(k) or {}
        for p in path:
            v = (v or {}).get(p) if isinstance(v, dict) else None
        return v
    lines = ["=" * 60, "  BBDD 조서 생성 완료", "=" * 60,
             f"  100 RECAP 장기  : {(_g('100_recap','장기차입금') or {}).get('filled','-')} 시설",
             f"  100 이자비용    : 합 {format(int(_g('100_interest','이자비용합') or 0), ',')}",
             f"  200 조회/대출   : {(_g('200') or {}).get('filled','-')} 건",
             f"  300 미지급이자  : {(_g('300','t1') or {}).get('filled','-')} 건",
             f"  400 주석        : {(_g('400') or {}).get('filled','-')} 건",
             f"  조회서기타      : 대출 {(_g('iq','loans') or {}).get('filled','-')}·담보 {(_g('iq','collateral') or {}).get('filled','-')}"]
    warns = r.get("warnings") or []
    if warns:
        lines.append("")
        lines.append(f"  ⚠ 경고 {len(warns)}건(해당 부분만 비움):")
        lines += [f"      - {w}" for w in warns]
    else:
        lines.append("\n  ✓ 경고 없음.")
    summary = "\n".join(lines)
    _line("\n" + summary)

    report_path = out_dir / f"BBDD_{company}_실행리포트.txt"
    try:
        report_path.write_text(summary, encoding="utf-8")
    except Exception:
        report_path = None

    _line("\n" + "=" * 60)
    _line(f"  결과 파일 : {output}")
    if report_path:
        _line(f"  실행 리포트: {report_path}")
    _line("=" * 60)
    _line("  ※ 초안입니다. 차입/상환/대체 이동분, 미지급이자 재계산, 결론은 감사인이 보완하세요.")
    return 0


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)

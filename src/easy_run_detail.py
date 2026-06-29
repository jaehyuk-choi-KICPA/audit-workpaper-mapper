# -*- coding: utf-8 -*-
"""기타자산(D200)·기타부채(CC200) 상세 시트 생성 — 쉬운 실행기(비전문가용, 2단계 묶음 EXE).

전제: 먼저 '조서생성'(정산표 기반 총괄표 일괄 생성)을 돌려 `출력조서/{회사}/` 에 D·CC 완성본
(총괄표 채워진)이 있어야 한다. 이 EXE는 그 산출물의 **두번째 시트(상세)**를 거래처원장(잔액명세서)
기반으로 채운다.

사용법:
  1) `입력자료/` 에 거래처원장(잔액명세서) 엑셀을 둔다(파일명에 '원장' 또는 '잔액' 포함, 또는 필수자료3 토큰).
  2) `출력조서/{회사}/` 에 1단계 산출물(D_4000_…·CC_4000_…)이 있어야 한다.
  3) 루트의 `기타자산부채_상세생성.bat` 더블클릭 (또는 python src/easy_run_detail.py)
  4) `출력조서/{회사}/` 의 D·CC 완성본 두번째 시트가 채워진다(제자리 갱신, 보조시트·컨트롤 보존).

경로·플래그를 몰라도 된다. 원장만 입력자료에 두고, 1단계를 먼저 돌려두면 나머지는 자동.
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

from detail_pipeline import build_detail_all
from pipeline import build_lead_all
from run_lead_all import REGISTRY
from workspace import (select_workspace, config_dir, template_dir,
                       prompt_company_info)

# config·양식 = 회사 무관 공유 자산. 입력·변환·출력은 회사별 워크스페이스에서 받는다.
CONFIG_DIR = config_dir()
TEMPLATE_DIR = template_dir()
INPUT_DIR = PARSED_DIR = OUTPUT_DIR = None  # main()에서 워크스페이스로 설정

_LEDGER_PATTERNS = ("필수자료3*.xls*", "*거래처*원장*.xls*", "*거래처*잔액*.xls*",
                    "*잔액*현황*.xls*", "*원장*.xls*", "*잔액*.xls*")

# D·CC 총괄표를 자체 생성하기 위한 레지스트리(조서생성의 D·CC 부분 재사용 → 독립 실행).
_DCC = [it for it in REGISTRY if it["code"] in ("D", "CC")]


def _line(msg=""):
    print(msg, flush=True)


def _find_ledger():
    for pat in _LEDGER_PATTERNS:
        cands = [h for h in glob.glob(str(INPUT_DIR / "**" / pat), recursive=True)
                 if "~$" not in h and "결산보고서" not in Path(h).name]
        if cands:
            return sorted(cands)[0]
    return None


def _find_settlement():
    cands = [h for h in glob.glob(str(INPUT_DIR / "**" / "*정산표*.xls*"), recursive=True)
             if "~$" not in h]
    return sorted(cands)[0] if cands else None


def _has_dcc(out_dir: Path) -> bool:
    """출력조서에 1단계 D·CC 산출물이 있는지 확인."""
    return bool(glob.glob(str(out_dir / "**" / "D_*기타자산*.xls*"), recursive=True) or
                glob.glob(str(out_dir / "**" / "CC_*기타부채*.xls*"), recursive=True) or
                glob.glob(str(out_dir / "**" / "D_4000_*.xls*"), recursive=True))


def main():
    _line("=" * 56)
    _line("  기타자산(D200)·기타부채(CC200) 상세 시트 생성")
    _line("=" * 56)

    global INPUT_DIR, PARSED_DIR, OUTPUT_DIR
    ws = select_workspace(line=_line)
    if ws is None:
        return 2
    INPUT_DIR, PARSED_DIR, OUTPUT_DIR = ws.input_dir, ws.parsed_dir, ws.output_dir
    _line(f"\n[회사] {ws.company}   (작업폴더: {ws.input_dir.parent})")

    ledger = _find_ledger()
    if ledger:
        _line(f"\n[거래처원장] {Path(ledger).name}")
    else:
        _line("\n[안내] 거래처원장(잔액명세서)을 찾지 못했습니다(파일명에 '원장'/'잔액' 포함).")
        _line(f"       없으면 거래처별 명세는 비우고 tie/check만 채웁니다. 폴더: {INPUT_DIR}")

    # D·CC 완성본(총괄표)이 없으면 → 정산표로 직접 선생성(독립 실행). 있으면 그대로 사용.
    if not _has_dcc(ws.output_dir):
        settlement = _find_settlement()
        base = [it for it in _DCC if (TEMPLATE_DIR / it["template"]).exists()]
        if settlement and base:
            _line("\n[안내] D·CC 총괄표가 없어 정산표로 먼저 생성합니다(독립 실행).")
            date, preparer, reviewer = prompt_company_info(ws, input, _line)
            try:
                build_lead_all(
                    settlement=settlement, registry=base, config_dir=str(CONFIG_DIR),
                    template_root=str(TEMPLATE_DIR), output_dir=str(ws.output_dir),
                    parsed_dir=str(ws.parsed_dir),
                    params={"회사명": ws.company, "날짜": date,
                            "preparer": preparer, "reviewer": reviewer},
                    progress=lambda code, ok: _line("   [%s] %s 총괄표" % ("V" if ok else "X", code)),
                )
            except Exception as e:
                _line(f"\n[오류] D·CC 총괄표 생성 실패: {type(e).__name__}: {e}")
                return 1
        if not _has_dcc(ws.output_dir):
            _line("\n[안내] D·CC 총괄표를 만들 수 없습니다.")
            _line(f"       정산표를 입력자료에 넣거나, 먼저 '조서생성'을 돌리세요. 폴더: {ws.output_dir}")
            return 2

    all_reports = {}
    _line(f"── {ws.company} 생성 중...")

    def _progress(code, ok):
        _line("   [%s] %s" % ("V" if ok else "X", code))

    try:
        reports = build_detail_all(
            output_root=str(ws.output_dir), ledger_path=ledger,
            config_dir=str(CONFIG_DIR),
            parsed_dir=str(ws.parsed_dir),
            progress=_progress)
        all_reports[ws.company] = reports
    except PermissionError:
        _line(f"\n[오류] 출력 파일이 열려 있어 저장할 수 없습니다. 닫고 다시 실행하세요: {ws.output_dir}")
        return 1
    except Exception as e:
        _line(f"\n[오류] {ws.company}: {type(e).__name__}: {e}")

    # 진단 요약
    lines = ["", "=" * 60, "  상세 생성 완료", "=" * 60]
    for cname, reports in all_reports.items():
        lines.append(f"[{cname}]")
        for code, r in reports.items():
            flag = "✗" if r.has_error else ("△" if r.has_warn else "✓")
            lines.append(f"  {flag} {code}")
            for st, lv, m in r.entries:
                if lv in ("warn", "error"):
                    lines.append(f"      - [{st}] {m}")
    summary = "\n".join(lines)
    _line(summary)
    try:
        rp = OUTPUT_DIR / "상세생성_실행리포트.txt"
        rp.write_text(summary, encoding="utf-8")
        _line(f"\n  실행 리포트: {rp}")
    except Exception:
        pass
    _line("  ※ 초안입니다. 2)증감분석 서술·3)Test 샘플링·Nature는 감사인이 검토·작성하세요.")
    return 0


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)

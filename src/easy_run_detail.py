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


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE = _base_dir()
INPUT_DIR  = Path(os.environ.get("A1_INPUT_DIR",  BASE / "입력자료"))
PARSED_DIR = Path(os.environ.get("A1_PARSED_DIR", BASE / "변환자료"))
OUTPUT_DIR = Path(os.environ.get("A1_OUTPUT_DIR", BASE / "출력조서"))
CONFIG_DIR = Path(getattr(sys, "_MEIPASS", str(BASE))) / "_internal" / "config"
# 개발(스크립트): 출력조서가 비면 _internal/output 사용
if not getattr(sys, "frozen", False) and not list(OUTPUT_DIR.glob("*/")):
    _dev = BASE / "_internal" / "output"
    if _dev.exists():
        OUTPUT_DIR = _dev

_LEDGER_PATTERNS = ("필수자료3*.xls*", "*거래처*원장*.xls*", "*거래처*잔액*.xls*",
                    "*잔액*현황*.xls*", "*원장*.xls*", "*잔액*.xls*")


def _line(msg=""):
    print(msg, flush=True)


def _find_ledger():
    for pat in _LEDGER_PATTERNS:
        cands = [h for h in glob.glob(str(INPUT_DIR / "**" / pat), recursive=True)
                 if "~$" not in h and "결산보고서" not in Path(h).name]
        if cands:
            return sorted(cands)[0]
    return None


def _company_dirs():
    """출력조서 하위에서 D·CC 산출물이 있는 회사 폴더 목록."""
    out = []
    for d in sorted(p for p in OUTPUT_DIR.iterdir() if p.is_dir()):
        has = glob.glob(str(d / "**" / "D_*기타자산*.xls*"), recursive=True) or \
              glob.glob(str(d / "**" / "CC_*기타부채*.xls*"), recursive=True) or \
              glob.glob(str(d / "**" / "D_4000_*.xls*"), recursive=True)
        if has:
            out.append(d)
    return out


def main():
    _line("=" * 56)
    _line("  기타자산(D200)·기타부채(CC200) 상세 시트 생성")
    _line("=" * 56)
    for d in (INPUT_DIR, PARSED_DIR, OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)

    ledger = _find_ledger()
    if ledger:
        _line(f"\n[거래처원장] {Path(ledger).name}")
    else:
        _line("\n[안내] 거래처원장(잔액명세서)을 찾지 못했습니다(파일명에 '원장'/'잔액' 포함).")
        _line(f"       없으면 거래처별 명세는 비우고 tie/check만 채웁니다. 폴더: {INPUT_DIR}")

    comps = _company_dirs()
    if not comps:
        _line("\n[안내] 출력조서에서 1단계 산출물(D·CC 완성본)을 찾지 못했습니다.")
        _line(f"       먼저 '조서생성'을 돌려 D·CC 총괄표를 만든 뒤 실행하세요. 폴더: {OUTPUT_DIR}")
        return 2
    _line(f"[대상 회사] {', '.join(p.name for p in comps)}\n")

    all_reports = {}
    for comp in comps:
        _line(f"── {comp.name} 생성 중...")

        def _progress(code, ok):
            _line("   [%s] %s" % ("V" if ok else "X", code))

        try:
            reports = build_detail_all(
                output_root=str(comp), ledger_path=ledger,
                config_dir=str(CONFIG_DIR),
                parsed_dir=str(PARSED_DIR / comp.name),
                progress=_progress)
            all_reports[comp.name] = reports
        except PermissionError:
            _line(f"\n[오류] 출력 파일이 열려 있어 저장할 수 없습니다. 닫고 다시 실행하세요: {comp}")
            return 1
        except Exception as e:
            _line(f"\n[오류] {comp.name}: {type(e).__name__}: {e}")

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

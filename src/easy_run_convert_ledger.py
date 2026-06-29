# -*- coding: utf-8 -*-
"""잔액명세서(거래처원장·계정과목별 거래처잔액) → 이상적 양식 변환 — 폴더 자동감지 EXE.

사용법: 이 EXE를 잔액명세서 파일들이 있는 폴더(또는 상위 폴더)에 두고 더블클릭.
  - 폴더(하위 포함)에서 잔액명세서/원장 파일을 자동 감지해 각각 이상적 양식으로 변환한다.
  - 결과: 같은 폴더의 `변환결과/{원본명}__이상적잔액명세서.xlsx` (계정과목·거래처별 기초·기말 표준표).
  - 이미 변환된 산출물(__이상적)·임시파일은 건너뛴다.

지원 형식: 계정별 다중시트((코드)계정명·N_계정명(코드)) / 단일 통합시트(계정과목 컬럼) /
           거래처별 블록(거래처 헤더 + 계정행). .xls/.xlsx/.xlsm 모두 인식. (오프라인·API 불필요)
"""

import glob
import os
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from extractors import parse_ledger, write_ideal_ledger

# 잔액명세서 감지 토큰(파일명 부분일치). 분개장·정산표·조회서 등은 제외.
_INCLUDE = ("원장", "잔액", "거래처", "계정과목별")
_EXCLUDE = ("분개", "정산표", "조회서", "__이상적", "이상적양식", "결산보고서")
_OUT_DIR = "변환결과"


def _line(m=""):
    print(m, flush=True)


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _discover(root: Path):
    hits = []
    for f in glob.glob(str(root / "**" / "*.xls*"), recursive=True):
        name = Path(f).name
        if name.startswith("~$") or _OUT_DIR in Path(f).parts:
            continue
        low = name
        if any(k in low for k in _EXCLUDE):
            continue
        if any(k in low for k in _INCLUDE):
            hits.append(f)
    return sorted(set(hits))


def main():
    base = _base_dir()
    _line("=" * 56)
    _line("  잔액명세서 → 이상적 양식 변환 (폴더 자동감지)")
    _line("=" * 56)
    _line(f"\n[검색 폴더] {base}")
    files = _discover(base)
    if not files:
        _line("\n[안내] 잔액명세서 파일을 찾지 못했습니다(파일명에 '원장'/'잔액'/'계정과목별' 포함).")
        _line("       이 프로그램과 같은 폴더(또는 하위)에 파일을 두고 다시 실행하세요.")
        return 2
    _line(f"[감지] {len(files)}개 파일\n")

    out_dir = base / _OUT_DIR
    ok = fail = 0
    for f in files:
        name = Path(f).name
        try:
            recs = parse_ledger(f)
            nz = sum(1 for r in recs if str(r.get("잔액") or "").strip() not in ("", "None"))
            outp = out_dir / (Path(f).stem + "__이상적잔액명세서.xlsx")
            write_ideal_ledger(recs, str(outp))
            n_acct = len({r.get("계정과목") for r in recs})
            _line(f"  [V] {name[:48]}")
            _line(f"      → {recs.__len__()}행 · {n_acct}계정 · {outp.name}")
            ok += 1
        except Exception as e:
            _line(f"  [X] {name[:48]}  실패: {type(e).__name__}: {str(e)[:70]}")
            fail += 1

    _line("\n" + "=" * 56)
    _line(f"  완료 — 성공 {ok} · 실패 {fail}")
    if ok:
        _line(f"  결과 폴더: {out_dir}")
    _line("=" * 56)
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)

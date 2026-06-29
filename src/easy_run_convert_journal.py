# -*- coding: utf-8 -*-
"""분개장 → 이상적 분개장 양식 변환 — 폴더 자동감지 EXE.

사용법: 이 EXE를 분개장 파일들이 있는 폴더(또는 상위 폴더)에 두고 더블클릭.
  - 폴더(하위 포함)에서 분개장 파일을 자동 감지해 각각 이상적 양식으로 변환한다.
  - 결과: `변환결과/{원본명}__이상적분개장.xlsx` (날짜·번호·계정과목·거래처·적요·차변·대변).
  - 차변합=대변합 균형을 점검해 표시(추출 정확성 확인).

여러 시트 중 가장 큰(데이터) 시트를 자동 선택, .xls/.xlsx/.xlsm 인식, 이중계정·코드형 양식 흡수.
(오프라인·API 불필요)
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

from extractors import parse_journal, write_ideal_journal

_INCLUDE = ("분개", "journal", "전표")
_EXCLUDE = ("__이상적", "이상적양식", "정산표", "원장", "조회서")
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
        if any(k in name for k in _EXCLUDE):
            continue
        if any(k in name.lower() for k in _INCLUDE):
            hits.append(f)
    return sorted(set(hits))


def _amt(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def main():
    base = _base_dir()
    _line("=" * 56)
    _line("  분개장 → 이상적 양식 변환 (폴더 자동감지)")
    _line("=" * 56)
    _line(f"\n[검색 폴더] {base}")
    files = _discover(base)
    if not files:
        _line("\n[안내] 분개장 파일을 찾지 못했습니다(파일명에 '분개'/'전표'/'journal' 포함).")
        _line("       이 프로그램과 같은 폴더(또는 하위)에 파일을 두고 다시 실행하세요.")
        return 2
    _line(f"[감지] {len(files)}개 파일\n")

    out_dir = base / _OUT_DIR
    ok = fail = 0
    for f in files:
        name = Path(f).name
        try:
            rows = parse_journal(f)
            if not rows:
                _line(f"  [X] {name[:48]}  추출 0행(헤더 인식 실패 가능)")
                fail += 1
                continue
            outp = out_dir / (Path(f).stem + "__이상적분개장.xlsx")
            write_ideal_journal(rows, str(outp))
            dt = sum(_amt(r.get("차변")) for r in rows)
            ct = sum(_amt(r.get("대변")) for r in rows)
            bal = "균형 OK" if abs(dt - ct) < 1 else f"불균형(차{dt:,.0f}≠대{ct:,.0f})"
            _line(f"  [V] {name[:48]}")
            _line(f"      → {len(rows)}라인 · 차대 {bal} · {outp.name}")
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

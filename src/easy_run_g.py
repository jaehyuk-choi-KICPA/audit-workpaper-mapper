# -*- coding: utf-8 -*-
"""G 유무형자산 조서(G-0) 생성 — 쉬운 실행기 (비전문가용 대화형).

사용법:
  1) `입력자료/` 폴더에 **정산표**(별도정산표·CF(정산표)_25·분개장_25 시트 포함)와
       **고정자산 관리대장**(파일명에 '관리대장')을 넣는다.
  2) `참고자료/` 폴더에 **유무형자산 내용연수·상각방법 메모**(*내용연수*.txt)를 넣는다.
  3) `양식자료/`(또는 _internal/양식)에 G 완성본 템플릿(G_*유무형자산.xls*)을 둔다.
  4) `G조서생성.bat` 더블클릭 (또는 python src/easy_run_g.py)
  5) 회사명·기준일·작성자·검토자 입력 → `출력조서/{회사}/` 에 G 조서 생성.

생성 시트: 고정자산 관리대장_25 · G100 총괄표(이동컬럼·회계정책) · G300 감가비검토 · G200 취득/처분 Test.
(완성본에 무손실 이식 — 보조시트·컨트롤·매크로 보존. 주석·등본·손상 시트는 감사인 보완.)
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

from g_pipeline import build_g


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE = _base_dir()
INPUT_DIR    = Path(os.environ.get("A1_INPUT_DIR",    BASE / "입력자료"))
PARSED_DIR   = Path(os.environ.get("A1_PARSED_DIR",   BASE / "변환자료"))
TEMPLATE_DIR = Path(os.environ.get("A1_TEMPLATE_DIR", BASE / "양식자료"))
REF_DIR      = Path(os.environ.get("A1_REF_DIR",      BASE / "참고자료"))
OUTPUT_DIR   = Path(os.environ.get("A1_OUTPUT_DIR",   BASE / "출력조서"))
CONFIG_DIR   = Path(getattr(sys, "_MEIPASS", str(BASE))) / "_internal" / "config"
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


def _find(pattern, root):
    cands = [h for h in glob.glob(str(root / "**" / pattern), recursive=True) if "~$" not in h]
    return sorted(cands)[0] if cands else None


def _find_g_template():
    for pat in ("G_*유무형*.xls*", "G_*template*.xls*", "G[_-]*.xls*", "*유무형자산*.xls*"):
        hit = _find(pat, TEMPLATE_DIR)
        if hit:
            return hit
    return None


def main():
    _line("=" * 56)
    _line("  G 유무형자산 조서(G-0) 생성")
    _line("=" * 56)
    for d in (INPUT_DIR, PARSED_DIR, TEMPLATE_DIR, REF_DIR, OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)

    settlement = _find("*정산표*.xls*", INPUT_DIR)
    register = _find("*관리대장*.xls*", INPUT_DIR)
    template = _find_g_template()
    if not settlement:
        _line(f"\n[안내] 정산표를 찾지 못했습니다(파일명에 '정산표'). 입력자료 폴더:\n       {INPUT_DIR}")
        return 2
    if not template:
        _line(f"\n[안내] G 완성본 템플릿을 찾지 못했습니다(G_*유무형자산.xlsx). 양식자료 폴더:\n       {TEMPLATE_DIR}")
        return 2
    _line(f"\n[정산표] {Path(settlement).name}")
    _line(f"[관리대장] {Path(register).name if register else '(없음 — 관리대장_25·G300 재계산 제외)'}")
    _line(f"[양식]   {Path(template).name}")
    _line(f"[참고자료] {REF_DIR}  (유무형자산 내용연수·상각방법 메모)")

    _line("\n[회사 정보 입력]")
    company  = _ask("회사명", example="주식회사 OO")
    date     = _ask("기준일", example="2025-12-31")
    preparer = _ask("작성자(Preparer)", required=False, example="CJH")
    reviewer = _ask("검토자(Reviewer)", required=False, example="KHK")

    out_dir = OUTPUT_DIR / company
    out_file = out_dir / f"G_4000_유무형자산_{company}{Path(template).suffix}"
    _line("\n생성 중... (분개장 파싱에 십수 초 소요될 수 있습니다)\n")

    try:
        done, warn = build_g(
            settlement=settlement, register=register, ref_dir=str(REF_DIR),
            template=template, output=str(out_file),
            params={"회사명": company, "날짜": date, "preparer": preparer, "reviewer": reviewer},
            parsed_dir=str(PARSED_DIR / company), config_dir=str(CONFIG_DIR),
            progress=lambda s, ok: _line("  [V] %s" % s),
        )
    except PermissionError:
        _line(f"\n[오류] 출력 파일이 열려 있어 저장할 수 없습니다. 닫고 다시 실행해 주세요:\n       {out_dir}")
        return 1
    except Exception as e:
        _line(f"\n[오류] 예기치 못한 문제: {type(e).__name__}: {e}")
        return 1

    lines = ["=" * 60, f"  생성 완료 — 시트 {len(done)}개", "=" * 60]
    for s in done:
        lines.append(f"  ✓ {s}")
    if warn:
        lines.append("")
        lines.append(f"  ⚠ 경고 {len(warn)}건:")
        lines += [f"      - {w}" for w in warn]
    summary = "\n".join(lines)
    _line("\n" + summary)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        (out_dir / f"{company}_G실행리포트.txt").write_text(summary, encoding="utf-8")
    except Exception:
        pass
    _line("\n" + "=" * 60)
    _line(f"  결과 파일 : {out_file}")
    _line("=" * 60)
    _line("  ※ 초안입니다. 주석·등본검토·손상평가·결론은 감사인이 검토·보완하세요.")
    return 0


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)

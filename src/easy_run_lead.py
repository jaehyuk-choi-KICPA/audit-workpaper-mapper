# -*- coding: utf-8 -*-
"""정산표 기반 조서 총괄표 일괄 생성 — 쉬운 실행기 (비전문가용 대화형).

사용법:
  1) `입력자료/` 폴더에 **정산표**(별도정산표 + 수정사항집계 시트 포함 엑셀)를 넣는다.
       파일명에 '정산표'가 들어가면 자동 인식(예: 필수자료4_정산표_OO.xlsx).
  2) `양식자료/` 폴더에 16개 조서 완성본 템플릿을 둔다(코드_4000_template_계정명.xls*).
  3) 루트의 `조서생성.bat` 더블클릭 (또는 python src/easy_run_lead.py)
  4) 화면 안내대로 회사명·기준일·작성자·검토자 입력
  5) `출력조서/{회사명}/` 에 16개 조서 생성(완성본에 총괄표 채워 — 보조시트·컨트롤 보존)

경로·플래그를 몰라도 된다. 정산표와 양식만 폴더에 맞춰주면 나머지는 자동.
"""

import glob
import os
import sys
import warnings
from pathlib import Path

# openpyxl이 머리말/꼬리말을 못 읽을 때 내는 무해한 경고 숨김(결과 영향 없음).
warnings.filterwarnings("ignore", message=".*[Hh]eader or footer.*")

sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")
except Exception:
    pass

from pipeline import build_lead_all
from run_lead_all import REGISTRY
from workspace import select_workspace, template_dir, config_dir, prompt_company_info

# 양식·config = 회사 무관 공유 자산. 입력·변환·출력은 회사별 워크스페이스에서 받는다.
TEMPLATE_DIR = template_dir()
CONFIG_DIR = config_dir()
INPUT_DIR = None  # main()에서 워크스페이스로 설정


def _line(msg=""):
    print(msg, flush=True)


def _ask(label: str, required: bool = True, example: str = "") -> str:
    hint = f" (예: {example})" if example else ""
    while True:
        val = input(f"  {label}{hint}: ").strip()
        if val or not required:
            return val
        _line("  ▷ 값을 입력해 주세요.")


def _find_settlement(input_dir: Path) -> "str | None":
    """입력자료에서 정산표 엑셀을 찾는다(파일명에 '정산표' 포함, 임시파일 제외)."""
    cands = [h for h in glob.glob(str(input_dir / "**" / "*정산표*.xls*"), recursive=True)
             if "~$" not in h]
    return sorted(cands)[0] if cands else None


def main():
    _line("=" * 56)
    _line("  정산표 기반 조서 총괄표 일괄 생성")
    _line("=" * 56)

    ws = select_workspace(line=_line)
    if ws is None:
        return 2
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    _line(f"\n[회사] {ws.company}   (작업폴더: {ws.input_dir.parent})")

    settlement = _find_settlement(ws.input_dir)
    if not settlement:
        _line("\n[안내] 정산표 파일을 찾지 못했습니다(파일명에 '정산표' 포함).")
        _line(f"       아래 폴더에 정산표를 넣고 다시 실행해 주세요:\n       {ws.input_dir}")
        return 2
    _line(f"\n[정산표] {Path(settlement).name}")

    # 양식자료에 16조서 템플릿이 있는지 점검(없는 건 건너뛰고 진행)
    present = [it for it in REGISTRY if (TEMPLATE_DIR / it["template"]).exists()]
    missing = [it["code"] for it in REGISTRY if not (TEMPLATE_DIR / it["template"]).exists()]
    _line(f"[양식] {TEMPLATE_DIR}  (확인 {len(present)}/{len(REGISTRY)}조서)")
    if missing:
        _line(f"       미발견: {', '.join(missing)} — 해당 조서는 생성에서 제외됩니다.")
    if not present:
        _line("       템플릿이 하나도 없습니다. 양식자료에 완성본을 넣어 주세요.")
        return 2

    _line("\n[회사 정보 입력]")
    company  = ws.company
    date, preparer, reviewer = prompt_company_info(ws, input, _line)

    out_dir = ws.output_dir
    _line("\n생성 중... (조서별 추출→매핑→이식, 수십 초 소요될 수 있습니다)\n")
    n_total = len(present)
    counter = {"n": 0}

    def _progress(code, ok):
        counter["n"] += 1
        mark = "V" if ok else "X"
        _line("  [%s] %-5s  (%d/%d)" % (mark, code, counter["n"], n_total))

    try:
        reports, comp = build_lead_all(
            settlement=settlement, registry=present, config_dir=str(CONFIG_DIR),
            template_root=str(TEMPLATE_DIR), output_dir=str(out_dir),
            parsed_dir=str(ws.parsed_dir),   # 작업/{회사}/변환자료 — 회사별 격리
            params={"회사명": company, "날짜": date, "preparer": preparer, "reviewer": reviewer},
            progress=_progress,
        )
    except PermissionError:
        _line(f"\n[오류] 출력 파일이 열려 있어 저장할 수 없습니다. 닫고 다시 실행해 주세요:\n       {out_dir}")
        return 1
    except Exception as e:
        _line(f"\n[오류] 예기치 못한 문제로 생성이 중단됐습니다: {type(e).__name__}: {e}")
        return 1

    # 진단 요약
    nerr = sum(1 for r in reports.values() if r.has_error)
    nwarn = sum(1 for r in reports.values() if r.has_warn and not r.has_error)
    lines = ["=" * 60, f"  생성 완료 — {len(reports)}조서 (오류 {nerr} · 경고 {nwarn})", "=" * 60]
    for code, r in reports.items():
        flag = "✗" if r.has_error else ("△" if r.has_warn else "✓")
        lines.append(f"  {flag} {code}")
        for st, lv, m in r.entries:
            if lv in ("warn", "error"):
                lines.append(f"      - [{st}] {m}")
    # 미매핑 계정 리포트(금액·성격·귀속제안) — 정산표 계정 중 어느 조서에도 안 담긴 것.
    # 제조원가명세서 항목은 제외(Q가 총계로 흡수). 귀속제안은 규칙기반(오프라인).
    udetail = comp.get("unmapped_detail") or []
    if udetail:
        lines.append("")
        lines.append(f"  ⚠ 미매핑 계정 {len(udetail)}건 — 어느 조서에도 안 담김(수동 확인·편입 필요):")
        lines.append("      %-22s %18s  %-10s %s" % ("계정", "금액(수정후)", "성격", "귀속제안"))
        lines.append("      " + "-" * 70)
        for d in udetail:
            lines.append("      %-22s %18s  %-10s %s" % (
                str(d["계정"])[:22], format(int(d["금액"] or 0), ","), d["성격"], d["제안"]))
    else:
        lines.append("")
        lines.append("  ✓ 미매핑 계정 없음(정산표 전 계정이 조서에 담김).")
    unmapped = comp.get("unmapped_adjustments") or []
    if unmapped:
        lines.append("")
        lines.append(f"  ⚠ 미매핑 수정분개 {len(unmapped)}건(어느 조서에도 안 붙음 — 수동 확인):")
        for no, accts in unmapped:
            lines.append(f"      #{no}: {', '.join(accts)}")
    summary = "\n".join(lines)
    _line("\n" + summary)

    report_path = out_dir / "조서생성_실행리포트.txt"
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        report_path.write_text(summary, encoding="utf-8")
    except Exception:
        report_path = None

    _line("\n" + "=" * 60)
    _line(f"  결과 폴더 : {out_dir}")
    if report_path:
        _line(f"  실행 리포트: {report_path}")
    _line("=" * 60)
    _line("  ※ 초안입니다. 보조시트(Test·주석·증감분석)와 결론·수정사항은 감사인이 검토·보완하세요.")
    return 0


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)

# -*- coding: utf-8 -*-
"""조서생성기 — 통합 실행기(하나만 골라서 / 여러 개 / 전체).

회사·회사정보(기준일·작성자·검토자)를 **한 번만** 정하고, 메뉴에서 만들 조서를
골라 생성한다. 7개 개별 EXE를 하나로 합친 오케스트레이터.

  · 회사 선택 1회 + 회사정보 입력 1회(이미 저장돼 있으면 Enter)
  · 조서를 하나만 골라도 되고(예: '4'), 여러 개(예: '1,3,4') 또는 전체('A')
  · 선택한 조서를 의존성 순서(조서생성 → 개별 → 기타상세)로 차례로 생성

각 조서 모듈의 main()을 그대로 재사용한다(workspace.push_active로 재선택·재입력 생략).
"""

import importlib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")
except Exception:
    pass

import workspace as W

# 메뉴 순서 = 실행 순서. 7개 모두 독립이라 단독/임의조합 가능.
#   기타상세는 D·CC 총괄표가 없으면 정산표로 자체 선생성(독립). 조서생성을 함께 고르면
#   그 산출물을 재사용하므로 중복 없이 맨 뒤에 둔다.
MENU = [
    ("1", "조서생성",  "정산표 → 16개 조서 총괄표 일괄",              "easy_run_lead"),
    ("2", "A-0",      "현금및현금성자산 (총괄표+보조시트 상세)",        "easy_run_a0"),
    ("3", "A-1",      "은행·금융기관 조회 (완전성·대사)",              "easy_run"),
    ("4", "BBDD",     "장단기차입금 상세",                            "easy_run_bbdd"),
    ("5", "G",        "유무형자산 (롤포워드)",                        "easy_run_g"),
    ("6", "R-S",      "판매관리비·기타손익 보조시트",                  "easy_run_rs"),
    ("7", "기타상세",  "기타자산D200·기타부채CC200 상세 (정산표·잔액)",  "easy_run_detail"),
]


def _line(m=""):
    print(m, flush=True)


def _parse_selection(raw: str):
    """'A'/'전체' → 전체, '1,3,4' / '1 3 4' → 해당 번호. 메뉴 순서대로 정렬·중복제거."""
    raw = raw.strip()
    if raw.lower() in ("a", "all", "전체"):
        return [row for row in MENU]
    toks = [t for t in raw.replace(",", " ").split() if t]
    nums = {t for t in toks if t.isdigit()}
    return [row for row in MENU if row[0] in nums]  # 메뉴(의존성) 순서 유지


def main():
    _line("=" * 60)
    _line("  조서생성기 — 통합 실행기 (하나만 / 여러 개 / 전체)")
    _line("=" * 60)

    # 1) 회사 선택 + 회사정보 1회
    ws = W.select_workspace(line=_line)
    if ws is None:
        return 2
    _line(f"\n[회사] {ws.company}   (작업폴더: {ws.input_dir.parent})")
    _line("\n[회사 정보]  (저장돼 있으면 Enter로 유지)")
    date, preparer, reviewer = W.prompt_company_info(ws, input, _line)

    # 2) 만들 조서 선택
    _line("\n[생성할 조서 선택]")
    for num, code, desc, _ in MENU:
        _line(f"  {num}) {code:8} {desc}")
    _line("  A) 전체 (의존성 순서대로)")
    while True:
        sel = input("\n  번호(쉼표로 여러 개) 또는 A: ").strip()
        chosen = _parse_selection(sel)
        if chosen:
            break
        _line("  ▷ 1~7 번호(쉼표 가능) 또는 A 를 입력해 주세요.")

    _line("\n  선택: " + ", ".join(c[1] for c in chosen))

    # 3) 선택한 조서를 차례로 생성 (회사·정보는 고정 — 각 모듈이 조용히 재사용)
    W.push_active(ws, {"date": date, "preparer": preparer, "reviewer": reviewer})
    results = []
    try:
        for num, code, desc, modname in chosen:
            _line("\n" + "─" * 60)
            _line(f"▶ {code} 생성")
            _line("─" * 60)
            try:
                mod = importlib.import_module(modname)
                rc = mod.main()
            except Exception as e:
                _line(f"[오류] {code}: {type(e).__name__}: {e}")
                rc = 1
            results.append((code, rc))
    finally:
        W.clear_active()

    # 4) 종합 요약
    _line("\n" + "=" * 60)
    _line("  생성 요약")
    _line("=" * 60)
    for code, rc in results:
        mark = "✓ 완료" if rc == 0 else ("· 건너뜀(입력 부족)" if rc == 2 else "✗ 오류")
        _line(f"  {mark:18} {code}")
    _line(f"\n  결과 위치: {ws.output_dir}")
    _line("  ※ 모든 산출물은 초안입니다. 감사인 검토·보완 후 완성하세요.")
    return 0 if all(rc == 0 for _, rc in results) else 1


if __name__ == "__main__":
    rc = main()
    try:
        input("\n[Enter] 키를 누르면 창이 닫힙니다...")
    except EOFError:
        pass
    sys.exit(rc)

"""A-1 조서 자동 생성 실행기 (CLI).

데이터 폴더에서 입력 3종(발송 Control Sheet · 거래처원장/결산보고서 · 조회서 취합엑셀)을
자동 탐색하고, 회사명·날짜·작성자만 입력하면 A-1(A100/A200/A300)을 한 파일로 생성한다.

사용 예:
  python src/run_a1.py --data "회사데이터폴더" \
      --company "회사명" --date 2025-12-31 --preparer CJH --reviewer KHK \
      --template "양식/_A-1_template_....xlsx" --output "결과.xlsx"

파일 자동탐색이 빗나가면 --control/--confirm/--ledger 로 직접 지정할 수 있다.
출력물은 초안이다 (은행연합회·우편회신·FN은 감사인 영역으로 비워둠/플래그).
"""

import argparse
import glob
import sys
from pathlib import Path

# src 를 import 경로에 추가 (직접 실행 대응)
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pipeline import build_a1


def _find_one(data_dir: str, patterns: list[str], exclude: tuple = ()) -> "str | None":
    """data_dir 하위에서 패턴에 맞는 첫 파일을 반환. 임시(~$) 제외."""
    for pat in patterns:
        hits = sorted(
            h for h in glob.glob(str(Path(data_dir) / "**" / pat), recursive=True)
            if "~$" not in h and not any(e in Path(h).name for e in exclude)
        )
        if hits:
            if len(hits) > 1:
                print(f"  ⚠ 여러 파일 매칭({pat}) → 첫 번째 사용: {Path(hits[0]).name}")
            return hits[0]
    return None


def discover_inputs(data_dir: str) -> dict:
    """입력자료 폴더에서 역할별 파일을 자동 탐색한다.

    표준 파일명(설명서 안내)을 **최우선**으로 인식하고, 못 찾으면 기존(레거시) 이름 패턴으로
    폴백한다. 표준명은 토큰으로 '시작'만 하면 되고 뒤에 회사명·날짜를 붙여도 된다.
      필수자료1_ = 발송 Control Sheet / 필수자료2_ = 조회서 취합 /
      필수자료3_ = 잔액 소스(_유형1 거래처잔액현황 · _유형2 결산보고서) /
      보조자료_A1 = 개별 잔액명세서(현금성 등, A-1 보강용·선택)
    """
    # 표준 파일명(설명서 안내)으로 역할을 인식한다. 토큰으로 시작만 하면 뒤는 자유.
    #   필수자료1_=발송CS / 필수자료2_=조회서 / 필수자료3_=잔액(_유형1·_유형2) / 보조자료_A1=개별명세서
    control = _find_one(data_dir, ["필수자료1*"], exclude=("필수자료2", "필수자료3", "보조자료"))
    confirm = _find_one(data_dir, ["필수자료2*"], exclude=("필수자료1", "필수자료3", "보조자료"))
    ledger  = _find_one(data_dir, ["필수자료3*"], exclude=("보조자료",))
    aux_a1  = _find_one(data_dir, ["보조자료_A1*"])
    return {"control_sheet": control, "confirm_xlsx": confirm,
            "ledger_src": ledger, "aux_a1": aux_a1}


def main(argv=None):
    ap = argparse.ArgumentParser(description="A-1 조서 자동 생성")
    ap.add_argument("--data", help="회사 데이터 폴더 (입력 3종 자동탐색)")
    ap.add_argument("--company", required=True, help="회사명")
    ap.add_argument("--date", required=True, help="기준일 (예: 2025-12-31)")
    ap.add_argument("--preparer", required=True, help="작성자")
    ap.add_argument("--reviewer", default="", help="검토자")
    ap.add_argument("--template", required=True, help="작업용 템플릿(.xlsx)")
    ap.add_argument("--config-dir", default=str(Path(__file__).resolve().parent.parent / "_internal" / "config"))
    ap.add_argument("--output", required=True, help="결과 파일 경로")
    # 자동탐색 override
    ap.add_argument("--control", help="발송 Control Sheet 직접 지정")
    ap.add_argument("--confirm", help="조회서 취합엑셀 직접 지정")
    ap.add_argument("--ledger", help="거래처원장/결산보고서 직접 지정")
    args = ap.parse_args(argv)

    found = discover_inputs(args.data) if args.data else {}
    control = args.control or found.get("control_sheet")
    confirm = args.confirm or found.get("confirm_xlsx")
    ledger = args.ledger or found.get("ledger_src")

    missing = [n for n, v in [("Control Sheet", control), ("조회서 취합엑셀", confirm),
                              ("거래처원장/결산보고서", ledger)] if not v]
    if missing:
        print(f"[오류] 입력 파일을 찾지 못함: {', '.join(missing)}")
        print("  --data 폴더를 확인하거나 --control/--confirm/--ledger 로 직접 지정하세요.")
        return 2

    print("입력 파일:")
    print(f"  Control Sheet : {Path(control).name}")
    print(f"  조회서 취합   : {Path(confirm).name}")
    print(f"  잔액 소스     : {Path(ledger).name}"
          f" ({'결산보고서' if '결산보고서' in Path(ledger).name else '거래처원장'})")

    report = build_a1(
        control_sheet=control, ledger_src=ledger, confirm_xlsx=confirm,
        template=args.template, config_dir=args.config_dir, output=args.output,
        params={"회사명": args.company, "날짜": args.date,
                "preparer": args.preparer, "reviewer": args.reviewer},
    )
    print("\n" + report.render())
    print(f"\n결과 파일 → {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

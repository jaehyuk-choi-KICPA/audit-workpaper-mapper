"""A-0 총괄표 자동 생성 실행기 (CLI) — 1차 범위: 본문 + 수정사항.

정산표(별도정산표·수정사항집계 시트)를 단일 소스로 A-0 총괄표의 본문과 '4.수정사항'을 채운다.
나머지 섹션(Recap·증감분석·Test·서술·결론)은 감사인 영역으로 비워둔다(초안).

사용 예:
  python src/run_a0.py --settlement "정산표.xlsx" \
      --template "양식/A-0_4000_A000_template.xlsx" --output "결과.xlsx"
  또는 --data 폴더에서 *정산표* 파일을 자동 탐색.
"""

import argparse
import glob
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pipeline import build_a0


def _find_settlement(data_dir: str) -> "str | None":
    for pat in ("*wp_정산표*.xls*", "*정산표*.xls*"):
        hits = sorted(h for h in glob.glob(str(Path(data_dir) / "**" / pat), recursive=True)
                      if "~$" not in h)
        if hits:
            return hits[0]
    return None


def main(argv=None):
    ap = argparse.ArgumentParser(description="A-0 총괄표 자동 생성 (본문 + 수정사항)")
    ap.add_argument("--data", help="회사 데이터 폴더 (정산표 자동탐색)")
    ap.add_argument("--settlement", help="정산표 파일 직접 지정")
    ap.add_argument("--template", required=True, help="A-0 작업용 템플릿(.xlsx)")
    ap.add_argument("--config-dir",
                    default=str(Path(__file__).resolve().parent.parent / "_internal" / "config"))
    ap.add_argument("--output", required=True, help="결과 파일 경로")
    ap.add_argument("--company", default="", help="회사명(선택)")
    args = ap.parse_args(argv)

    settlement = args.settlement or (_find_settlement(args.data) if args.data else None)
    if not settlement:
        print("[오류] 정산표 파일을 찾지 못함. --settlement 또는 --data 폴더를 지정하세요.")
        return 2

    print(f"정산표: {Path(settlement).name}")
    report = build_a0(
        settlement=settlement, template=args.template,
        config_dir=args.config_dir, output=args.output,
        params={"회사명": args.company} if args.company else {},
    )
    print("\n" + report.render())
    print(f"\n결과 파일 → {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

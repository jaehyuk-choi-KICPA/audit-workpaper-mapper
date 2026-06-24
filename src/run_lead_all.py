"""잔액형 총괄표 일괄 생성 실행기.

정산표 1개로 등록된 모든 잔액형 조서의 총괄표(본문+수정사항)를 한 번에 생성한다.
정산표 파싱은 변환자료에 캐시되어 조서 간 공유(재파싱 0). 라우팅 완전성도 함께 점검.

사용:
  python src/run_lead_all.py --settlement "정산표.xlsx" \
      --template-root "_internal/양식" --output "_internal/output" --company "회사명"
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pipeline import build_lead_all

# 조서 레지스트리: code / config(YAML) / template(양식 루트 기준) / kind(생략=engine).
# kind: engine=a0.py 섹션 렌더 / refill·capital·cogs·sales=고정격자 특수 생성기.
REGISTRY = [
    {"code": "A-0", "config": "a0.yaml", "template": "A-0_4000_A000_template.xlsx"},
    {"code": "C",   "config": "c.yaml",  "template": "C_4000_template.xlsx"},
    {"code": "AA",  "config": "aa.yaml", "template": "AA_4000_template.xlsx"},
    {"code": "CC",  "config": "cc.yaml", "template": "CC_4000_template.xlsx"},
    {"code": "D",   "config": "d.yaml",  "template": "D_4000_template.xlsx"},
    {"code": "EE",  "config": "ee.yaml", "template": "EE_4000_template.xlsx"},
    {"code": "R",   "config": "r.yaml",  "template": "R_4000_template.xlsx"},
    {"code": "S",   "config": "s.yaml",  "template": "S_4000_template.xlsx"},
    # 특수형(고정격자 — 엔진 섹션 렌더로 안 맞는 조서)
    {"code": "R-2",  "config": "r2.yaml",   "template": "R2_4000_template.xlsx",   "kind": "refill"},
    {"code": "Q",    "config": "q.yaml",    "template": "Q_4000_template.xlsx",    "kind": "cogs"},
    {"code": "P",    "config": "p.yaml",    "template": "P_4000_template.xlsx",    "kind": "sales"},
    {"code": "GG",   "config": "gg.yaml",   "template": "GG_4000_template.xlsx",   "kind": "capital"},
    {"code": "BBDD", "config": "bbdd.yaml", "template": "BBDD_4000_template.xlsx", "kind": "refill"},
]


def main(argv=None):
    ap = argparse.ArgumentParser(description="잔액형 총괄표 일괄 생성")
    ap.add_argument("--settlement", required=True)
    ap.add_argument("--template-root", default=str(Path(__file__).resolve().parent.parent / "_internal" / "양식"))
    ap.add_argument("--config-dir", default=str(Path(__file__).resolve().parent.parent / "_internal" / "config"))
    ap.add_argument("--output", required=True)
    ap.add_argument("--company", default="")
    args = ap.parse_args(argv)

    reports, comp = build_lead_all(
        settlement=args.settlement, registry=REGISTRY,
        config_dir=args.config_dir, template_root=args.template_root,
        output_dir=args.output,
        params={"회사명": args.company} if args.company else {},
    )
    for code, rep in reports.items():
        gen = [m for st, lv, m in rep.entries if st == "생성"]
        out = [m for st, lv, m in rep.entries if st == "출력"]
        flag = "✗" if rep.has_error else ("△" if out else "✓")
        print(f"[{flag}] {code}: {'; '.join(gen)}" + (f"  | 출력경고 {len(out)}" if out else ""))
    print("\n=== 라우팅 완전성 ===")
    print(f"소유 대분류 {comp['owned']}종")
    if comp["duplicate"]:
        print("⚠ 중복소유:", comp["duplicate"])
    print(f"미매핑(아직 조서 없음) {len(comp['unmapped'])}종:")
    print("  " + ", ".join(comp["unmapped"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

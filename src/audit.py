"""완전성 감사: 정산표 전 대분류가 빠짐없이 조서에 매핑되는지 점검(회사 무관).

account_routing.yaml(exact + keyword_rules)로 각 대분류를 조서 코드에 배정하고,
조서별로 묶어 보여준다. exact/keyword로 못 잡아 기본(판관비 R)으로 떨어진 건 '?'로 표시
(BS 계정이 R로 새면 라우팅 누락 신호). 빌드된 조서 여부도 함께 표시.

사용: python src/audit.py [정산표경로]
"""
import sys, re
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import yaml
from extractors import parse_trial_balance

BUILT = {"A", "B", "C", "AA", "CC", "D", "E", "EE", "R", "R-2", "GG", "Q", "P", "S", "BBDD"}  # 생성기 있는 조서
PENDING = {"G", "H", "TAX"}  # G·H 변동분해형(보류), TAX 법인세는 S에 포함


def _norm(s):
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def route(name, routing):
    n = _norm(name)
    ex = {_norm(k): v for k, v in routing.get("exact", {}).items()}
    if n in ex:
        return ex[n], False
    for rule in routing.get("keyword_rules", []):
        kw = rule.get("contains")
        if kw and _norm(kw) in n:
            return rule["code"], False
        if rule.get("pl_expense_default"):
            return rule["code"], True   # 기본 라우팅(불확실)
    return None, True


def main(settlement=None):
    base = Path(__file__).resolve().parent.parent
    S = settlement or str(next(p for p in base.glob("_internal/data/**/*정산표*.xls*")
                               if "~$" not in str(p)))
    routing = yaml.safe_load((base / "_internal/config/account_routing.yaml").read_text(encoding="utf-8"))
    tb = parse_trial_balance(S)

    by_code = {}
    for r in tb:
        if (r["기초"] in (None, 0)) and (r["기말"] in (None, 0)) and (r["수정후"] in (None, 0)):
            continue
        code, default = route(r["대분류"], routing)
        by_code.setdefault(code, []).append((r["대분류"], default))

    print(f"=== 완전성 감사: {Path(S).name} ===")
    n_unmapped = n_default = 0
    for code in sorted(by_code, key=lambda c: (c is None, str(c))):
        accts = by_code[code]
        uniq = []
        seen = set()
        for a, d in accts:
            if a in seen:
                continue
            seen.add(a); uniq.append((a, d))
        status = "✗미구현" if code in PENDING else ("?미매핑" if code is None else "✓")
        flag = "" if code in BUILT else f" [{status}]"
        names = ", ".join(a + ("?" if d else "") for a, d in uniq)
        print(f"  [{code or '미매핑'}]{flag} {len(uniq)}종: {names}")
        if code is None:
            n_unmapped += len(uniq)
        n_default += sum(1 for _, d in uniq if d)
    print(f"\n빌드됨: {sorted(BUILT)}")
    print(f"라우팅O·생성기X(후속): {sorted(PENDING)}")
    print(f"미매핑(어디에도 못 감): {n_unmapped}종 | 기본(R)으로 떨어진 '?'표시: {n_default}종 — BS가 섞였으면 라우팅 보완")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else None)

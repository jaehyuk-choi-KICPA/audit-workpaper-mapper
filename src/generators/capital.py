"""GG 자본 총괄표 전용 생성기.

자본 총괄표는 일반 잔액형(대분류별/섹션별 소계)과 구조가 다르다:
  - 로마자 구분(Ⅰ.자본금~Ⅴ.이익잉여금)이 **상위 합계 행**이고 그 아래 leaf 계정이 온다.
    (완성본은 로마/자본총계가 이미 하위 참조 SUM 수식 → leaf만 갱신하면 자동 재계산)
  - 수정사항: 이익잉여금(미처분이익잉여금)에 **수정분개의 이익잉여금 Effect 합계**가 집계된다
    (개별 계정 K-L이 아니라 손익→이익잉여금 전체 효과).

따라서 엔진 렌더 대신 **완성본 격자 보존 in-place 리필**을 한다: leaf 행의 계정명을 정산표
대분류에 매칭(최장 공통 접두사)해 기초/기말을 채우고, 미처분이익잉여금 행 수정사항에 집계액을 쓴다.
"""

import re
from collections import defaultdict

import openpyxl

_ROMAN = ("Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ")
_NUMF = '#,##0;[Red]\\(#,##0\\);"-"'


def _norm(s) -> str:
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def _common_prefix(a: str, b: str) -> int:
    n = 0
    for x, y in zip(a, b):
        if x == y:
            n += 1
        else:
            break
    return n


def fill_capital(template: str, tb_rows: list, adj_entries: list, output: str, cfg: dict) -> dict:
    """자본 총괄표를 정산표로 in-place 리필한다.

    cfg: {sheet, header_row, data_end_anchor, name_col, base_col, end_col, adj_col}
    """
    sheet = cfg["sheet"]
    nci = cfg.get("name_col", 2)        # 계정명 열(1-indexed)
    bcol = cfg.get("base_col", "C")     # 기초
    ecol = cfg.get("end_col", "D")      # 기말(회사제시)
    acol = cfg.get("adj_col", "G")      # 수정사항
    hr = cfg.get("header_row", 9)
    end_anchor = _norm(cfg.get("data_end_anchor", "자본총계"))

    # 정산표 대분류별 기초/기말 합산
    by_class = defaultdict(lambda: {"기초": 0, "기말": 0})
    for rec in tb_rows:
        g = by_class[_norm(rec["대분류"])]
        g["기초"] += rec["기초"] or 0
        g["기말"] += rec["기말"] or 0

    # 명시적 매핑(자본 leaf 라벨 → 정산표 대분류) — 자본 계정은 명칭이 비슷해 fuzzy 오매칭 위험
    # (예: '매도가능증권평가익'(자본)이 '매도가능증권'(자산)에 걸림). 명시 매핑으로 정확히 잡는다.
    name_map = {_norm(k): _norm(v) for k, v in cfg.get("name_map", {}).items()}

    def find(name):
        n = _norm(name)
        if name_map:
            tgt = name_map.get(n)
            return by_class.get(tgt) if tgt else None
        if n in by_class:
            return by_class[n]
        best, bl = None, 3
        for k in by_class:
            c = _common_prefix(k, n)
            if c > bl:
                bl, best = c, k
        return by_class[best] if best else None

    # 이익잉여금 수정사항 = 수정분개의 이익잉여금 Effect 합계
    re_adj = sum(l["이익잉여금"] for e in adj_entries for l in e["lines"] if l.get("이익잉여금"))

    wb = openpyxl.load_workbook(template)
    ws = wb[sheet]
    n_filled = 0
    for r in range(hr + 1, ws.max_row + 1):
        b = ws.cell(r, nci).value
        nb = _norm(b)
        if not nb:
            continue
        if nb.startswith(end_anchor):
            break
        if nb.startswith(_ROMAN) or "자본총계" in nb:
            continue                       # 로마/총계는 수식 자동 재계산
        rec = find(b)
        if rec:
            ws[f"{bcol}{r}"] = rec["기초"]; ws[f"{bcol}{r}"].number_format = _NUMF
            ws[f"{ecol}{r}"] = rec["기말"]; ws[f"{ecol}{r}"].number_format = _NUMF
            if "미처분이익잉여금" in nb and re_adj:
                ws[f"{acol}{r}"] = re_adj; ws[f"{acol}{r}"].number_format = _NUMF
            n_filled += 1
        else:
            # ★ 블랭크화: 정산표에 없는 leaf(타 회사 잔재)는 값을 비운다(0). 라벨/구조는 보존.
            ws[f"{bcol}{r}"] = 0; ws[f"{ecol}{r}"] = 0
            ws[f"{acol}{r}"] = None

    for col in (bcol, ecol, "E", acol, "H"):
        cur = ws.column_dimensions[col].width or 0
        ws.column_dimensions[col].width = max(cur, 16)

    Path = __import__("pathlib").Path
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return {"n_filled": n_filled, "re_adj": re_adj}

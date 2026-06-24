"""P(매출)·Q(매출원가) 총괄표 생성기.

이 두 조서는 회사마다 계정 구성이 크게 달라(매출 종류·상품매출원가 유무) 고정행 리필로는
부족하다. 따라서:
  - Q(매출원가): 고정 격자(미완성공사·당기총공사비용·공사매출원가)를 보존하며 당기총공사비용을
    제조원가명세서 총계로 채우고, **상품매출원가가 있으면 공사매출원가 아래에 행을 삽입**해 매칭.
  - P(매출): IS 매출 표를 **있는 매출 대분류만 동적 렌더**(회사마다 매출 종류가 달라서),
    BS 교차참조 행(공사선수금·매출채권)은 이름으로 찾아 채운다.

제조원가명세서 행(제조원가=True)과 손익 매출원가 행(제조원가=False)은 별도정산표 파서가
구분해 둔다. 회사 무관 동작을 위해 계정명은 키워드 규칙으로 매칭한다.
"""

import re
from collections import defaultdict
from copy import copy
from pathlib import Path

import openpyxl

_NUMF = '#,##0;[Red]\\(#,##0\\);"-"'
_PCTF = '0%;[Red]\\(0%\\)'


def _norm(s) -> str:
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def _agg_rule(tb_rows, rule):
    """tb_rows에서 규칙(제조 flag + 포함/제외 키워드)에 맞는 행을 합산. (기초/기말/수정사항)"""
    want_mfg = rule.get("제조")
    inc = [_norm(x) for x in rule.get("포함", [])]
    exc = [_norm(x) for x in rule.get("제외", [])]
    agg = {"기초": 0, "기말": 0, "수정사항": 0}
    hit = False
    for rec in tb_rows:
        if want_mfg is not None and bool(rec.get("제조원가")) != bool(want_mfg):
            continue
        text = _norm(rec["대분류"]) + _norm(rec["계정명"])
        if inc and not any(k in text for k in inc):
            continue
        if exc and any(k in text for k in exc):
            continue
        hit = True
        for k in ("기초", "기말", "수정사항"):
            agg[k] += rec[k] or 0
    return agg, hit


def _copy_row_style(ws, src_r, dst_r, ncols):
    """src_r 행의 셀 스타일을 dst_r 행으로 복사(테두리/서식 보존)."""
    for c in range(1, ncols + 1):
        sc = ws.cell(src_r, c)
        dc = ws.cell(dst_r, c)
        if sc.has_style:
            dc.font = copy(sc.font); dc.border = copy(sc.border)
            dc.fill = copy(sc.fill); dc.alignment = copy(sc.alignment)
            dc.number_format = sc.number_format


def _widen(ws, cols, minw=16):
    for col in cols:
        cur = ws.column_dimensions[col].width or 0
        ws.column_dimensions[col].width = max(cur, minw)


# ─────────────────────────────── Q: 매출원가 ───────────────────────────────

def fill_cogs(template, tb_rows, output, cfg):
    """Q 매출원가 총괄표. 공사원가 롤포워드(기초미완성·당기총공사비용·기말미완성)를 제조원가
    명세서 행으로 채워 공사매출원가가 손익 도급공사매출원가와 일치하게 하고, 상품매출원가가
    있으면 공사매출원가 아래 '상품매출원가'+'매출원가계' 행을 삽입한다.

    cfg: {sheet, name_col, base_col, end_col, adj_col, formulas{col:tmpl},
          name_map{라벨: 규칙}, cogs_anchor, commodity_label, commodity_rule, total_label}
    참고: 타계정 전입/전출액은 (에서/으로)가 상쇄되어 공사매출원가에 영향 없으므로 채우지 않는다.
    """
    sheet = cfg["sheet"]
    nci = cfg.get("name_col", 2)
    bcol, ecol, acol = cfg["base_col"], cfg["end_col"], cfg["adj_col"]
    formulas = cfg.get("formulas", {})

    wb = openpyxl.load_workbook(template)
    ws = wb[sheet]

    def put(r, rec):
        ws[f"{bcol}{r}"] = rec["기초"]; ws[f"{bcol}{r}"].number_format = _NUMF
        ws[f"{ecol}{r}"] = rec["기말"]; ws[f"{ecol}{r}"].number_format = _NUMF
        if rec.get("수정사항"):
            ws[f"{acol}{r}"] = rec["수정사항"]; ws[f"{acol}{r}"].number_format = _NUMF
        for col, tmpl in formulas.items():
            ws[f"{col}{r}"] = tmpl.format(r=r)
            ws[f"{col}{r}"].number_format = _PCTF if col == cfg.get("pct_col") else _NUMF

    # 1) 롤포워드 행(기초미완성·당기총공사비용·기말미완성)을 제조원가명세서 행으로 채움
    name_map = {_norm(k): v for k, v in cfg.get("name_map", {}).items()}
    cogs_anchor = _norm(cfg["cogs_anchor"])
    cogs_row = None
    for r in range(1, ws.max_row + 1):
        nm = _norm(ws.cell(r, nci).value)
        if nm in name_map:
            agg, _ = _agg_rule(tb_rows, name_map[nm])
            put(r, agg)
        if nm.startswith(cogs_anchor) and cogs_row is None:
            cogs_row = r

    # 2) 상품매출원가(있을 때만) → 공사매출원가 아래에 행 삽입 + 매출원가계
    comm_agg, has_comm = _agg_rule(tb_rows, cfg["commodity_rule"])
    n_extra = 0
    if has_comm and any(comm_agg[k] for k in ("기초", "기말", "수정사항")) and cogs_row:
        ws.insert_rows(cogs_row + 1, 2)
        comm_r, total_r = cogs_row + 1, cogs_row + 2
        ncols = ws.max_column
        _copy_row_style(ws, cogs_row, comm_r, ncols)
        _copy_row_style(ws, cogs_row, total_r, ncols)
        # 상품매출원가 행
        ws.cell(comm_r, nci).value = cfg.get("commodity_label", "상품매출원가")
        put(comm_r, comm_agg)
        # 매출원가계 = 공사매출원가 + 상품매출원가
        ws.cell(total_r, nci).value = cfg.get("total_label", "매출원가계")
        for col in (bcol, ecol, acol):
            ws[f"{col}{total_r}"] = f"={col}{cogs_row}+{col}{comm_r}"
            ws[f"{col}{total_r}"].number_format = _NUMF
        for col, tmpl in formulas.items():
            ws[f"{col}{total_r}"] = tmpl.format(r=total_r)
            ws[f"{col}{total_r}"].number_format = _PCTF if col == cfg.get("pct_col") else _NUMF
        n_extra = 2

    _widen(ws, [bcol, ecol, acol] + list(formulas.keys()))
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return {"commodity_cogs": comm_agg["기말"] if has_comm else 0, "rows_inserted": n_extra}


# ─────────────────────────────── P: 매출 ───────────────────────────────

def fill_sales(template, tb_rows, output, cfg):
    """P 매출 총괄표. IS 매출 표를 '있는 매출 대분류'만 동적 렌더하고, BS 교차참조 행
    (공사선수금·매출채권)은 이름으로 찾아 채운다.

    cfg: {sheet, is_header_row, is_start_row, is_total_label, name_col,
          base_col, end_col, adj_col, formulas, is_groups[대분류...],
          bs_rows[{label, name}], pct_col?}
    """
    sheet = cfg["sheet"]
    nci = cfg.get("name_col", 2)
    bcol, ecol, acol = cfg["base_col"], cfg["end_col"], cfg["adj_col"]
    formulas = cfg.get("formulas", {})
    is_start = cfg["is_start_row"]
    total_label = _norm(cfg["is_total_label"])

    by_class = defaultdict(lambda: {"기초": 0, "기말": 0, "수정사항": 0, "계정명": None})
    by_name = {}
    for rec in tb_rows:
        if rec.get("제조원가"):           # 제조원가명세서 행은 매출 아님
            continue
        g = by_class[_norm(rec["대분류"])]
        for k in ("기초", "기말", "수정사항"):
            g[k] += rec[k] or 0
        if g["계정명"] is None:
            g["계정명"] = rec["대분류"]
        by_name[_norm(rec["계정명"])] = {"기초": rec["기초"] or 0, "기말": rec["기말"] or 0,
                                        "수정사항": rec["수정사항"] or 0}

    wb = openpyxl.load_workbook(template)
    ws = wb[sheet]
    ncols = ws.max_column

    def setfmt_formulas(r):
        for col, tmpl in formulas.items():
            ws[f"{col}{r}"] = tmpl.format(r=r)
            ws[f"{col}{r}"].number_format = _PCTF if col == cfg.get("pct_col") else _NUMF

    def put(r, rec, name):
        ws.cell(r, nci).value = name
        ws[f"{bcol}{r}"] = rec["기초"]; ws[f"{bcol}{r}"].number_format = _NUMF
        ws[f"{ecol}{r}"] = rec["기말"]; ws[f"{ecol}{r}"].number_format = _NUMF
        ws[f"{acol}{r}"] = rec.get("수정사항") or 0; ws[f"{acol}{r}"].number_format = _NUMF
        setfmt_formulas(r)

    # 1) 현재 IS 합계 행 위치 탐색
    total_row = None
    for r in range(is_start, ws.max_row + 1):
        if _norm(ws.cell(r, nci).value).startswith(total_label):
            total_row = r
            break
    if total_row is None:
        total_row = is_start + 1

    # 2) 렌더할 매출 대분류(있고 0 아님)
    present = [g for g in cfg["is_groups"]
               if _norm(g) in by_class and any(by_class[_norm(g)][k] for k in ("기초", "기말", "수정사항"))]
    needed = max(len(present), 1)
    avail = total_row - is_start            # 합계 위 데이터 행 수
    delta = needed - avail
    if delta > 0:
        ws.insert_rows(is_start, delta)     # 데이터 행 삽입(합계·BS 아래로 밀림)
        for i in range(delta):
            _copy_row_style(ws, is_start + delta, is_start + i, ncols)  # 도너=원래 첫 데이터행
        total_row += delta
    # 데이터 행 비우고 렌더
    for r in range(is_start, total_row):
        for col in (bcol, ecol, acol):
            ws[f"{col}{r}"] = None
        ws.cell(r, nci).value = None
    r = is_start
    for g in present:
        rec = by_class[_norm(g)]
        put(r, rec, rec["계정명"] or g)
        r += 1
    # 3) 합계 행: SUM 범위 재작성(삽입으로 범위가 어긋날 수 있음)
    ws.cell(total_row, nci).value = cfg.get("is_total_label", "합계")
    for col in (bcol, ecol, acol):
        ws[f"{col}{total_row}"] = f"=SUM({col}{is_start}:{col}{total_row - 1})"
        ws[f"{col}{total_row}"].number_format = _NUMF
    setfmt_formulas(total_row)

    # 4) BS 교차참조 행: 이름으로 찾아 채움(공사선수금·매출채권 등)
    n_bs = 0
    for spec in cfg.get("bs_rows", []):
        lbl = _norm(spec["label"])
        nm = _norm(spec["name"])
        rec = by_name.get(nm) or by_class.get(nm)
        if not rec:
            continue
        for rr in range(total_row + 1, ws.max_row + 1):
            if _norm(ws.cell(rr, nci).value).startswith(lbl):
                ws[f"{bcol}{rr}"] = rec["기초"]; ws[f"{bcol}{rr}"].number_format = _NUMF
                ws[f"{ecol}{rr}"] = rec["기말"]; ws[f"{ecol}{rr}"].number_format = _NUMF
                ws[f"{acol}{rr}"] = rec.get("수정사항") or 0; ws[f"{acol}{rr}"].number_format = _NUMF
                setfmt_formulas(rr)
                n_bs += 1
                break

    _widen(ws, [bcol, ecol, acol] + list(formulas.keys()))
    Path(output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return {"n_sales": len(present), "n_bs": n_bs}

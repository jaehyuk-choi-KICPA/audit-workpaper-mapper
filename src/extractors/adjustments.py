"""수정사항집계 파서 — 감사 수정분개표.

정산표의 **수정사항집계(25) 시트**(당기 수정분개)를 entry(분개 묶음) 단위로 구조화한다.
A-0 등 조서 총괄표의 '수정사항' 섹션에 관련 분개를 **그대로 재현**하기 위함이므로,
부호 변환 없이 원본 차변/대변 구조를 보존한다.

시트 구조(헤더 2행):
  행1: # | 계정과목 | 금액 | Effect | Description
  행2:      차변 대변 | 차변 대변 | 손익 이익잉여금
  - entry 시작 = '#' 열에 정수가 있는 행. 다음 정수 전까지 한 entry.
  - 한 줄 = 차변(계정과목 차변열+금액 차변열) 또는 대변(계정과목 대변열+금액 대변열).
  - 설명만 있는 줄은 entry 비고로 모은다. 줄 없는 빈 entry는 스킵.

출력:
  [{"no": int, "lines": [{"side","계정","금액","손익","이익잉여금","설명"}...],
    "notes": [str...]}, ...]
시트/헤더 없으면 빈 리스트.
"""

import warnings
from pathlib import Path

import openpyxl

from ._headers import normalize

_DEFAULT_SHEET = "수정사항집계(25)"


def _is_amount(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _as_int(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    s = str(v).strip() if v is not None else ""
    return int(s) if s.isdigit() else None


def _find_header(rows: list[tuple]) -> "tuple[int, dict] | None":
    """(데이터 시작행 인덱스, {필드: 열인덱스}) 반환. 0-indexed.

    행1(계정과목/금액/Effect/Description) + 행2(차변/대변/손익/이익잉여금) 구조에 앵커.
    """
    for i in range(len(rows) - 1):
        n1 = [normalize(c) for c in rows[i]]
        n2 = [normalize(c) for c in rows[i + 1]]
        col_no = next((j for j, n in enumerate(n1) if n == "#"), None)
        col_계정 = next((j for j, n in enumerate(n1) if n == "계정과목"), None)
        col_금액 = next((j for j, n in enumerate(n1) if n == "금액"), None)
        col_effect = next((j for j, n in enumerate(n1) if n.lower().startswith("effect")), None)
        col_desc = next((j for j, n in enumerate(n1) if "description" in n.lower()), None)
        if None in (col_계정, col_금액):
            continue
        # 행2가 차변/대변 서브헤더인지 확인
        if not (normalize(rows[i + 1][col_계정] if col_계정 < len(rows[i + 1]) else "") == "차변"):
            continue
        cm = {
            "no": col_no,
            "차변계정": col_계정, "대변계정": col_계정 + 1,
            "금액차변": col_금액, "금액대변": col_금액 + 1,
            "손익": col_effect, "이익잉여금": (col_effect + 1) if col_effect is not None else None,
            "설명": col_desc,
        }
        return i + 2, cm   # 데이터는 서브헤더(행2) 다음부터
    return None


def parse_adjustments(path: str, sheet_name: str = _DEFAULT_SHEET) -> list[dict]:
    """수정사항집계 시트를 entry 리스트로 반환."""
    p = Path(path)
    wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        warnings.warn(f"[수정사항 파서] 시트 '{sheet_name}' 없음. 존재: {wb.sheetnames}")
        return []
    rows = list(wb[sheet_name].iter_rows(values_only=True))
    wb.close()

    found = _find_header(rows)
    if found is None:
        warnings.warn("[수정사항 파서] 헤더(계정과목/금액 + 차변/대변)를 찾지 못함.")
        return []
    start, cm = found

    def g(row, key):
        j = cm.get(key)
        return row[j] if j is not None and j < len(row) else None

    entries: list[dict] = []
    cur = None
    for row in rows[start:]:
        # 표1 끝: '계'(총계) 행에서 중단(이후 '(2) 미수정왜곡표시' 등 다른 표의 헤더 흡수 방지).
        # '계'는 '#'(no) 열에 온다.
        if normalize(g(row, "no")) == "계":
            break
        no = _as_int(g(row, "no"))
        if no is not None:                       # 새 entry 시작
            if cur and (cur["lines"] or cur["notes"]):
                entries.append(cur)
            cur = {"no": no, "lines": [], "notes": []}
        if cur is None:
            continue
        손익 = g(row, "손익"); 이잉 = g(row, "이익잉여금"); 설명 = g(row, "설명")
        dr_acc = g(row, "차변계정"); cr_acc = g(row, "대변계정")
        added = False
        if dr_acc is not None and str(dr_acc).strip():
            cur["lines"].append({
                "side": "차변", "계정": str(dr_acc).strip(),
                "금액": g(row, "금액차변") if _is_amount(g(row, "금액차변")) else None,
                "손익": 손익 if _is_amount(손익) else None,
                "이익잉여금": 이잉 if _is_amount(이잉) else None,
                "설명": str(설명).strip() if 설명 is not None and str(설명).strip() else None,
            })
            added = True
        if cr_acc is not None and str(cr_acc).strip():
            cur["lines"].append({
                "side": "대변", "계정": str(cr_acc).strip(),
                "금액": g(row, "금액대변") if _is_amount(g(row, "금액대변")) else None,
                "손익": 손익 if _is_amount(손익) else None,
                "이익잉여금": 이잉 if _is_amount(이잉) else None,
                "설명": str(설명).strip() if 설명 is not None and str(설명).strip() else None,
            })
            added = True
        if not added and 설명 is not None and str(설명).strip():
            cur["notes"].append(str(설명).strip())

    if cur and (cur["lines"] or cur["notes"]):
        entries.append(cur)
    return entries

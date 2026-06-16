"""금융기관조회서 취합엑셀 INVESTMENT 시트 파서.

증권사 회신 1번 항목(유가증권 등 금융상품)에서 계좌별 잔액을 추출한다.
단기매매주식 등 단기금융상품의 **조회서 회신 소스**로 사용한다.

A-200에서 단기금융상품의 base는 거래처원장(단기매매증권 등)이며,
이 파서 결과는 금액 일치 시 구분/계좌/통화를 역으로 채우는 데 쓰인다.
(거래처원장에 단기금융 계정이 없는 회사는 해당 섹션이 비어 있을 수 있다.)

시트 구조: BANK/INSURANCE와 동일한 섹션 패턴 (마커 col1, 1번 섹션만 사용).
헤더(예): 조서번호 | 조회대상회사 | 사업자번호 | 금융기관명 | 조회기준일 |
          계좌번호 | 금융상품의 종류 | 금융상품 금액 | 예수금 | ...
"""

import re
from pathlib import Path

import openpyxl

from ._sections import find_section_bounds

_SYNONYMS: dict[str, str] = {
    "조서번호":       "조서번호",
    "금융기관명":     "금융기관명",
    "계좌번호":       "계좌번호",
    "금융상품의 종류": "금융상품종류",
    "금융상품의종류":  "금융상품종류",
    "금융상품 금액":   "금액",
    "금융상품금액":    "금액",
}

_REQUIRED = {"금융기관명", "금융상품종류", "금액"}

_DEFAULT_SHEET = "INVESTMENT"


def _map_headers(row: tuple) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(row):
        if cell is None:
            continue
        std = _SYNONYMS.get(str(cell).strip())
        if std and std not in mapping:
            mapping[std] = idx
    return mapping if _REQUIRED.issubset(mapping.keys()) else {}


def _to_amount(value) -> float:
    if isinstance(value, bool):
        return 0
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        s = re.sub(r"[^\d.\-]", "", value)
        try:
            return float(s) if s else 0
        except ValueError:
            return 0
    return 0


def parse_invest(path: str, sheet_name: str = _DEFAULT_SHEET) -> list[dict]:
    """INVESTMENT 1번 섹션의 계좌별 행을 반환한다.

    Returns:
        [{"조서번호","금융기관명","계좌번호","금융상품종류","금액"}, ...]
        (시트 없으면 빈 리스트 — 단기금융 미보유 회사 대응)
    """
    p = Path(path)
    wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return []

    rows = list(wb[sheet_name].iter_rows(values_only=True))
    wb.close()

    bounds = find_section_bounds(rows, "1.")
    if bounds is None:
        return []
    sec_idx, sec_end = bounds

    header_idx, col_map = None, None
    for i in range(sec_idx + 1, min(sec_idx + 6, sec_end)):
        cm = _map_headers(rows[i])
        if cm:
            header_idx, col_map = i, cm
            break
    if col_map is None:
        return []

    def g(row, key):
        idx = col_map.get(key)
        return row[idx] if idx is not None and idx < len(row) else None

    result = []
    for i in range(header_idx + 1, sec_end):
        row = rows[i]
        inst = g(row, "금융기관명")
        if inst is None or str(inst).strip() == "":
            continue
        result.append({
            "조서번호":     (str(g(row, "조서번호")).strip() if g(row, "조서번호") else None),
            "금융기관명":   str(inst).strip(),
            "계좌번호":     g(row, "계좌번호"),
            "금융상품종류": g(row, "금융상품종류"),
            "금액":         _to_amount(g(row, "금액")),
        })
    return result

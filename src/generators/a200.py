"""A-200 (조회서 대사) 생성기.

회사제시 명세서(거래처원장)를 base로 두고, 조회서 취합엑셀(BANK/INSURANCE)을
**역방향 매핑**해 STEP5 데이터 표를 구성한다.

핵심 원칙 (회계사 조언 반영):
  - 회사제시 금액(col8) = 거래처원장 잔액 (명세서가 모집단의 기준).
  - 거래처원장엔 계좌번호가 불완전하므로, **금액으로 조회서 회신금액과 1:1 매칭**해
    일치 시 구분/계좌번호/통화/이자율/만기를 조회서에서 역으로 채운다.
  - 금액 불일치분은 감사인 영역 → '확인필요' 표시만 붙이고 수치 판단은 하지 않는다.
  - 차이금액(col16) = 수식(=회신-회사제시) 온존. 일치여부 5개열 = 전부 공란.

분류:
  - 현금성자산: 거래처원장 보통예금  ↔  BANK 1번(현금성)
  - 장기금융상품: 거래처원장 장기금융상품(잔액>0)  ↔  INSURANCE 1번 그룹(해약환급금>0)
  - 단기금융상품(INVESTMENT)은 대상 회사에 따라 발생 가능 → 확장 훅만 둠

전자조회분만 취합엑셀에 존재한다. 우편조회분(명세서엔 있으나 취합엑셀에 없음)은
회신금액을 채울 수 없으므로 '확인필요(우편)'로 표시한다.
"""

import re
from collections import defaultdict
from copy import copy

from .base import WorkpaperGenerator


_CASH_LEDGER_ACCOUNTS = ("보통예금",)
_LONGTERM_LEDGER_KEYWORD = "장기금융"

_CONFIRM_ELECTRONIC = "전자"
_CONFIRM_POSTAL = "우편"
_V = "V"
_FLAG_REVIEW = "확인필요"


def _amt(v) -> int:
    """금액 셀을 정수로 안전 변환 (None/비숫자 → 0)."""
    if v is None or isinstance(v, bool):
        return 0
    if isinstance(v, (int, float)):
        return round(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return round(float(s)) if s else 0
    except ValueError:
        return 0


def _clean_inst_name(s) -> str:
    """금융기관명 정규화: 괄호 수식어 제거 + 공백 제거. 매칭 키 생성용."""
    if not s:
        return ""
    s = re.sub(r"\(.*?\)", "", str(s))
    return re.sub(r"\s+", "", s).strip()


# ---------------------------------------------------------------------------
# 현금성자산: 거래처원장 보통예금 ↔ BANK 현금성 (금액 1:1 역매핑)
# ---------------------------------------------------------------------------

def _build_cash_rows(ledger_cash: list[dict], bank_cash: list[dict]) -> tuple[list[dict], list[dict]]:
    """반환: (현금성 표 행들, 매칭 안 된 BANK 행들)."""
    # BANK를 금액별로 인덱싱 (1:1 소비)
    by_amt: dict[int, list[dict]] = defaultdict(list)
    for b in bank_cash:
        by_amt[_amt(b["금액"])].append(b)

    rows = []
    for L in ledger_cash:
        amt = _amt(L["잔액"])
        ledger_name = _clean_inst_name(L["거래처명"])

        chosen = None
        cands = by_amt.get(amt)
        if cands:
            # 같은 금액 후보 중 금융기관명이 부합하는 것 우선
            for i, b in enumerate(cands):
                bname = _clean_inst_name(b["금융기관명"])
                if bname and (bname in ledger_name or ledger_name in bname):
                    chosen = cands.pop(i)
                    break
            if chosen is None:
                chosen = cands.pop(0)

        rows.append(_cash_row(L, chosen))

    leftover = [b for lst in by_amt.values() for b in lst]
    return rows, leftover


def _cash_row(ledger: dict, bank: "dict | None") -> dict:
    """거래처원장 1행(+매칭된 BANK 행)을 A200 표준 행 dict로."""
    company_amt = _amt(ledger["잔액"])
    matched = bank is not None
    return {
        "계정분류":   "현금성자산",
        "금융기관":   (bank["금융기관명"] if matched else ledger["거래처명"]),
        "구분":       (bank["금융상품종류"] if matched else None),
        "계좌번호":   (bank["계좌번호"] if matched else None),
        "통화":       (bank["통화"] if matched else None),
        "금액":       company_amt,            # col8 회사제시
        "이자율":     (bank["연이자율"] if matched else None),
        "만기":       (bank["만기"] if matched else None),
        "사용제한":   "해당사항없음",
        "회신여부":   (_V if matched else None),
        "회수방법":   (_CONFIRM_ELECTRONIC if matched else None),
        "조회서번호": None,                    # 감사인 cross-ref 영역
        "회신금액":   (_amt(bank["금액"]) if matched else None),  # col15
        "_확인필요":  (not matched),           # 회신 미매칭 → 감사인 확인
    }


# ---------------------------------------------------------------------------
# 장기금융상품: 거래처원장 장기금융(잔액>0) ↔ INSURANCE 그룹(해약환급금>0)
# ---------------------------------------------------------------------------

def _build_longterm_rows(ledger_lt: list[dict], ins_groups: list[dict]) -> list[dict]:
    rows = []
    for L in ledger_lt:
        company_amt = _amt(L["잔액"])
        if company_amt <= 0:
            continue  # 0잔액(손해보험·해지펀드 등) 제외
        ledger_name = _clean_inst_name(L["거래처명"])

        match = None
        for g in ins_groups:
            gname = _clean_inst_name(g["금융기관명"])
            if gname and (ledger_name in gname or gname in ledger_name):
                match = g
                break

        matched = match is not None
        rows.append({
            "계정분류":   "장기금융상품",
            "금융기관":   L["거래처명"],
            "구분":       (match["보험의종류"] if matched else None),
            "계좌번호":   (match["증권번호"] if matched else None),
            "통화":       ("KRW" if matched else None),
            "금액":       company_amt,
            "이자율":     "-",
            "만기":       "-",
            "사용제한":   "해당사항없음",
            "회신여부":   (_V if matched else None),
            "회수방법":   (_CONFIRM_ELECTRONIC if matched else _CONFIRM_POSTAL),
            "조회서번호": None,
            "회신금액":   (_amt(match["조회서금액"]) if matched else None),
            "_확인필요":  (not matched),  # 우편조회분 → 감사인이 회신 별도 입력
        })
    return rows


# ---------------------------------------------------------------------------
# 공개 변환 API
# ---------------------------------------------------------------------------

def build_a200_data(ledger_rows: list[dict], bank_rows: list[dict],
                    insurance_groups: list[dict]) -> dict:
    """A-200 STEP5 데이터 + STEP2 모집단 요약을 계산해 반환한다.

    Returns dict:
      cash:        현금성자산 표 행들
      longterm:    장기금융상품 표 행들
      cash_leftover_bank:  조회서엔 회신됐으나 명세서와 미매칭된 BANK 행들 (감사인 확인)
      population_total:    모집단 총금액 (회사제시 금액 합)
    """
    ledger_cash = [r for r in ledger_rows if r.get("계정과목") in _CASH_LEDGER_ACCOUNTS]
    ledger_lt = [r for r in ledger_rows
                 if _LONGTERM_LEDGER_KEYWORD in (r.get("계정과목") or "")]
    bank_cash = [b for b in bank_rows if b.get("계정분류") == "현금성자산"]

    cash, leftover = _build_cash_rows(ledger_cash, bank_cash)
    longterm = _build_longterm_rows(ledger_lt, insurance_groups)

    population = sum(r["금액"] for r in cash) + sum(r["금액"] for r in longterm)
    return {
        "cash": cash,
        "longterm": longterm,
        "cash_leftover_bank": leftover,
        "population_total": population,
    }


# ---------------------------------------------------------------------------
# 엑셀 출력 (템플릿 = 완성본에서 데이터 제거한 작업용 파일)
# ---------------------------------------------------------------------------

class A200Generator(WorkpaperGenerator):
    """A-200 시트의 STEP2/STEP5/STEP6을 채운다.

    공통 4항목(회사명·Preparer·날짜·Reviewer)도 채운다.
    Step1·3·4 및 표 헤더(R42~44)는 템플릿 그대로 보존한다.
    """

    def fill(self, data: dict, params: dict) -> None:
        """data = build_a200_data() 결과, params = 회사명/preparer/reviewer/date 등."""
        self.open_template()
        cfg = self.config
        self._fill_header(cfg["header"], params)
        self._fill_step2(cfg["step2"], cfg.get("defaults", {}), data, params)
        self._fill_step5_step6(cfg, data)

    # --- 공통 4항목 ---
    def _fill_header(self, h: dict, params: dict) -> None:
        if params.get("회사명"):
            self.write_cell(h["company_cell"], f"{h.get('company_prefix','')}{params['회사명']}")
        if params.get("preparer"):
            self.write_cell(h["preparer_cell"], params["preparer"])
        if params.get("reviewer"):
            self.write_cell(h["reviewer_cell"], params["reviewer"])
        if params.get("date"):
            self.write_cell(h["date_cell"], params["date"])

    # --- Step2 모집단 ---
    def _fill_step2(self, s2: dict, defaults: dict, data: dict, params: dict) -> None:
        self.write_cell(s2["population_cell"], data["population_total"])
        self.write_cell(s2["source_name_cell"],
                        params.get("source_name") or defaults.get("source_name"))
        self.write_cell(s2["completeness_cell"],
                        params.get("completeness") or defaults.get("completeness"))

    # --- Step5 표 + Step6 결론 ---
    def _fill_step5_step6(self, cfg: dict, data: dict) -> None:
        s5 = cfg["step5"]
        cols = s5["columns"]
        start = s5["data_start_row"]
        diff_c = s5["diff_col"]
        conc_c = s5["conclusion_col"]
        match_cs = s5["match_cols"]
        hcol = s5["amount_company_col"]
        ocol = s5["amount_confirm_col"]
        review_flag = cfg.get("defaults", {}).get("review_flag", "확인필요")

        self._clear_region(start, s5["clear_until_row"])

        r = start
        no = 1

        def write_data_row(rec: dict, no_val: int):
            nonlocal r
            for key, col in cols.items():
                val = no_val if key == "No." else rec.get(key)
                self.ws[f"{col}{r}"] = val
            # 차이금액 = 수식 (회신 - 회사제시)
            self.ws[f"{diff_c}{r}"] = f"={ocol}{r}-{hcol}{r}"
            # 일치여부 5개열 = 공란 (감사인 영역)
            for mc in match_cs:
                self.ws[f"{mc}{r}"] = None
            # 결론 = 확인필요 플래그만
            self.ws[f"{conc_c}{r}"] = review_flag if rec.get("_확인필요") else None
            r += 1

        def write_subtotal(label: str, first: int, last: int):
            nonlocal r
            self._merge(s5["subtotal_label_merge"].format(r=r))
            self.ws[f"{cols['No.']}{r}"] = label
            if last >= first:
                self.ws[f"{hcol}{r}"] = f"=SUM({hcol}{first}:{hcol}{last})"
                self.ws[f"{ocol}{r}"] = f"=SUM({ocol}{first}:{ocol}{last})"
                self.ws[f"{diff_c}{r}"] = f"={ocol}{r}-{hcol}{r}"
            r += 1

        # 현금성자산
        cash_first = r
        for rec in data["cash"]:
            write_data_row(rec, no); no += 1
        write_subtotal("현금성자산 합계", cash_first, r - 1)

        # 장기금융상품
        lt_first = r
        for rec in data["longterm"]:
            write_data_row(rec, no); no += 1
        write_subtotal("장기금융상품 합계", lt_first, r - 1)

        # Step6 결론
        r += cfg.get("step6", {}).get("gap_after_subtotal", 2)
        self._write_conclusion(r, data, review_flag)

    def _write_conclusion(self, row: int, data: dict, review_flag: str) -> None:
        cash_review = sum(1 for x in data["cash"] if x["_확인필요"])
        lt_review = sum(1 for x in data["longterm"] if x["_확인필요"])
        leftover = len(data["cash_leftover_bank"])
        label_col = self.config["step5"]["columns"]["No."]

        lines = ["Step 6. 결론 (초안 — 감사인 검토 필요)"]
        lines.append(
            f"현금성자산: 명세서 {len(data['cash'])}건 중 조회서 회신금액과 1:1 대사. "
            f"금액 불일치 {cash_review}건은 {review_flag} (장부·회신 차이, 우편조회 등)."
        )
        if leftover:
            lines.append(
                f"조회서에는 회신되었으나 명세서 금액과 일치하지 않는 {leftover}건 존재 → 감사인 확인."
            )
        lines.append(
            f"장기금융상품: 명세서 {len(data['longterm'])}건. "
            f"우편조회분 {lt_review}건은 회신금액 별도 입력 필요({review_flag})."
        )
        for i, text in enumerate(lines):
            self.ws[f"{label_col}{row + i}"] = text

    # --- 유틸 ---
    def _clear_region(self, start: int, end: int) -> None:
        """데이터/합계/Step6 영역의 값·병합을 초기화."""
        for rng in list(self.ws.merged_cells.ranges):
            if rng.min_row >= start:
                self.ws.unmerge_cells(str(rng))
        for row in range(start, end + 1):
            for col in range(1, self.ws.max_column + 1):
                self.ws.cell(row=row, column=col).value = None

    def _merge(self, rng: str) -> None:
        try:
            self.ws.merge_cells(rng)
        except Exception:
            pass

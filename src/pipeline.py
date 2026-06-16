"""A-1 조서 일괄 생성 파이프라인.

입력 3종(발송 Control Sheet · 거래처원장 · 조회서 취합엑셀)을 파싱해
하나의 작업용 템플릿에 A-100/A-200/A-300 세 시트를 모두 채워 A-1을 완성한다.

체이닝 방식: 템플릿을 복사한 작업 파일에 세 생성기를 순차 적용한다.
각 생성기는 자기 시트만 수정하므로 다른 시트는 보존된다.

무시(감사인 수행영역): FN·월보 시트, 은행연합회 자료 열, 우편조회 회신금액.
"""

import shutil
import warnings
from pathlib import Path

from cache import cached_ideal_ledger
from extractors import (
    parse_cs,
    parse_bank, parse_bank_loans, parse_bank_collateral,
    parse_insurance, parse_invest,
)
from generators import (
    A100Generator, A200Generator, build_a200_data,
    A300Generator, build_a300_data,
    ControlSheetGenerator, build_ref_map,
)


class RunReport:
    """실행 진단 리포트. 입력/생성/출력/확인필요 단계별 항목을 모은다.

    엑셀을 못 보는 사용자가 '무엇이 입력값 문제이고 무엇이 출력 양식 문제인지' 알도록,
    그리고 일부가 실패해도 어디가 비었는지 알도록 한다 (절대 전체실패 금지 + 진단).
    """

    def __init__(self):
        self.entries: list[tuple] = []   # (stage, level, msg)  level: ok|warn|error

    def add(self, stage, level, msg):
        self.entries.append((stage, level, str(msg)))

    @property
    def has_error(self):
        return any(lv == "error" for _, lv, _ in self.entries)

    @property
    def has_warn(self):
        return any(lv in ("warn", "error") for _, lv, _ in self.entries)

    def render(self) -> str:
        icon = {"ok": "✓", "warn": "△", "error": "✗"}
        if self.has_error:
            summary = "부분 출력 — 일부 자료가 비었습니다([입력]의 ✗ 확인). 나머지는 정상 생성됨"
        elif self.has_warn:
            summary = "정상 출력 — 단, [확인필요]/빈 자료(△) 항목은 감사인 검토 필요"
        else:
            summary = "정상 출력"
        lines = ["=" * 60, f"  실행 리포트 — {summary}", "=" * 60]
        for stage in ("입력", "생성", "출력", "확인필요"):
            rows = [(lv, m) for st, lv, m in self.entries if st == stage]
            if not rows:
                continue
            lines.append(f"[{stage}]")
            for lv, m in rows:
                lines.append(f"  {icon.get(lv, '·')} {m}")
        return "\n".join(lines)


def build_a1(*, control_sheet, ledger_src, confirm_xlsx, template, config_dir,
             output, params, parsed_dir=None, ref_dir=None):
    """A-1 조서(A-100/200/300)를 한 파일로 생성한다.

    Args:
        control_sheet: 발송 Control Sheet 경로
        ledger_src:    잔액 소스 — 거래처원장(.xls/.xlsx) 또는 결산보고서(.xlsx).
                       파일명에 '결산보고서'가 있으면 결산보고서 변환 경로로 자동 라우팅.
        confirm_xlsx:  조회서 취합엑셀 경로 (BANK/INSURANCE/INVESTMENT)
        template:      작업용 템플릿(.xlsx, A100/A200/A300 시트 포함)
        config_dir:    셀 매핑 YAML 디렉터리 (a100/a200/a300/settlement_report.yaml)
        output:        결과 파일 경로
        params:        공통 헤더 {회사명, 날짜, preparer, reviewer} + 선택 항목
        parsed_dir:    변환자료 폴더 (이상적 양식 캐시 저장/재사용; None이면 출력 폴더 옆)

    Returns:
        RunReport (단계별 진단). 결과 파일은 output 경로에 생성됨(부분 실패해도 항상 생성).
    """
    cfg = Path(config_dir)
    out = Path(output)
    parsed = parsed_dir or str(out.parent / "변환자료")
    report = RunReport()

    # C-2차 학습 동의어 주입(오프라인·무료): 과거 해소된 헤더 매핑을 이번 실행의 헤더 인식에
    # 자동 적용한다. 캐시가 비었으면 무영향(기존 동작과 동일). LLM/키 불필요.
    try:
        from format_adapt import LearnedHeaders
        from extractors._headers import set_learned
        _lh = LearnedHeaders(parsed)
        for _pk in ("ledger", "insurance"):
            syn = _lh.synonyms_for(_pk)
            if syn:
                set_learned(_pk, syn)
                report.add("입력", "ok", f"학습 동의어 적용({_pk}): {len(syn)}건")
    except Exception:
        pass  # 학습 캐시는 보조 — 실패해도 본 흐름 막지 않음

    # 회복력: 회사마다 입력 구조가 달라 일부 파서가 실패할 수 있다. 한 파서 실패가
    # 전체 A-1 생성을 막지 않도록 감싸서, 해당 부분만 비우고 진단을 모은다(부분 출력 보장).
    def _safe_parse(label, fn, *args, default=None, **kwargs):
        try:
            with warnings.catch_warnings(record=True) as caught:
                warnings.simplefilter("always")
                res = fn(*args, **kwargs)
            for w in caught:                       # 파서 내부 경고(예: 잔액불일치) 포착
                m = str(w.message)
                if "default style" in m:
                    continue
                report.add("입력", "warn", f"{label}: {m.splitlines()[0][:160]}")
            n = len(res) if hasattr(res, "__len__") else "?"
            report.add("입력", "ok" if res else "warn",
                       f"{label}: {n}건" + ("" if res else " (빈 값 — 해당 부분 비움)"))
            return res
        except Exception as e:
            report.add("입력", "error", f"{label} 처리 실패({type(e).__name__}: {e}) → 해당 부분 비움")
            return [] if default is None else default

    # ---- STEP 1. 데이터 추출 (전부 안전 래핑) ----
    cs_rows = _safe_parse("발송 Control Sheet", parse_cs, control_sheet)
    ledger_rows = _safe_parse("거래처원장/결산보고서", cached_ideal_ledger, ledger_src, parsed, str(cfg))
    bank_rows = _safe_parse("BANK(현금성/퇴직)", parse_bank, confirm_xlsx)
    ins_groups = _safe_parse("INSURANCE(장기금융)", parse_insurance, confirm_xlsx)
    invest_rows = _safe_parse("INVESTMENT(단기금융)", parse_invest, confirm_xlsx)
    loan_rows = _safe_parse("대출(BANK 2-2)", parse_bank_loans, confirm_xlsx)
    coll_rows = _safe_parse("담보(BANK 9)", parse_bank_collateral, confirm_xlsx)
    ref_map = _safe_parse("취합 요약(조서번호)", build_ref_map, confirm_xlsx, default={})

    # 외화 환율 메모(참고자료) 로드 — 없으면 빈 dict(=환산 비활성, KRW 그대로)
    fx_rates = {}
    if ref_dir:
        from extractors import load_fx_rates
        fx_rates = _safe_parse("기말환율 메모(참고자료)", load_fx_rates, ref_dir, default={}) or {}

    # ---- STEP 2~4. 변환 ----
    a200_data = build_a200_data(ledger_rows, bank_rows, ins_groups, invest_rows, fx_rates=fx_rates)
    a300_data = build_a300_data(bank_rows, loan_rows, coll_rows)

    # 외화 기말환산 불일치 집계(진단)
    fx_diff = sum(1 for r in a200_data["cash"] if r.get("_기말환산불일치"))
    if fx_diff:
        report.add("확인필요", "warn",
                   f"외화 기말환산 불일치 {fx_diff}건(환율×조회서외화값 ≠ 명세서 원화) — A-0에서 조정")

    # 확인필요(감사인 영역) 집계
    cash_review = sum(1 for r in a200_data["cash"] if r["_확인필요"])
    lt_review = sum(1 for r in a200_data["longterm"] if r["_확인필요"])
    if cash_review or lt_review:
        report.add("확인필요", "warn",
                   f"A200 금액 불일치/우편 등 확인필요: 현금성 {cash_review}건, 장기금융 {lt_review}건")

    # ---- STEP 3. 조서 생성 (시트별 순차, 한 시트 실패가 나머지를 막지 않음) ----
    work = str(out)
    shutil.copy(template, work)

    def _safe_gen(sheet, gen_cls, yaml_name, fill_args):
        try:
            g = gen_cls(work, str(cfg / yaml_name))
            g.fill(*fill_args)
            g.save(work)
            report.add("생성", "ok", f"{sheet} 생성")
        except Exception as e:
            report.add("생성", "error", f"{sheet} 생성 실패({type(e).__name__}: {e}) → 이 시트는 양식 기본값")

    _safe_gen("A-100", A100Generator, "a100.yaml", (cs_rows, params))
    _safe_gen("A-200", A200Generator, "a200.yaml", (a200_data, params))
    _safe_gen("A-300", A300Generator, "a300.yaml", (a300_data, params))
    _safe_gen("Control Sheet", ControlSheetGenerator, "control_sheet.yaml",
              (control_sheet, ref_map, params))

    # ---- 양식무결성 HOOK(시각 확인 대체): 생성 후 격자·컬럼·제목 검사 ----
    try:
        for issue in _check_form_integrity(work, n_cs=len(cs_rows), config_dir=config_dir):
            report.add("출력", "warn", issue)
    except Exception as e:
        report.add("출력", "warn", f"양식무결성 검사 자체 실패({type(e).__name__}: {e})")

    return report


def _check_form_integrity(path, *, n_cs: int, config_dir: str) -> list:
    """생성된 A-1의 양식 무결성(격자·컬럼잘림·제목/헤더)을 검사하는 표준 게이트.

    양식 붕괴는 조서 신뢰성 훼손 → 데이터 채우기 후 반드시 이 검사를 거친다.
    config_dir은 build_a1가 받은 값을 그대로 사용(.exe frozen 모드 경로 안전).
    """
    from validator import (validate_grid, validate_a200_form,
                           check_column_fit, check_cells_present,
                           check_outside_col_fill)
    import yaml

    cfg = Path(config_dir)
    out = []
    # Control Sheet: 격자 + 컬럼 잘림 + 제목·헤더 존재 (시각 확인 대체)
    cs_cfg = yaml.safe_load((cfg / "control_sheet.yaml").read_text(encoding="utf-8"))
    cs_sheet = cs_cfg["sheet"]
    cs_start, cs_end = cs_cfg["data_start_row"], cs_cfg["data_start_row"] + n_cs - 1
    out += validate_grid(path, cs_sheet, header_row=3, data_start=cs_cfg["data_start_row"],
                         n_rows=n_cs, cols=list(range(1, 11)))
    out += check_column_fit(path, cs_sheet, cols=list(range(1, 11)),
                            data_start=cs_start, data_end=cs_end, header_row=3)
    out += check_cells_present(path, cs_sheet, [cs_cfg.get("title_cell", "A1"), "B3", "E3", "G3"])
    # A100 바깥열 회색 음영 보존 검사
    a1_cfg = yaml.safe_load((cfg / "a100.yaml").read_text(encoding="utf-8"))
    outside_a1 = a1_cfg.get("layout", {}).get("outside_gray_cols", [])
    if outside_a1:
        out += check_outside_col_fill(path, a1_cfg["sheet"], outside_gray_cols=outside_a1)
    # A200 전용(소계·수식·테두리) + 바깥열 회색 음영 보존 검사
    a2 = yaml.safe_load((cfg / "a200.yaml").read_text(encoding="utf-8"))
    out += validate_a200_form(path, a2["sheet"], a2)
    outside_a2 = a2.get("layout", {}).get("outside_gray_cols", [])
    if outside_a2:
        out += check_outside_col_fill(path, a2["sheet"], outside_gray_cols=outside_a2)
    return out

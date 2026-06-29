# -*- coding: utf-8 -*-
"""회사별 워크스페이스 폴더 해석 — 모든 easy_run 진입점이 공유.

한 자동화프로그램 폴더 안에서 여러 회사를 깔끔히 관리하기 위한 공통 계층.
지금까지 7개 진입점이 각자 INPUT_DIR/PARSED_DIR/... 블록과 회사명 손입력을
복제했는데, 이 모듈 하나로 통일한다.

배포 루트 구조:
  자동화프로그램/
    양식자료/            ← 공유(회사 무관): 조서 템플릿 (회사마다 복제하지 않음)
    _internal/config/    ← 공유: 셀 매핑 (스크립트=폴더, .exe=내장 _MEIPASS)
    작업/
      {회사}/            ← 회사명 = 폴더명(자동 인식, 손입력 불필요)
        입력자료/
        변환자료/        ← 캐시. 회사별로 격리되어 stale·충돌 원천 차단
        참고자료/        ← 전기조서·환율·내용연수 메모(회사별)
        출력조서/

환경변수 재정의:
  A1_WORK_DIR      작업 루트(기본 BASE/작업)
  A1_TEMPLATE_DIR  양식자료(공유)
  A1_INPUT_DIR / A1_PARSED_DIR / A1_REF_DIR / A1_OUTPUT_DIR
       → 하나라도 설정되면 '레거시 평면 모드'(작업/ 미사용, 그 경로 직접 사용).
         dev·테스트·CI 호환용. 회사명은 그대로 입력받는다.
"""

import json
import os
import sys
from pathlib import Path


def base_dir() -> Path:
    """프로그램 루트. .exe(frozen)면 실행파일 폴더, 스크립트면 프로젝트 루트."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def template_dir() -> Path:
    """조서 템플릿 폴더(회사 무관 공유). dev에서 비어 있으면 _internal/양식 폴백."""
    base = base_dir()
    tdir = Path(os.environ.get("A1_TEMPLATE_DIR", base / "양식자료"))
    if not getattr(sys, "frozen", False) and not list(tdir.glob("*.xls*")):
        dev = base / "_internal" / "양식"
        if dev.exists():
            return dev
    return tdir


def config_dir() -> Path:
    """셀 매핑 config 폴더(공유). 스크립트=_internal/config, .exe=내장(_MEIPASS)."""
    return Path(getattr(sys, "_MEIPASS", str(base_dir()))) / "_internal" / "config"


def work_root() -> Path:
    return Path(os.environ.get("A1_WORK_DIR", base_dir() / "작업"))


INPUT_README_NAME = "★먼저읽기_입력파일안내.txt"

INPUT_README_TEXT = """\
================================================================
  이 폴더(입력자료)에 원본 엑셀을 넣으세요
================================================================

파일 이름을 아래 '역할 한 단어'로 바꿔 넣으면 자동 인식됩니다.
(역할 단어만 들어가면 되고, 뒤에 회사명·날짜를 붙여도 됩니다.
 예: 조회서_○○_2512.xlsx)

  정산표      ← 정산표 (대부분의 조서 생성 기준)
  조회서      ← 금융기관 조회서 취합엑셀 (회신 취합본)
  잔액        ← 거래처잔액 현황(거래처원장) 또는 결산보고서 — 둘 중 하나만
  분개장      ← 분개장
  발송CS      ← 발송 Control Sheet
  관리대장    ← 고정자산 관리대장 (유무형자산용)
  개별명세_A1 ← 개별 잔액명세서(현금성 계좌표 등) — A-1 보강용(선택)

[조서별 필요 파일]
  조서생성(16개 총괄표) … 정산표
  A-0 현금 …………… 정산표 (+조회서·잔액)
  A-1 금융기관 조회 … 발송CS · 조회서 · 잔액 (+개별명세_A1)
  BBDD 차입금 ……… 정산표 · 조회서 · 잔액 · 분개장
  G 유무형자산 …… 정산표 · 관리대장 (+분개장)
  R-S 판관비·기타손익 … 정산표
  기타상세(D·CC) … 정산표 · 잔액

※ 예전 이름(필수자료1_발송CS / 필수자료2_조회서 / 필수자료3_잔액 등)도 인식합니다.
※ 파일을 다 넣은 뒤 '조서생성기'를 실행하세요.
================================================================
"""


def _ensure_input_readme(input_dir: Path) -> None:
    """입력자료 폴더에 안내문을 넣는다(없을 때만 — 사용자 수정 보존)."""
    try:
        readme = Path(input_dir) / INPUT_README_NAME
        if not readme.exists():
            readme.write_text(INPUT_README_TEXT, encoding="utf-8")
    except Exception:
        pass


# 참고자료 시드 — A-0(외화 환산)·G(유무형 상각) 참조용 템플릿. 인차지가 값만 채운다.
# (회사 기밀 아님 — 양식·예시·작성안내. 파서가 읽는 표준 포맷.)
REF_SEEDS = {
    "기말환율.txt": """\
# 기말환율 메모 — 외화 계좌 환산용
#
# 사용법: 외화 계좌가 있으면, 아래에 '통화코드=환율' 한 줄씩 적으세요.
#   · 환율 = 기말 기준, 1 외화당 원화 금액
#   · '#' 로 시작하는 줄과 빈 줄은 무시됩니다.
#   · 외화 계좌가 없으면 이 파일은 없어도 됩니다.
#
# 아래는 예시(가상) 값입니다. 실제 기말환율로 바꿔 적으세요.

USD=1434.90
EUR=1521.30
JPY=9.45
CNY=196.30
""",
    "유무형자산_내용연수_상각방법.txt": """\
# ─────────────────────────────────────────────────────────────
#  유무형자산 세부계정 · 내용연수 · 상각방법 메모
#  (G 유무형자산 조서 자동화 입력 — 인차지가 작성)
# ─────────────────────────────────────────────────────────────
#  · 한 줄에 자산 하나.  '#'으로 시작하는 줄과 빈 줄은 무시됩니다.
#  · 컬럼 구분은 파이프( | ).  앞뒤 공백은 자동 정리됩니다.
#  · 구분    : 유형 / 무형
#  · 계정과목 : 총괄표·정산표에 나오는 자산 계정명 (예: 건물, 차량운반구, 소프트웨어)
#  · 내용연수 : 숫자(년).  '5' 또는 '5년' 모두 가능
#  · 상각방법 : 정액 / 정률.  둘 다 적으면(예: 정률,정액) 회사제시 금액을 보고 자동 추론
#    ※ 한 분류 안에 정액·정률이 섞이는 경우가 많습니다(예: 차량·비품·시설장치).
#      섞일 수 있으면 '정률,정액'으로 적으면 자산별로 알아서 맞춰 재계산합니다.
#
#  구분 | 계정과목      | 내용연수 | 상각방법
# ─────────────────────────────────────────────────────────────
유형 | 건물        | 40 | 정액
유형 | 구축물      | 20 | 정액
유형 | 차량운반구  | 5  | 정률,정액
유형 | 공구와기구  | 5  | 정률,정액
유형 | 비품        | 5  | 정률,정액
유형 | 시설장치    | 5  | 정률,정액
무형 | 소프트웨어  | 5  | 정액
""",
}


def _ensure_ref_seeds(ref_dir: Path) -> None:
    """참고자료 폴더에 환율·상각 템플릿을 넣는다(없을 때만 — 사용자 값 보존)."""
    for name, text in REF_SEEDS.items():
        try:
            f = Path(ref_dir) / name
            if not f.exists():
                f.write_text(text, encoding="utf-8")
        except Exception:
            pass


def workspace_for(company: str) -> "Workspace":
    """회사명으로 작업/{회사}/ 4폴더 + 입력자료 안내문 + 참고자료 시드를 만들어 반환."""
    cdir = work_root() / company
    ws = Workspace(company,
                   cdir / "입력자료",
                   cdir / "변환자료",
                   cdir / "참고자료",
                   cdir / "출력조서").ensure()
    _ensure_input_readme(ws.input_dir)
    _ensure_ref_seeds(ws.ref_dir)
    return ws


# ── 활성 선택 오버라이드 (오케스트레이터용) ─────────────────────────────────
# 통합 실행기(조서생성기)가 회사·회사정보를 한 번만 고른 뒤, 각 조서 모듈의
# main()을 그대로 호출해도 다시 묻지 않고 그 선택을 조용히 재사용하게 한다.
_ACTIVE_WS = None
_ACTIVE_META = None


def push_active(ws: "Workspace", meta: dict) -> None:
    global _ACTIVE_WS, _ACTIVE_META
    _ACTIVE_WS, _ACTIVE_META = ws, dict(meta or {})


def clear_active() -> None:
    global _ACTIVE_WS, _ACTIVE_META
    _ACTIVE_WS = _ACTIVE_META = None


class Workspace:
    """한 회사의 작업공간. 입력·변환(캐시)·참고·출력 4폴더를 회사별로 격리한다."""

    def __init__(self, company: str, input_dir: Path, parsed_dir: Path,
                 ref_dir: Path, output_dir: Path):
        self.company = company
        self.input_dir = Path(input_dir)
        self.parsed_dir = Path(parsed_dir)
        self.ref_dir = Path(ref_dir)
        self.output_dir = Path(output_dir)

    def ensure(self) -> "Workspace":
        for d in (self.input_dir, self.parsed_dir, self.ref_dir, self.output_dir):
            d.mkdir(parents=True, exist_ok=True)
        return self

    def __repr__(self):
        return f"<Workspace {self.company!r} @ {self.output_dir.parent}>"


def _is_legacy_env() -> bool:
    return any(os.environ.get(k) for k in
               ("A1_INPUT_DIR", "A1_PARSED_DIR", "A1_REF_DIR", "A1_OUTPUT_DIR"))


def _legacy_workspace(company: str) -> Workspace:
    """레거시 평면 모드(env 재정의). 변환·출력은 회사 하위로 격리(기존 동작 유지)."""
    base = base_dir()
    inp = Path(os.environ.get("A1_INPUT_DIR", base / "입력자료"))
    parsed = Path(os.environ.get("A1_PARSED_DIR", base / "변환자료"))
    ref = Path(os.environ.get("A1_REF_DIR", base / "참고자료"))
    out = Path(os.environ.get("A1_OUTPUT_DIR", base / "출력조서"))
    return Workspace(company, inp, parsed / company, ref, out / company).ensure()


def _flat_root_workspace(company: str) -> Workspace:
    """기존 평면 배포(작업/ 도입 전) 호환: 루트 5폴더를 그대로 쓰되 변환·출력만 회사 격리."""
    base = base_dir()
    return Workspace(company,
                     base / "입력자료",
                     base / "변환자료" / company,
                     base / "참고자료",
                     base / "출력조서" / company).ensure()


def _ask_company(ask, line) -> str:
    while True:
        v = ask("  회사명(폴더로 사용, 예: ○○_2512): ").strip()
        if v:
            return v
        line("  ▷ 회사명을 입력해 주세요.")


def select_workspace(line=print, ask=None) -> "Workspace | None":
    """회사 워크스페이스를 대화형으로 선택/생성해 반환.

    작업/ 하위 회사 폴더를 번호 메뉴로 보여주고, 번호 선택 또는 새 회사명 입력.
    회사명은 폴더명이 되어 자동 인식된다(손입력 불일치 제거).
    레거시 env 재정의 시엔 회사명만 묻고 그 경로를 사용한다.
    """
    if ask is None:
        ask = input

    if _ACTIVE_WS is not None:   # 오케스트레이터가 이미 고른 회사 — 재선택 생략
        return _ACTIVE_WS

    if _is_legacy_env():
        return _legacy_workspace(_ask_company(ask, line))

    root = work_root()
    root.mkdir(parents=True, exist_ok=True)
    companies = sorted(p.name for p in root.iterdir() if p.is_dir())

    # 기존 평면 배포(작업/ 비어 있고 루트 입력자료에 파일이 있음) 폴백 옵션
    flat_inp = base_dir() / "입력자료"
    has_flat = (not companies) and flat_inp.exists() and \
        any(flat_inp.glob("*.xls*"))

    line("\n[회사 선택]  (작업/ 폴더 안에서 회사별로 격리 관리)")
    for i, name in enumerate(companies, 1):
        line(f"  {i}) {name}")
    if has_flat:
        line("  0) 기존 입력자료 그대로 사용(루트 평면 폴더)")
    line("  · 새 회사를 추가하려면 회사명을 그대로 입력하세요(예: ○○_2512)")

    while True:
        sel = ask("\n  번호 또는 새 회사명: ").strip()
        if not sel:
            line("  ▷ 번호나 회사명을 입력해 주세요.")
            continue
        if has_flat and sel == "0":
            return _flat_root_workspace(_ask_company(ask, line))
        if sel.isdigit():
            idx = int(sel)
            if 1 <= idx <= len(companies):
                company = companies[idx - 1]
                break
            line(f"  ▷ 1~{len(companies)} 사이 번호를 입력해 주세요(또는 새 회사명).")
            continue
        # 숫자가 아니면 새 회사명으로 간주
        company = sel
        break

    return workspace_for(company)


# ── 회사정보(기준일·작성자·검토자) 1회 저장 → 모든 EXE가 재사용 ──────────────
# 같은 감사면 회사마다 기준일·작성자·검토자가 고정인데 EXE 7개에서 매번 재입력하는
# 낭비를 없앤다. 회사 워크스페이스에 저장해 두고, 다음 실행부터는 Enter만 치면 된다.

def _meta_path(ws: "Workspace") -> Path:
    # 변환자료는 모든 모드에서 회사별로 격리되므로 회사정보 보관처로 안전.
    return ws.parsed_dir / "_회사정보.json"


def read_company_meta(ws: "Workspace") -> dict:
    try:
        return json.loads(_meta_path(ws).read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_company_meta(ws: "Workspace", meta: dict) -> None:
    try:
        ws.parsed_dir.mkdir(parents=True, exist_ok=True)
        _meta_path(ws).write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def prompt_company_info(ws: "Workspace", ask=None, line=print,
                        *, preparer_required: bool = True) -> tuple:
    """기준일·작성자·검토자를 묻되, 회사에 저장된 값이 있으면 기본값으로 제시(Enter=유지).

    반환: (기준일, 작성자, 검토자). 입력 후 회사 워크스페이스에 저장 → 다음 EXE부터 자동.
    """
    if ask is None:
        ask = input
    if _ACTIVE_META is not None:   # 오케스트레이터가 이미 받은 회사정보 — 재입력 생략
        return (_ACTIVE_META.get("date", ""), _ACTIVE_META.get("preparer", ""),
                _ACTIVE_META.get("reviewer", ""))
    meta = read_company_meta(ws)

    def _one(label, key, example, required):
        saved = (meta.get(key) or "").strip()
        if saved:
            hint = f" [Enter=저장값 '{saved}']"
        else:
            hint = f" (예: {example})" if example else ""
        while True:
            v = ask(f"  {label}{hint}: ").strip()
            if not v:
                v = saved
            if v or not required:
                return v
            line("  ▷ 값을 입력해 주세요.")

    if any(meta.get(k) for k in ("date", "preparer", "reviewer")):
        line("  (저장된 회사정보를 불러왔습니다 — 그대로 쓰려면 Enter)")
    date = _one("기준일", "date", "2025-12-31", True)
    preparer = _one("작성자(Preparer)", "preparer", "CJH", preparer_required)
    reviewer = _one("검토자(Reviewer)", "reviewer", "KHK", False)
    write_company_meta(ws, {"date": date, "preparer": preparer, "reviewer": reviewer})
    return date, preparer, reviewer

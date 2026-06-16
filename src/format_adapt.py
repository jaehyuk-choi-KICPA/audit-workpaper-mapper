# -*- coding: utf-8 -*-
"""포맷 호환 C-2차 — 마스킹 LLM 부분보정 + 학습 동의어 캐시.

C-1차(정확+부분일치 동의어, `extractors/_headers.py`)로도 *필수 헤더*를 못 찾은 경우에만,
**그 헤더 행 한 줄**을 좁게 LLM에 물어 매핑을 보정한다(전체 파일 재해석 X → 토큰 최소·자가치유).

3겹 안전망 (CLAUDE.md 설계원칙 #5 / 플랜 C):
  1) 학습 캐시(오프라인·무료): 과거에 해소된 매핑(헤더라벨→필드)을 먼저 적용 → LLM 없이 자동 인식.
  2) 마스킹 LLM(미해결 필드만): 헤더 라벨(구조 메타, 비민감)은 평문, **셀 값은 전부 마스킹**해
     "이 중 '잔액'은 몇 번 열?" 식 좁은 질문. 답은 코드가 검증(인덱스 범위·열 타입) 후 사용.
  3) 학습 저장(복리): LLM/사용자가 해소한 매핑을 캐시에 저장 → 다음 실행부터 오프라인 인식.

보안 모델 (엄수):
  - **API 키는 코드/exe에 내장 금지.** 외부 `.env`/환경변수(ANTHROPIC_API_KEY)에서만 로드.
    키 없으면 LLM 비활성 = 100% 오프라인(1·3겹만 동작). exe 배포 ≠ 키 배포.
  - **egress 전 마스킹 필수**: 실제 셀 값은 절대 전송하지 않는다(타입·자릿수 메타만). `mask_cell` 참조.
    이 모듈은 `build_payload`가 만든 페이로드만 전송하며, 페이로드엔 원본 값 문자열이 없음을
    `test_masking`(실데이터로 0 누출 검증)이 보증한다.
  - **동의(consent) + 전송내용 표시**: 호출 전 무엇이 나가는지 보여주고 Y/N. 결과는 'AI 추정—확인필요' 표기.
"""

import json
import os
import re
from pathlib import Path


# ---------------------------------------------------------------------------
# 값 타입 추론 + 마스킹 (egress 전 — 실값 0 누출)
# ---------------------------------------------------------------------------

_LONG_DIGIT_RE = re.compile(r"^[\d\-]{6,}$")  # 계좌·증권번호류(6자리+ 숫자/하이픈)


def col_type(values) -> str:
    """열의 대표 타입 추론: number | long_digit | text | empty.

    잔액 등 금액은 number, 계좌·증권번호는 long_digit(긴 숫자열), 거래처명 등은 text.
    """
    seen = [v for v in values if v is not None and str(v).strip() != ""]
    if not seen:
        return "empty"
    n_num = n_long = n_text = 0
    for v in seen:
        if isinstance(v, bool):
            n_text += 1
        elif isinstance(v, (int, float)):
            n_num += 1
        else:
            s = str(v).strip()
            digits = re.sub(r"[,\s]", "", s)
            if _LONG_DIGIT_RE.match(digits):
                n_long += 1
            elif re.fullmatch(r"-?[\d,]+(\.\d+)?", s):
                n_num += 1
            else:
                n_text += 1
    # 다수결 (long_digit는 number보다 우선해 계좌번호를 금액으로 오인하지 않도록)
    if n_long >= max(n_num, n_text) and n_long > 0:
        return "long_digit"
    if n_num >= n_text and n_num > 0:
        return "number"
    return "text"


def mask_cell(value, ctype: str) -> str:
    """셀 값을 타입·대략적 형태만 남기고 마스킹. 원본 문자/숫자는 절대 남기지 않는다.

    number    → 자릿수만('#'*n)         예: 1,234,567 → '#######'
    long_digit→ 자릿수만('#'*n)         예: 110-1234-5678 → '###########'
    text      → 길이만('○'*min(n,8))    예: '(주)○○○○' → '○○○○○○'
    empty/None→ '∅'
    """
    if value is None or str(value).strip() == "":
        return "∅"
    s = str(value).strip()
    if ctype in ("number", "long_digit"):
        n = len(re.sub(r"[^\d]", "", s)) or 1
        return "#" * min(n, 16)
    return "○" * min(len(s), 8)


def build_payload(header_cells, sample_rows, unmapped_fields, *, max_samples=4):
    """LLM에 보낼 **마스킹된** 페이로드 생성.

    Args:
        header_cells:    헤더 행 셀(평문 라벨 — 구조 메타, 비민감).
        sample_rows:     헤더 아래 데이터 표본 행들(값은 마스킹됨).
        unmapped_fields: 아직 못 찾은 표준 필드명 리스트(예: ['잔액','거래처명']).
        max_samples:     열당 마스킹 표본 수.

    Returns:
        {"columns":[{idx, header, type, samples:[masked...]}, ...],
         "needed": [...필드...]}  — 원본 값 문자열은 포함되지 않는다(검증: test_masking).
    """
    ncol = len(header_cells)
    columns = []
    for idx in range(ncol):
        raw_vals = [row[idx] for row in sample_rows if idx < len(row)]
        ctype = col_type(raw_vals)
        masked = [mask_cell(v, ctype) for v in raw_vals[:max_samples]]
        header = header_cells[idx]
        columns.append({
            "idx": idx,
            "header": ("" if header is None else str(header).strip()),
            "type": ctype,
            "samples": masked,
        })
    return {"columns": columns, "needed": list(unmapped_fields)}


# ---------------------------------------------------------------------------
# 학습 동의어 캐시 (복리 — 오프라인)
# ---------------------------------------------------------------------------

class LearnedHeaders:
    """해소된 헤더 매핑을 로컬에 누적. {parser_key: {정규화헤더라벨: 표준필드}}.

    변환자료 폴더 옆에 learned_headers.json으로 저장 → 다음 실행부터 오프라인 인식.
    """

    FILENAME = "learned_headers.json"

    def __init__(self, store_dir):
        self.path = Path(store_dir) / self.FILENAME
        self.data: dict = {}
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except Exception:
                self.data = {}

    def synonyms_for(self, parser_key: str) -> dict:
        """C-1차 _headers.map_headers의 synonyms에 합칠 학습 동의어(평문 라벨→필드)."""
        return dict(self.data.get(parser_key, {}))

    def learn(self, parser_key: str, label: str, field: str):
        from extractors._headers import normalize
        self.data.setdefault(parser_key, {})[normalize(label)] = field

    def save(self):
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2),
                                 encoding="utf-8")
        except Exception:
            pass  # 캐시는 보조 — 실패해도 본 흐름 막지 않음


# ---------------------------------------------------------------------------
# 키 로드 + LLM 질의 (키 없으면 비활성 = 오프라인)
# ---------------------------------------------------------------------------

def load_api_key() -> "str | None":
    """외부 환경/.env에서만 키 로드. 없으면 None(LLM 비활성). 코드/exe 내장 금지."""
    key = os.environ.get("ANTHROPIC_API_KEY")
    if key:
        return key
    try:
        from dotenv import dotenv_values
        for cand in (Path.cwd() / ".env", Path(__file__).resolve().parent.parent / ".env"):
            if cand.exists():
                v = dotenv_values(str(cand)).get("ANTHROPIC_API_KEY")
                if v:
                    return v
    except Exception:
        pass
    return None


_MODEL = "claude-opus-4-8"


def suggest_via_llm(payload: dict, *, api_key: str, model: str = _MODEL) -> dict:
    """마스킹 페이로드로 미해결 필드의 열 인덱스를 LLM에 질의(좁은 질문).

    Returns: {필드: 열인덱스(int)} — 호출자가 인덱스 범위·타입을 다시 검증해야 함.
    네트워크/파싱 실패는 빈 dict 반환(폴백 — 본 흐름 막지 않음).
    """
    try:
        import anthropic
    except ImportError:
        return {}

    cols_desc = "\n".join(
        f"  열{c['idx']}: 라벨='{c['header']}' 타입={c['type']} 표본={c['samples']}"
        for c in payload["columns"]
    )
    needed = payload["needed"]
    prompt = (
        "다음은 한 회계 표의 헤더와 (마스킹된) 표본이다. 실제 값은 전송되지 않았고 "
        "'#'=숫자자릿수, '○'=문자길이, '∅'=빈칸이다.\n"
        f"{cols_desc}\n\n"
        f"다음 표준 필드 각각이 몇 번 열인지 판단하라(라벨 의미+타입 기준): {needed}\n"
        "확신이 없으면 null. JSON만 출력: {\"필드명\": 열번호(int) 또는 null, ...}"
    )
    try:
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model=model, max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )
        text = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        m = re.search(r"\{.*\}", text, re.S)
        raw = json.loads(m.group(0)) if m else {}
    except Exception:
        return {}

    out = {}
    for field, idx in raw.items():
        if isinstance(idx, int):
            out[field] = idx
    return out


def resolve_unmapped(*, parser_key, header_cells, sample_rows, unmapped_fields,
                     expected_types=None, learned: "LearnedHeaders | None" = None,
                     api_key=None, consent=None, announce=None):
    """미해결 필드를 학습캐시→(동의 시)마스킹 LLM 순으로 보정. 검증 통과분만 반환.

    Args:
        parser_key:     학습 캐시 분류 키(예: 'ledger', 'insurance').
        expected_types: {필드: 허용타입집합} — LLM 답 검증용(예: {'잔액':{'number'}}).
        learned:        LearnedHeaders(없으면 학습 단계 생략).
        api_key:        없으면 LLM 비활성(오프라인).
        consent:        callable(payload)->bool. egress 전 동의. None이면 거부(전송 안 함).
        announce:       callable(str) — 전송 내용/상태 안내 출력(선택).

    Returns:
        ({필드: 열인덱스}, source) source∈{'learned','llm','none'}.
    """
    # 학습 캐시는 호출자(map_headers)가 synonyms로 이미 합치므로, 여기 도달 = 캐시로도 미해결.
    if not unmapped_fields:
        return {}, "none"
    if not api_key:
        if announce:
            announce("LLM 보정 비활성(API 키 없음) — 오프라인 진단만 제공")
        return {}, "none"

    payload = build_payload(header_cells, sample_rows, unmapped_fields)
    if announce:
        announce("아래 마스킹 정보만 전송됩니다(실제 값 없음):\n" +
                 json.dumps(payload, ensure_ascii=False, indent=2))
    if consent is None or not consent(payload):
        if announce:
            announce("사용자 미동의 → 전송하지 않음(오프라인 진단만)")
        return {}, "none"

    suggested = suggest_via_llm(payload, api_key=api_key)

    # 코드 검증: 인덱스 범위 + (기대 타입이 있으면) 열 타입 일치
    col_by_idx = {c["idx"]: c for c in payload["columns"]}
    verified = {}
    for field, idx in suggested.items():
        if field not in unmapped_fields or idx not in col_by_idx:
            continue
        if expected_types and field in expected_types:
            if col_by_idx[idx]["type"] not in expected_types[field]:
                continue  # 타입 불일치 → 신뢰 안 함
        verified[field] = idx
        if learned is not None:
            learned.learn(parser_key, col_by_idx[idx]["header"], field)
    if learned is not None and verified:
        learned.save()
    return verified, ("llm" if verified else "none")

"""헤더 매핑 공용 유틸 — 유연 탐지(C-1차).

회사마다 헤더 라벨이 조금씩 달라도(공백·접두/접미·동의어) 흡수하기 위한 2단계 매칭:
  1차) 정확 동의어 매칭 — 기존 동작 그대로(검증됨). 회귀 위험 0.
  2차) 부분일치(substring) 폴백 — 1차로 못 찾은 *필수* 필드만, 미사용 열에서 키워드 포함 검색.

값-타입 폴백(잔액=숫자열, 계좌=장수숫자 등)은 금액열이 여러 개라 라벨 없이 구분이 어렵다 →
여기서는 라벨 기반(정확+부분일치)까지만 책임지고, 라벨이 완전히 다른 경우는 진단/LLM(C-2차)에 넘긴다.

C-2차(학습 동의어 캐시)는 이 모듈의 synonyms/sub_keywords에 매핑을 주입하는 식으로 확장한다.
"""

import re


def normalize(s) -> str:
    """컬럼명 정규화: 내부 공백 제거 + strip. '거 래 처 명' → '거래처명'."""
    return re.sub(r"\s+", "", str(s).strip()) if s is not None else ""


# 학습 동의어 레지스트리(C-2차) — {parser_key: {정규화라벨: 표준필드}}.
# format_adapt가 해소·저장한 매핑을 파이프라인이 여기에 주입하면, map_headers의 1차 정확매칭이
# 학습분까지 자동 인식한다(LLM 없이 오프라인). 키 없으면 비어 있어 기존 동작과 동일.
_LEARNED: dict = {}


def set_learned(parser_key: str, mapping: dict):
    """학습 동의어를 등록(누적). mapping: {정규화라벨: 표준필드}."""
    _LEARNED.setdefault(parser_key, {}).update(mapping or {})


def map_headers(cells, *, synonyms, required, sub_keywords=None, norm=normalize,
                parser_key=None):
    """헤더 셀들을 {표준필드: 열인덱스}로 매핑.

    Args:
        cells:        헤더 행(셀 값 시퀀스).
        synonyms:     {정규화된_라벨: 표준필드} — 1차 정확 매칭.
        required:     필수 표준필드 집합. 미충족이면 None 반환(예외 없음).
        sub_keywords: [(표준필드, [키워드...]), ...] — 2차 부분일치 폴백(구체적 우선).
                      None이면 폴백 생략.
        norm:         정규화 함수(기본 normalize).
        parser_key:   주어지면 _LEARNED[parser_key]의 학습 동의어(C-2차)를 1차 매칭에 합침.

    Returns:
        매핑 dict (required 충족 시) 또는 None.
    """
    if parser_key and _LEARNED.get(parser_key):
        synonyms = {**_LEARNED[parser_key], **synonyms}   # 기본 동의어 우선(학습분은 빈자리만 채움)

    indexed = [(idx, norm(c)) for idx, c in enumerate(cells) if c is not None and norm(c)]

    mapping: dict = {}
    used: set = set()
    # 1차 — 정확 매칭
    for idx, n in indexed:
        std = synonyms.get(n)
        if std and std not in mapping:
            mapping[std] = idx
            used.add(idx)

    # 2차 — 부분일치 폴백: 못 찾은 필수 필드만, 미사용 열에서
    for field, kws in (sub_keywords or []):
        if field in mapping:
            continue
        for kw in kws:
            kwn = norm(kw)
            hit = next((idx for idx, n in indexed if idx not in used and kwn in n), None)
            if hit is not None:
                mapping[field] = hit
                used.add(hit)
                break

    return mapping if set(required).issubset(mapping.keys()) else None

"""파일 해시 기반 파싱 캐시 유틸."""

import json
from hashlib import sha256
from pathlib import Path
from typing import Callable


def file_hash(path: str) -> str:
    """파일 내용의 SHA-256 앞 16자리. 내용이 바뀌면 해시도 바뀐다."""
    return sha256(Path(path).read_bytes()).hexdigest()[:16]


def load_with_cache(path: str, parse_fn: Callable, **kwargs) -> list | dict:
    """엑셀 파일을 파싱하고 결과를 JSON으로 캐시한다.

    파일 내용의 SHA-256 앞 16자리를 캐시 키로 사용한다.
    파일이 변경되면 해시가 달라지므로 자동으로 재파싱된다.
    캐시는 _internal/cache/ 에 저장되어 git 추적에서 제외된다.

    사용 예:
        rows = load_with_cache(path, parse_ledger)
        rows = load_with_cache(path, parse_cs, sheet_name="My Sheet")
    """
    cache_dir = Path("_internal/cache")
    cache_dir.mkdir(parents=True, exist_ok=True)

    cache_file = cache_dir / f"{Path(path).stem}_{file_hash(path)}.json"

    if cache_file.exists():
        return json.loads(cache_file.read_text(encoding="utf-8"))

    result = parse_fn(path, **kwargs)
    cache_file.write_text(
        json.dumps(result, ensure_ascii=False, default=str),
        encoding="utf-8",
    )
    return result


def cached_ideal_ledger(source_path: str, parsed_dir: str, config_dir: str) -> list[dict]:
    """잔액 소스(거래처원장/결산보고서)를 이상적 양식으로 변환하되,
    소스 파일 해시로 캐시된 결과를 재사용한다.

    `변환자료/{원본명}__ideal_{hash}.xlsx` 가 있으면 재파싱 없이 그대로 읽는다.
    원본이 바뀌면 해시가 달라져 자동 재변환된다. 여러 조서 생성기가 이 함수를
    호출하면 **동일 변환 스냅샷을 공유**하므로 일관성·속도가 확보된다.

    Args:
        source_path: 원본 엑셀 (거래처원장 .xls/.xlsx 또는 결산보고서 .xlsx)
        parsed_dir:  변환자료 폴더
        config_dir:  결산보고서 변환에 쓰는 config 디렉터리 (settlement_report.yaml)

    Returns:
        이상적 양식 records (read_ideal_ledger 결과)
    """
    # 지연 import (순환 의존 방지)
    from extractors import (
        parse_ledger, parse_cash_schedule, parse_settlement_report,
        write_ideal_ledger, read_ideal_ledger,
    )

    src = Path(source_path)
    out_dir = Path(parsed_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    cached = out_dir / f"{src.stem}__ideal_{file_hash(source_path)}.xlsx"

    if cached.exists():
        return read_ideal_ledger(str(cached))

    if "결산보고서" in src.name:
        recs = parse_settlement_report(source_path,
                                       str(Path(config_dir) / "settlement_report.yaml"))
    else:
        # 거래처원장(계정별 다중시트) 우선, 형식 미인식이면 잔액 스냅샷(현금성자산 등)으로 폴백.
        try:
            recs = parse_ledger(source_path)
        except ValueError:
            recs = parse_cash_schedule(source_path, resolver=_build_cash_resolver(parsed_dir))

    write_ideal_ledger(recs, str(cached))
    return read_ideal_ledger(str(cached))


def _build_cash_resolver(parsed_dir: str):
    """잔액 스냅샷 헤더가 C-1차로 안 잡힐 때만 쓰는 C-2차 마스킹 LLM 폴백 resolver.

    API 키가 외부(.env/환경변수)에 있을 때만 활성. 없으면 None(=오프라인, LLM 미호출).
    셀 값은 전부 마스킹되어 전송되며(실값 0 — 검증 완료), 결과는 타입검증 후 사용·학습캐시 저장.
    """
    try:
        from format_adapt import load_api_key, resolve_unmapped, LearnedHeaders
    except Exception:
        return None
    key = load_api_key()
    if not key:
        return None
    learned = LearnedHeaders(parsed_dir)

    def resolver(header_cells, sample_rows, missing_fields):
        import warnings
        verified, _ = resolve_unmapped(
            parser_key="cash", header_cells=header_cells, sample_rows=sample_rows,
            unmapped_fields=missing_fields,
            expected_types={"거래처명": {"text"}, "잔액": {"number"}, "계정과목": {"text"}},
            learned=learned, api_key=key, consent=lambda pl: True,
            announce=lambda m: warnings.warn("[포맷보정] " + m.splitlines()[0]),
        )
        return verified

    return resolver

"""파일 해시 기반 파싱 캐시 유틸."""

import datetime as _dt
import json
from hashlib import sha256
from pathlib import Path
from typing import Callable


def file_hash(path: str) -> str:
    """파일 내용의 SHA-256 앞 16자리. 내용이 바뀌면 해시도 바뀐다."""
    return sha256(Path(path).read_bytes()).hexdigest()[:16]


# ---------------------------------------------------------------------------
# 타입 보존 JSON 코덱
#
# 파서 결과에는 엑셀 날짜셀(datetime/date)이 그대로 담길 수 있다(예: 대출일·
# 최종만기일). 일반 JSON으로 직렬화하면 문자열로 변질되어, 생성기가 셀에 쓸 때
# 날짜 서식이 깨진다(양식 무결성 위반). 날짜류는 태그를 붙여 왕복 보존한다.
# 그 외 비직렬화 타입은 str 폴백(기존 동작과 동일).
# ---------------------------------------------------------------------------

def _json_default(o):
    if isinstance(o, (_dt.datetime, _dt.date, _dt.time)):
        return {"__kind__": type(o).__name__, "__iso__": o.isoformat()}
    return str(o)


def _decode_obj(d: dict):
    kind = d.get("__kind__")
    if kind and "__iso__" in d:
        iso = d["__iso__"]
        try:
            if kind == "datetime":
                return _dt.datetime.fromisoformat(iso)
            if kind == "date":
                return _dt.date.fromisoformat(iso)
            if kind == "time":
                return _dt.time.fromisoformat(iso)
        except ValueError:
            return iso  # 손상된 값은 문자열로(견고성)
    return d


def load_with_cache(path: str, parse_fn: Callable, *,
                    cache_dir: str | None = None, tag: str | None = None,
                    **kwargs) -> list | dict:
    """엑셀 파일을 파싱하고 결과를 JSON으로 캐시한다.

    파일 내용의 SHA-256 앞 16자리를 캐시 키로 사용한다.
    파일이 변경되면 해시가 달라지므로 자동으로 재파싱된다.

    Args:
        cache_dir: 캐시 저장 폴더. None이면 `_internal/cache`(개발 기본).
                   파이프라인은 배포 모델에 맞춰 `변환자료/`(parsed_dir)를 넘긴다.
        tag:       같은 파일을 여러 파서가 캐시할 때 키 충돌을 막는 식별자
                   (예: 취합엑셀을 bank/insurance/invest가 각각 파싱). None이면
                   `parse_fn.__name__`을 사용한다.
        **kwargs:  parse_fn에 그대로 전달.

    사용 예:
        rows = load_with_cache(path, parse_ledger)
        rows = load_with_cache(path, parse_bank, cache_dir=parsed, tag="bank")
    """
    cdir = Path(cache_dir) if cache_dir else Path("_internal/cache")
    cdir.mkdir(parents=True, exist_ok=True)

    label = tag or getattr(parse_fn, "__name__", "parse")
    cache_file = cdir / f"{Path(path).stem}__{label}_{file_hash(path)}.json"

    if cache_file.exists():
        try:
            return json.loads(cache_file.read_text(encoding="utf-8"),
                              object_hook=_decode_obj)
        except Exception:
            pass  # 캐시 손상 → 아래에서 재파싱

    result = parse_fn(path, **kwargs)
    try:
        cache_file.write_text(
            json.dumps(result, ensure_ascii=False, default=_json_default),
            encoding="utf-8",
        )
    except Exception:
        pass  # 캐시 기록 실패해도 결과는 반환(견고성 — 절대 전체실패 금지)
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

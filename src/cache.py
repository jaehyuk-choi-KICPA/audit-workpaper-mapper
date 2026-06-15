"""파일 해시 기반 파싱 캐시 유틸."""

import json
from hashlib import sha256
from pathlib import Path
from typing import Callable


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

    file_bytes = Path(path).read_bytes()
    file_hash = sha256(file_bytes).hexdigest()[:16]
    cache_file = cache_dir / f"{Path(path).stem}_{file_hash}.json"

    if cache_file.exists():
        return json.loads(cache_file.read_text(encoding="utf-8"))

    result = parse_fn(path, **kwargs)
    cache_file.write_text(
        json.dumps(result, ensure_ascii=False, default=str),
        encoding="utf-8",
    )
    return result

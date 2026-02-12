import os
from pathlib import Path
from typing import Optional


_env_loaded = False


def _load_env_if_present() -> None:
    global _env_loaded
    if _env_loaded:
        return
    _env_loaded = True

    env_path = Path(__file__).parent.parent / ".env"
    if not env_path.exists():
        return

    for line in env_path.read_text(encoding="utf-8").splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#") or "=" not in raw:
            continue
        key, value = raw.split("=", 1)
        key = key.strip()
        value = value.strip().strip("\"").strip("'")
        if key:
            os.environ.setdefault(key, value)


def getVarVal(name: str, default: Optional[str] = None) -> Optional[str]:
    _load_env_if_present()
    return os.environ.get(name, default)

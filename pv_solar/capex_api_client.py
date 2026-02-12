# -*- coding: utf-8 -*-
"""CAPEX API helpers split from capex_pptx."""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional


def _get_env_local_value(key: str) -> Optional[str]:
    env_path = Path(__file__).resolve().parents[1] / ".env.local"
    if not env_path.exists():
        return None
    try:
        for raw in env_path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            if k.strip() == key:
                return v.strip().strip("'\"")
    except Exception:
        return None
    return None


DEFAULT_API_URL = "https://devapi.hamonize.com/api/v1/chat/sync"
DEFAULT_API_DOCS = [doc.strip() for doc in os.getenv("HAMONIZE_API_DOCUMENTS", "test2.xlsx").split(",") if doc.strip()]
DEFAULT_API_USERNAME = os.getenv("HAMONIZE_API_USERNAME", "ryan")
DEFAULT_API_KEY = _get_env_local_value("NEXT_PUBLIC_RAG_STATUS_KEY")


def _extract_text_from_api_response(data: Any) -> Optional[str]:
    if data is None:
        return None
    if isinstance(data, str):
        return data.strip()
    if isinstance(data, dict):
        candidates = [
            data.get("answer"),
            data.get("content"),
            data.get("message"),
            data.get("text"),
            data.get("data"),
        ]
        for candidate in candidates:
            if isinstance(candidate, str) and candidate.strip():
                return candidate.strip()
        return str(data)
    return str(data)


def call_slide_api(prompt: str,
                   *,
                   mentioned_documents: Optional[List[str]] = None,
                   rag: bool = False,
                   web: bool = False,
                   username: str = DEFAULT_API_USERNAME,
                   timeout: int = 60,
                   retries: int = 2,
                   backoff: int = 2) -> Optional[str]:
    api_key = DEFAULT_API_KEY
    if not api_key:
        print("⚠️ HAMONIZE_API_KEY 환경변수가 없어 API 호출을 건너뜁니다.")
        return None

    try:
        import requests  # type: ignore
    except ImportError:
        print("⚠️ requests 라이브러리가 없어 API 호출을 건너뜁니다.")
        return None

    api_url = os.getenv("HAMONIZE_API_URL", DEFAULT_API_URL)
    docs = mentioned_documents if mentioned_documents is not None else DEFAULT_API_DOCS

    payload: Dict[str, Any] = {
        "prompt": prompt,
        "options": {
            "mentionedDocuments": docs,
            "rag": rag,
            "web": web,
            "username": username,
        },
    }

    headers = {
        "accept": "application/json",
        "X-API-Key": api_key,
        "Content-Type": "application/json",
    }

    for attempt in range(1, retries + 2):
        try:
            resp = requests.post(api_url, headers=headers, json=payload, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            text = _extract_text_from_api_response(data)
            if text:
                return text
            print(f"⚠️ API 응답에서 텍스트를 찾지 못했습니다. 응답: {data}")
            return None
        except Exception as exc:  # pylint: disable=broad-except
            if attempt > retries:
                print(f"⚠️ API 호출 실패 ({prompt[:40]}...): {exc}")
                return None
            sleep_for = backoff * attempt
            print(f"⚠️ API 호출 실패, {sleep_for}s 후 재시도 {attempt}/{retries} ({prompt[:40]}...): {exc}")
            try:
                import time
                time.sleep(sleep_for)
            except Exception:
                pass


def get_api_text_or_default(prompt: str,
                            default: str,
                            *,
                            documents: Optional[List[str]] = None,
                            rag: bool = False,
                            web: bool = False,
                            username: str = DEFAULT_API_USERNAME) -> str:
    text = call_slide_api(prompt, mentioned_documents=documents, rag=rag, web=web, username=username)
    if text:
        return text
    return default

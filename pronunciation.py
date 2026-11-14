"""音声合成向けの読み仮名・発音正規化モジュール.

固定読み辞書 (pronunciation_dict.json) の適用と、
OpenAI 互換 API (LMproxy / Ollama / LM Studio / Gemini / OpenAI 等) を用いた
バックグラウンドでの自動カタカナ・ひらがな変換を提供します。
"""

from __future__ import annotations

import json
import logging
import os
import re
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

logger = logging.getLogger("pptx_voice_pipeline")

DEFAULT_DICT_PATH = Path(__file__).resolve().parent / "pronunciation_dict.json"
DEFAULT_AI_BASE_URL = "http://127.0.0.1:4000/v1"
DEFAULT_AI_MODEL = "gpt-4o-mini"

_SYSTEM_PROMPT = (
    "あなたは日本語音声合成（VOICEVOX / TTS）用の読み上げ原稿正規化アシスタントです。\n"
    "入力テキストの意味・情報量・語順を一切変えず、要約や補足もしないでください。\n"
    "日本語の漢字・ひらがな・句読点はそのまま維持し、残っている英単語・英字略語・"
    "記号・バージョン表記（例: v1.2）だけを、日本語の音声合成エンジンが自然に発音できる"
    "カタカナまたはひらがな表記に書き換えてください。\n"
    "出力は変換後のテキスト本文のみを返し、解説や引用符・Markdown装飾は付けないでください。"
)


def load_pronunciation_dict(dict_path: Path | None = None) -> dict[str, str]:
    """読み辞書 JSON を読み込む. ファイルが存在しない場合は空辞書を返す."""
    target = dict_path or DEFAULT_DICT_PATH
    if not target.exists():
        return {}
    raw = json.loads(target.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise TypeError(f"読み辞書の形式が不正です (dict が必要): {target}")
    return {
        str(k).strip(): str(v).strip()
        for k, v in raw.items()
        if str(k).strip() and str(v).strip()
    }


def apply_pronunciation_dict(text: str, mapping: dict[str, str]) -> str:
    """英数字の単語境界を考慮して、固定読み辞書を 1 パスで置換する."""
    if not text or not mapping:
        return text

    lower_map = {k.lower(): v for k, v in mapping.items() if k}
    sorted_keys = sorted(mapping.keys(), key=len, reverse=True)
    escaped = "|".join(re.escape(k) for k in sorted_keys if k)
    if not escaped:
        return text

    pattern = re.compile(rf"(?<![A-Za-z0-9])(?:{escaped})(?![A-Za-z0-9])", re.IGNORECASE)
    return pattern.sub(lambda m: lower_map.get(m.group(0).lower(), m.group(0)), text)


def validate_openai_base_url(base_url: str) -> str:
    """OpenAI 互換 API のベース URL を検証して正規化する."""
    raw = (base_url or DEFAULT_AI_BASE_URL).strip().rstrip("/")
    parsed = urllib.parse.urlparse(raw)
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        raise ValueError(f"AI Base URL が不正です: {base_url!r}")
    return raw


def _needs_ai_normalization(text: str) -> bool:
    """英字（ASCII アルファベット）が残っている場合のみ AI 変換対象とする."""
    return bool(re.search(r"[A-Za-z]", text))


def _build_ai_headers(api_key: str = "") -> dict[str, str]:
    """LMproxy (OpenAI 互換) 呼び出し用の HTTP ヘッダを構築する."""
    resolved_key = (
        api_key.strip()
        or os.environ.get("LMPROXY_API_KEY", "").strip()
        or os.environ.get("OPENAI_API_KEY", "").strip()
    )
    headers = {"Content-Type": "application/json"}
    if resolved_key:
        headers["Authorization"] = f"Bearer {resolved_key}"
    return headers


def fetch_lmproxy_models(
    base_url: str = DEFAULT_AI_BASE_URL,
    *,
    api_key: str = "",
    timeout: float = 5.0,
) -> dict[str, object]:
    """LMproxy (/models) から利用可能なモデル ID 一覧を取得する."""
    clean_base = validate_openai_base_url(base_url)
    req = urllib.request.Request(
        f"{clean_base}/models",
        headers=_build_ai_headers(api_key),
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    raw_items = data.get("data") or []
    models = sorted(
        {
            str(item["id"]).strip()
            for item in raw_items
            if isinstance(item, dict) and str(item.get("id", "")).strip()
        }
    )
    return {"url": clean_base, "models": models}


def normalize_reading_with_ai(
    text: str,
    *,
    base_url: str = DEFAULT_AI_BASE_URL,
    model: str = DEFAULT_AI_MODEL,
    api_key: str = "",
    timeout: float = 20.0,
) -> str:
    """OpenAI 互換 /chat/completions API で英文・略語の読みをカタカナ/ひらがな化する."""
    if not text.strip():
        return text

    clean_base = validate_openai_base_url(base_url)
    endpoint = f"{clean_base}/chat/completions"
    headers = _build_ai_headers(api_key)

    payload = {
        "model": (model or DEFAULT_AI_MODEL).strip(),
        "temperature": 0.1,
        "messages": [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user", "content": text},
        ],
    }

    req = urllib.request.Request(
        endpoint,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers=headers,
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    choices = data.get("choices") or []
    if not choices:
        return text
    content = str(choices[0].get("message", {}).get("content", "")).strip()
    if not content:
        return text

    # 異常な要約・膨張をガード
    if len(content) < max(1, int(len(text) * 0.3)) or len(content) > max(50, len(text) * 4):
        logger.warning("AI 読み変換結果の長さが異常なため元テキストを維持します。")
        return text
    return content


def prepare_tts_text(
    text: str,
    *,
    dict_mapping: dict[str, str] | None = None,
    ai_enabled: bool = False,
    ai_base_url: str = DEFAULT_AI_BASE_URL,
    ai_model: str = DEFAULT_AI_MODEL,
    ai_api_key: str = "",
    cache: dict[str, str] | None = None,
) -> str:
    """固定読み辞書と（有効時は）AI 変換を適用して音声合成用テキストを返す."""
    result = apply_pronunciation_dict(text, dict_mapping) if dict_mapping else text

    if not ai_enabled or not _needs_ai_normalization(result):
        return result

    cache_key = f"{ai_base_url}|{ai_model}|{result}"
    if cache is not None and cache_key in cache:
        return cache[cache_key]

    try:
        normalized = normalize_reading_with_ai(
            result,
            base_url=ai_base_url,
            model=ai_model,
            api_key=ai_api_key,
        )
        if cache is not None:
            cache[cache_key] = normalized
        return normalized
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, ValueError, OSError) as exc:
        logger.warning("AI 読み変換をスキップして辞書適用済みテキストを使用します: %s", exc)
        return result

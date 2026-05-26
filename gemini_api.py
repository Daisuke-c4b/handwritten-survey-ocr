"""Gemini API 呼び出しの共通エラー処理."""
from __future__ import annotations

import json
from typing import Any

import requests

GEMINI_MODELS_DOC_URL = "https://ai.google.dev/gemini-api/docs/models?hl=ja"


class GeminiApiError(Exception):
    """Gemini API 呼び出し失敗。ユーザー向けメッセージを含む."""

    def __init__(
        self,
        message: str,
        *,
        status_code: int | None = None,
        api_detail: str = "",
        model_name: str = "",
    ) -> None:
        self.status_code = status_code
        self.api_detail = api_detail
        self.model_name = model_name
        super().__init__(message)


def extract_api_error_detail(response: requests.Response | None) -> str:
    """Gemini API の JSON エラー応答からメッセージを抽出."""
    if response is None:
        return ""
    try:
        data = response.json()
        if "error" in data:
            err = data["error"]
            if isinstance(err, dict):
                return (
                    err.get("message")
                    or err.get("status")
                    or json.dumps(err, ensure_ascii=False)
                )
            return str(err)
    except Exception:
        pass
    text = (response.text or "").strip()
    return text[:500] if text else ""


def raise_for_gemini_response(response: requests.Response, model_name: str) -> None:
    """HTTP エラー応答を GeminiApiError に変換して送出."""
    if response.ok:
        return
    detail = extract_api_error_detail(response)
    summary = f"Gemini API エラー (HTTP {response.status_code})"
    if detail:
        summary += f": {detail}"
    raise GeminiApiError(
        summary,
        status_code=response.status_code,
        api_detail=detail,
        model_name=model_name,
    )


def wrap_gemini_exception(exc: Exception, model_name: str) -> GeminiApiError:
    """任意の例外を GeminiApiError に正規化."""
    if isinstance(exc, GeminiApiError):
        return exc
    if isinstance(exc, requests.HTTPError) and exc.response is not None:
        detail = extract_api_error_detail(exc.response)
        status = exc.response.status_code
        summary = f"Gemini API エラー (HTTP {status})"
        if detail:
            summary += f": {detail}"
        return GeminiApiError(
            summary,
            status_code=status,
            api_detail=detail,
            model_name=model_name,
        )
    return GeminiApiError(
        f"Gemini API 呼び出しに失敗しました: {exc}",
        model_name=model_name,
    )


def parse_generate_content_response(data: dict[str, Any]) -> str:
    """generateContent 応答 JSON からテキストを取り出す."""
    if "error" in data:
        err = data["error"]
        if isinstance(err, dict):
            msg = err.get("message") or err.get("status") or json.dumps(err, ensure_ascii=False)
        else:
            msg = str(err)
        raise GeminiApiError(f"Gemini API エラー: {msg}", api_detail=msg)
    return (
        data.get("candidates", [{}])[0]
        .get("content", {})
        .get("parts", [{}])[0]
        .get("text", "")
    ).strip()

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


def build_generate_content_payload(parts: list[dict[str, Any]]) -> dict[str, Any]:
    """文字起こし向けの generateContent リクエスト.

    Gemini 3 系は複雑な画像で思考が長引く。抽出タスクでは thinkingLevel を
    minimal に固定し、応答が返るまでの待ち時間を抑える。
    """
    return {
        "contents": [{"parts": parts}],
        "generationConfig": {
            "thinkingConfig": {"thinkingLevel": "minimal"},
        },
    }


def extract_candidate_text(data: dict[str, Any]) -> str:
    """generateContent 応答から本文を取り出す.

    Gemini 3 は parts の先頭に思考パートを置くことがある。先頭だけ読むと
    本文が空になり、呼び出し側が失敗とみなして再試行を繰り返す。
    """
    if "error" in data:
        err = data["error"]
        if isinstance(err, dict):
            msg = err.get("message") or err.get("status") or json.dumps(err, ensure_ascii=False)
        else:
            msg = str(err)
        raise GeminiApiError(f"Gemini API エラー: {msg}", api_detail=msg)

    candidates = data.get("candidates") or []
    if not candidates:
        return ""
    parts = (candidates[0].get("content") or {}).get("parts") or []
    answers: list[str] = []
    thoughts: list[str] = []
    for part in parts:
        if not isinstance(part, dict):
            continue
        text = (part.get("text") or "").strip()
        if not text:
            continue
        if part.get("thought"):
            thoughts.append(text)
        else:
            answers.append(text)
    if answers:
        return "\n".join(answers).strip()
    return "\n".join(thoughts).strip()


def parse_generate_content_response(data: dict[str, Any]) -> str:
    """generateContent 応答 JSON からテキストを取り出す."""
    return extract_candidate_text(data)

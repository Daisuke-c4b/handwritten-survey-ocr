"""API utilization tracker (calls, characters, duration, estimated cost).

OCR / TextEditor などの Gemini 呼び出し全てから利用される。
Streamlit session_state とは独立に動作する純粋モジュールで、
内部状態はモジュールレベルの dict に保持される（プロセス単位）。

Streamlit が再実行されてもプロセスは生き続けるため、累計値は持続する。
明示的に reset() でクリアできる。

注意:
- トークン数は Gemini が返す usage を必ずしも取れないため、
  文字数ベースで粗い見積もりを行う（日本語約 2 文字 ≒ 1 トークン）。
- 価格も公開価格から推定値を採用する。
  実際の請求金額とは異なる可能性があるため、目安値として表示する。
"""
from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field, asdict
from typing import Optional

# Gemini Flash-Lite 系の参考価格 (USD / 1M tokens)
# 実際の価格は変動するため、設定値として上書き可能とする。
_DEFAULT_INPUT_PRICE_PER_1M_TOKENS_USD = 0.075
_DEFAULT_OUTPUT_PRICE_PER_1M_TOKENS_USD = 0.30

# 日本語前提の粗い文字→トークン換算係数（1 トークン ≈ 2 文字）
_CHARS_PER_TOKEN_ESTIMATE = 2.0


@dataclass
class ApiCallRecord:
    """1 回の API 呼び出し記録."""

    timestamp: float
    purpose: str  # 例: "OCR(page)", "matome", "校正チェック", "分析"
    model: str
    duration_sec: float
    input_chars: int
    output_chars: int
    status: str  # "ok" | "error" | "retry"
    error: Optional[str] = None
    image_bytes: int = 0  # 画像が含まれる場合の追加バイト数


@dataclass
class _TrackerState:
    """累計トラッキング状態."""

    records: list[ApiCallRecord] = field(default_factory=list)
    input_price_per_1m: float = _DEFAULT_INPUT_PRICE_PER_1M_TOKENS_USD
    output_price_per_1m: float = _DEFAULT_OUTPUT_PRICE_PER_1M_TOKENS_USD
    chars_per_token: float = _CHARS_PER_TOKEN_ESTIMATE


_STATE = _TrackerState()
_LOCK = threading.Lock()


def reset() -> None:
    """累計記録をクリア."""
    with _LOCK:
        _STATE.records.clear()


def record(
    purpose: str,
    model: str,
    duration_sec: float,
    input_chars: int,
    output_chars: int,
    status: str = "ok",
    error: Optional[str] = None,
    image_bytes: int = 0,
) -> None:
    """API 呼び出し結果を記録."""
    with _LOCK:
        _STATE.records.append(
            ApiCallRecord(
                timestamp=time.time(),
                purpose=purpose,
                model=model,
                duration_sec=duration_sec,
                input_chars=input_chars,
                output_chars=output_chars,
                status=status,
                error=error,
                image_bytes=image_bytes,
            )
        )


def configure_pricing(
    input_price_per_1m: Optional[float] = None,
    output_price_per_1m: Optional[float] = None,
    chars_per_token: Optional[float] = None,
) -> None:
    """価格・換算係数の上書き設定."""
    with _LOCK:
        if input_price_per_1m is not None:
            _STATE.input_price_per_1m = input_price_per_1m
        if output_price_per_1m is not None:
            _STATE.output_price_per_1m = output_price_per_1m
        if chars_per_token is not None and chars_per_token > 0:
            _STATE.chars_per_token = chars_per_token


def _estimate_tokens(chars: int) -> float:
    if _STATE.chars_per_token <= 0:
        return 0.0
    return chars / _STATE.chars_per_token


def _estimate_cost_usd(input_chars: int, output_chars: int) -> float:
    in_tokens = _estimate_tokens(input_chars)
    out_tokens = _estimate_tokens(output_chars)
    return (
        in_tokens * _STATE.input_price_per_1m / 1_000_000
        + out_tokens * _STATE.output_price_per_1m / 1_000_000
    )


def summary() -> dict:
    """累計サマリーを dict で返す."""
    with _LOCK:
        records = list(_STATE.records)
        in_price = _STATE.input_price_per_1m
        out_price = _STATE.output_price_per_1m
        chars_per_token = _STATE.chars_per_token

    total_calls = len(records)
    total_ok = sum(1 for r in records if r.status == "ok")
    total_err = sum(1 for r in records if r.status == "error")
    total_in_chars = sum(r.input_chars for r in records)
    total_out_chars = sum(r.output_chars for r in records)
    total_duration = sum(r.duration_sec for r in records)
    total_image_bytes = sum(r.image_bytes for r in records)

    return {
        "total_calls": total_calls,
        "successful_calls": total_ok,
        "failed_calls": total_err,
        "total_input_chars": total_in_chars,
        "total_output_chars": total_out_chars,
        "total_duration_sec": round(total_duration, 2),
        "total_image_bytes": total_image_bytes,
        "estimated_input_tokens": int(_estimate_tokens(total_in_chars)),
        "estimated_output_tokens": int(_estimate_tokens(total_out_chars)),
        "estimated_cost_usd": round(_estimate_cost_usd(total_in_chars, total_out_chars), 6),
        "input_price_per_1m_usd": in_price,
        "output_price_per_1m_usd": out_price,
        "chars_per_token": chars_per_token,
        "average_duration_sec": round(
            total_duration / total_calls, 2
        ) if total_calls else 0.0,
    }


def per_purpose_summary() -> list[dict]:
    """目的（用途）別の集計を返す."""
    with _LOCK:
        records = list(_STATE.records)

    buckets: dict[str, dict] = {}
    for r in records:
        b = buckets.setdefault(
            r.purpose,
            {
                "purpose": r.purpose,
                "calls": 0,
                "duration_sec": 0.0,
                "input_chars": 0,
                "output_chars": 0,
                "errors": 0,
            },
        )
        b["calls"] += 1
        b["duration_sec"] += r.duration_sec
        b["input_chars"] += r.input_chars
        b["output_chars"] += r.output_chars
        if r.status == "error":
            b["errors"] += 1

    for b in buckets.values():
        b["duration_sec"] = round(b["duration_sec"], 2)
        b["estimated_cost_usd"] = round(
            _estimate_cost_usd(b["input_chars"], b["output_chars"]), 6
        )
    return sorted(buckets.values(), key=lambda x: -x["calls"])


def recent_records(limit: int = 50) -> list[dict]:
    """直近の呼び出し履歴を dict 化して返す."""
    with _LOCK:
        records = list(_STATE.records)[-limit:]
    return [asdict(r) for r in records]


class TimedCall:
    """`with TimedCall(...) as t: ... ; t.set_output(text)` で使うコンテキスト."""

    def __init__(
        self,
        purpose: str,
        model: str,
        input_chars: int = 0,
        image_bytes: int = 0,
    ) -> None:
        self.purpose = purpose
        self.model = model
        self.input_chars = input_chars
        self.image_bytes = image_bytes
        self.output_chars = 0
        self.status = "ok"
        self.error: Optional[str] = None
        self._start = 0.0

    def __enter__(self) -> "TimedCall":
        self._start = time.time()
        return self

    def set_output(self, text: str) -> None:
        self.output_chars = len(text or "")

    def set_input(self, text: str) -> None:
        self.input_chars = len(text or "")

    def add_image_bytes(self, n: int) -> None:
        self.image_bytes += max(0, int(n))

    def mark_error(self, e: BaseException) -> None:
        self.status = "error"
        self.error = repr(e)

    def __exit__(self, exc_type, exc, tb) -> bool:
        duration = time.time() - self._start
        if exc is not None and self.status != "error":
            self.status = "error"
            self.error = repr(exc)
        record(
            purpose=self.purpose,
            model=self.model,
            duration_sec=duration,
            input_chars=self.input_chars,
            output_chars=self.output_chars,
            status=self.status,
            error=self.error,
            image_bytes=self.image_bytes,
        )
        return False

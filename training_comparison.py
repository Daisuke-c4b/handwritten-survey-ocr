"""前回研修と今回研修のアンケート Word 比較."""

from __future__ import annotations

_TRAINING_COMPARE_PROMPT = """\
あなたは研修アンケート分析の専門家です。
前回研修と今回研修の文字起こし結果（Word から抽出したテキスト）を比較し、
研修改善のための実用的なレポートを Markdown 形式で作成してください。

【前回研修のアンケート】
{previous}

【今回研修のアンケート】
{current}

【レポート構成（厳守）】
## 概要
- 前回と今回の回答規模・設問構成の違い（あれば）

## センチメント・印象の変化
- ポジティブ / ネガティブ / 中立の傾向変化
- 顕著に改善した点・悪化した点

## トピック・キーワードの変化
- 前回に多かったトピック vs 今回に多かったトピック
- 新たに現れた意見・消えた意見

## 設問別の比較（可能な範囲で）
- 共通する設問について、回答傾向の違いを具体的に

## 研修改善への示唆
- 次回に活かすべき点（箇条書き 3〜5 件）
- 注意すべきリスクや継続課題

【注意】
- 推測は最小限にし、テキストに根拠のある記述を優先してください。
- 数値は概算で構いませんが、根拠なく断定しないでください。
- 出力は Markdown のみ（前置き・締めの挨拶は不要）。
"""


def build_training_comparison_prompt(previous: str, current: str) -> str:
    return _TRAINING_COMPARE_PROMPT.format(
        previous=previous.strip() or "（空）",
        current=current.strip() or "（空）",
    )


def generate_training_comparison(
    previous_text: str,
    current_text: str,
    caller,
) -> str:
    """Gemini で前回 vs 今回の比較レポートを生成."""
    prompt = build_training_comparison_prompt(previous_text, current_text)
    return caller.call_with_purpose(prompt, purpose="研修比較")

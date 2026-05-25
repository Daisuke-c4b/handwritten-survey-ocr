"""アンケート分析: 回答者単位ビュー / 定量サマリー / テンプレート自動生成.

Streamlit から切り離した純粋ロジック。
テスト容易性のため、外部 API 呼び出しは引数として渡される caller (TextEditor 等) に委譲する。
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Callable, Optional


# ---------------------------------------------------------------------------
# 質問・回答パーサ
# ---------------------------------------------------------------------------

_Q_HEADER_RE = re.compile(r"^(?:質問|Q|Ｑ)\s*(\d+)\s*[:：]\s*(.+)$")
_ANSWER_LINE_RE = re.compile(r"^[・\-\*]\s*(?:\[(P\.\d+|回答者\d+|F\..+?-P\.\d+)\]\s*)?(.+)$")


@dataclass
class ParsedAnswer:
    respondent_id: str  # 例: "P.1", "回答者1"。識別子がなければ "unknown"
    text: str


@dataclass
class ParsedQuestion:
    q_num: int
    q_text: str
    answers: list[ParsedAnswer] = field(default_factory=list)


def parse_consolidated_text(consolidated: str) -> list[ParsedQuestion]:
    """`質問N: ...\n・[P.N] 回答` 形式のテキストを構造化."""
    if not consolidated:
        return []

    questions: list[ParsedQuestion] = []
    current: Optional[ParsedQuestion] = None

    for raw in consolidated.splitlines():
        line = raw.strip()
        if not line:
            continue

        m = _Q_HEADER_RE.match(line)
        if m:
            if current is not None:
                questions.append(current)
            current = ParsedQuestion(
                q_num=int(m.group(1)),
                q_text=m.group(2).strip(),
            )
            continue

        if current is None:
            # 質問ヘッダの前にある行は無視
            continue

        am = _ANSWER_LINE_RE.match(line)
        if am:
            respondent_id = am.group(1) or "unknown"
            text = am.group(2).strip()
            if text and text != "（無回答）":
                current.answers.append(ParsedAnswer(respondent_id=respondent_id, text=text))
        else:
            # 箇条書きでない通常行は最後の回答に連結（複数行回答対応）
            if current.answers:
                current.answers[-1].text += " " + line

    if current is not None:
        questions.append(current)

    return questions


# ---------------------------------------------------------------------------
# 回答者単位ビュー
# ---------------------------------------------------------------------------

@dataclass
class RespondentBlock:
    respondent_id: str
    answers: list[tuple[int, str, str]]  # (q_num, q_text, answer_text)


def to_respondent_view(questions: list[ParsedQuestion]) -> list[RespondentBlock]:
    """質問単位の集約から回答者単位ビューへ再構成する."""
    grouped: dict[str, list[tuple[int, str, str]]] = {}
    for q in questions:
        for ans in q.answers:
            grouped.setdefault(ans.respondent_id, []).append((q.q_num, q.q_text, ans.text))

    blocks: list[RespondentBlock] = []
    # 回答者IDの並び順: P.N → 数値昇順、その他は辞書順
    def sort_key(rid: str) -> tuple[int, int, str]:
        m = re.match(r"P\.(\d+)", rid)
        if m:
            return (0, int(m.group(1)), rid)
        m2 = re.match(r"回答者(\d+)", rid)
        if m2:
            return (1, int(m2.group(1)), rid)
        return (2, 0, rid)

    for rid in sorted(grouped.keys(), key=sort_key):
        answers = sorted(grouped[rid], key=lambda x: x[0])
        blocks.append(RespondentBlock(respondent_id=rid, answers=answers))
    return blocks


def respondent_view_as_markdown(blocks: list[RespondentBlock]) -> str:
    """回答者単位ビューを Markdown で文字列化."""
    if not blocks:
        return "（回答者識別子が検出できませんでした。先に「質問ごとにまとめる」加工を実行してください）"
    lines: list[str] = []
    for b in blocks:
        lines.append(f"## [{b.respondent_id}] の回答")
        for q_num, q_text, ans in b.answers:
            lines.append(f"- **Q{q_num}: {q_text}**")
            lines.append(f"  - {ans}")
        lines.append("")
    return "\n".join(lines).rstrip()


# ---------------------------------------------------------------------------
# 定量サマリー
# ---------------------------------------------------------------------------

_QUANT_SUMMARY_PROMPT = """\
あなたはアンケート定量分析の専門家です。以下の集約済みアンケート回答テキストを分析し、
回答単位のセンチメント分類とトピック抽出を行い、JSON 形式で返してください。

【入力テキスト】
{text}

【出力形式（厳守: JSON のみ。Markdown コードフェンスは付けない）】
{{
  "answer_count": <分析した回答の総数: int>,
  "sentiment": {{
    "positive": <件数: int>,
    "negative": <件数: int>,
    "neutral": <件数: int>
  }},
  "topics": [
    {{"label": "<トピック名（短く）>", "count": <該当件数: int>, "examples": ["<代表表現1>", "<代表表現2>"]}},
    ...
  ],
  "keywords": [
    {{"word": "<頻出キーワード>", "count": <出現回数: int>}},
    ...
  ],
  "highlights": {{
    "positive": ["<代表的なポジティブ回答1>", ...],
    "negative": ["<代表的なネガティブ回答1>", ...]
  }}
}}

【注意】
- センチメントは「ポジティブ / ネガティブ / 中立」の3分類。
- トピックは類似回答をまとめた意味的なクラスタ名。最大 8 件、件数が多い順。
- keywords は研修・施策の改善に役立つ名詞句を中心に最大 10 件。
- highlights は各カテゴリ最大 3 件まで、原文をそのまま引用。
- 必ず JSON 単体で返してください。前後の説明文や Markdown 記号は禁止です。
"""


def build_quant_summary_prompt(consolidated_text: str) -> str:
    return _QUANT_SUMMARY_PROMPT.format(text=consolidated_text)


def parse_quant_summary_response(raw: str) -> dict:
    """Gemini からの応答を JSON parse。前後ノイズは堅牢に剥がす."""
    if not raw or not raw.strip():
        return {}
    text = raw.strip()
    # マークダウンコードフェンス除去
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```\s*$", "", text)
    # 最初の { から最後の } までを抜き取る（前置きが含まれた場合の保険）
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start : end + 1]
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"_raw": raw}


# ---------------------------------------------------------------------------
# 設問テンプレート自動生成
# ---------------------------------------------------------------------------

_TEMPLATE_GEN_PROMPT = """\
あなたはアンケート用紙の解析専門家です。以下は同一のアンケート用紙の文字起こし結果（複数ページ・複数回答者を含むことがあります）です。
この用紙に印字されている**質問文・タイトル・案内文・選択肢ラベル**など、
**「文字起こし対象から除外すべき固定テキスト」**を抽出して、1 行に 1 つずつ列挙してください。

【入力テキスト】
{text}

【出力形式（厳守）】
- 除外対象とする固定テキストを、1 行に 1 つずつ、装飾なしのプレーンテキストで列挙してください。
- 質問番号（例: 「1.」「Q1:」など）が付いている場合は番号も含めてそのまま出力してください。
- 手書きの回答内容や個別の意見は**絶対に含めないでください**。
- 同じ意味の重複行は 1 つに統合してください。
- 余計な見出し・注釈・前置き・締めの文章は付けないでください。テキストのみを行ごとに出力してください。
"""


def build_template_generation_prompt(text: str) -> str:
    return _TEMPLATE_GEN_PROMPT.format(text=text)


def parse_template_generation_response(raw: str) -> list[str]:
    """応答テキストから1行1項目のテンプレート候補を抽出."""
    if not raw:
        return []
    items: list[str] = []
    for line in raw.splitlines():
        s = line.strip()
        if not s:
            continue
        # 箇条書き記号を剥がす
        s = re.sub(r"^[-*・]\s*", "", s)
        # 番号付き箇条書き先頭の "1) " 等を剥がす（質問番号「1.」は保持したい場合があるため厳格にしない）
        if re.match(r"^\d+[\)）]\s+", s):
            s = re.sub(r"^\d+[\)）]\s+", "", s)
        if s and s not in items:
            items.append(s)
    return items


# ---------------------------------------------------------------------------
# 高レベル API（TextEditor などを caller として注入）
# ---------------------------------------------------------------------------

GeminiCaller = Callable[[str], str]


def generate_quant_summary(text: str, caller: GeminiCaller) -> dict:
    """テキストから定量サマリーを生成（Gemini 呼び出しは caller に委譲）."""
    if not text or not text.strip():
        return {}
    prompt = build_quant_summary_prompt(text)
    raw = caller(prompt)
    return parse_quant_summary_response(raw)


def generate_exclude_template(text: str, caller: GeminiCaller) -> list[str]:
    """テキストから除外テンプレート候補を生成."""
    if not text or not text.strip():
        return []
    prompt = build_template_generation_prompt(text)
    raw = caller(prompt)
    return parse_template_generation_response(raw)

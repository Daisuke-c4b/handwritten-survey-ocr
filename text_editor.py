import time
import requests
from ocr_processor import _get_api_key, MODEL_NAME

EDITING_MODES: dict[str, dict] = {
    "matome": {
        "label": "質問ごとにまとめる",
        "short_desc": "質問別に全回答を集約し、回答者識別子を付与",
        "description": (
            "複数ページ・複数回答の文字起こし結果から、"
            "印字された質問ごとに回答を集約します。"
            "同じ質問への回答がバラバラに並んでいる場合に、"
            "質問単位でまとめて見やすく整理します。"
            "各回答の先頭には [P.N] や [回答者N] などの識別子が付与され、"
            "質問1・質問2いずれかしか回答していない回答者がいても"
            "同一の回答者を追跡できます。"
        ),
        "prompt": (
            "あなたは厳密な分類を行うアシスタントです。\n"
            "以下の文字起こしテキストを、印字された質問ごとに回答を完全に分けて集約してください。\n\n"
            "【厳守事項】\n"
            "1. 各回答は「元のテキストでその回答が書かれていた質問」の下にのみ配置してください。\n"
            "   質問の境界を絶対に超えてはいけません。Q2 の回答を Q1 に混ぜることを禁止します。\n"
            "2. 同じ質問に対する複数の回答は、それぞれ独立した行として列挙してください。\n"
            "   勝手に統合・要約・言い換えしてはいけません（内容・意味は厳密に保持）。\n"
            "3. 回答者識別子の保持・付与:\n"
            "   - 入力に `[P.<数字>]` や `[ページ<数字>]` や `--- ページ<数字> ---` が含まれる場合、\n"
            "     各回答の先頭に必ず `[P.<数字>]` の形式で識別子を付けてください。\n"
            "   - 入力に識別子がない場合は、各回答ブロックに `[回答者1]`, `[回答者2]` …と\n"
            "     出現順に通し番号を付与してください。\n"
            "   - 同一回答者（同じ識別子）の回答は、複数の質問にまたがって配置されても識別子を一貫させてください。\n"
            "4. 質問1・質問2いずれかしか回答していない回答者がいても、その回答は該当する質問の下にだけ配置し、\n"
            "   無回答の質問の下に勝手に挿入してはいけません。\n"
            "5. 回答がない質問の下には「・（無回答）」と1行のみ記載してください。\n\n"
            "【出力形式（厳守）】\n"
            "質問1: <印字された質問文をそのまま転記>\n"
            "・[P.1] 回答内容\n"
            "・[P.2] 回答内容\n"
            "\n"
            "質問2: <印字された質問文をそのまま転記>\n"
            "・[P.1] 回答内容\n"
            "・[P.3] 回答内容\n"
            "\n"
            "【出力に関する追加ルール】\n"
            "- 余計な見出し・前置き・注釈・締めの文章は一切付けないでください。\n"
            "- 出力は上記の構造のみとしてください。\n"
            "- 質問文は元テキストに印字されている文言を改変せず転記してください。\n\n"
            "元のテキスト：\n{text}"
        ),
    },
    "seimon": {
        "label": "整文",
        "short_desc": "です・ます調に整えて文体を統一",
        "description": (
            "句読点の補完・文体の統一を行い、文章全体をです・ます調に整えます。"
            "話し言葉を丁寧な書き言葉に変えますが、内容・意味は変えません。"
            "アンケート回答を正式な記録文書として仕上げたいときに適しています。"
        ),
        "prompt": (
            "以下の文字起こしテキストを整文してください。"
            "句読点を適切に補完し、文章全体を「です・ます調」に統一してください。"
            "内容・意味は変えずに、整文後のテキストのみ出力してください。\n\n"
            "元のテキスト：\n{text}"
        ),
    },
    "youyaku": {
        "label": "要約",
        "short_desc": "要点を簡潔にまとめる",
        "description": (
            "文章の要点を抽出し、簡潔にまとめます。"
            "重要なポイントが失われないよう配慮しながら圧縮します。"
            "長い回答を短くまとめて報告書などに使いやすくしたいときに適しています。"
        ),
        "prompt": (
            "以下の文字起こしテキストを要約してください。"
            "重要なポイントを漏らさず、簡潔にまとめてください。"
            "要約後のテキストのみ出力してください。\n\n"
            "元のテキスト：\n{text}"
        ),
    },
    "custom": {
        "label": "カスタム",
        "short_desc": "独自のプロンプトで自由編集",
        "description": (
            "自分でプロンプトを入力して、自由にテキストを編集できます。"
            "「敬語に変換して」「箇条書きにして」「英語に翻訳して」など、"
            "用途に合わせた加工が可能です。"
        ),
        "prompt": "",
    },
}

_TEXT_QUALITY_CHECK_PROMPT = """
あなたは日本語校正の専門家です。以下のアンケートテキストを慎重に読み、
「誤字脱字」「違和感のある文体」「不自然な日本語」を検出してレポートしてください。

【検出カテゴリ】
1. 誤字脱字（明らかな誤字・脱字・送り仮名の誤り・漢字の誤変換）
2. 違和感のある文体（敬体「です・ます」と常体「だ・である」の混在、口語と書き言葉の混在など）
3. 不自然な日本語（文法的な誤り、係り受けの不整合、意味の通じにくい表現）

【出力形式（厳守: Markdown）】
問題が検出された場合は、必ず以下の Markdown 形式で列挙してください。

## 検出件数: N 件

### 1. [カテゴリ] 該当箇所
- **該当箇所**: 「（問題のある表現を原文から正確に抜き出す）」
- **指摘内容**: （何が問題か簡潔に1〜2文で）
- **修正案**: 「（自然で正確な表現に直したもの）」
- **前後の文脈**: 「（該当箇所を含む1文ほど）」

### 2. [カテゴリ] ...
（以下同様）

問題が一つも見つからない場合は、以下のみを出力してください:

## 問題は検出されませんでした

【注意】
- 「カテゴリ」は「誤字脱字」「違和感」「不自然」のいずれか1つを必ず角括弧内に書いてください。
- 原文の書き換え結果（全文）は出力しないでください。指摘と修正案のみで十分です。
- 推測ではなく、明確に問題と判断できる箇所のみを指摘してください。
- 同じ問題が複数回出現する場合は1件にまとめ、「他にも同種の箇所あり」と注記してください。
- 余計な前置きや締めの文章は付けず、上記の Markdown 構造のみを出力してください。

【対象テキスト】
{text}
"""


_SURVEY_ANALYSIS_PROMPT = """
あなたはアンケート分析の専門家です。以下のアンケート回答をもとに、詳細な分析レポートを作成してください。

【アンケート回答】
{combined}

---

以下の観点で分析し、それぞれ具体的に記述してください：

## 1. 全体的な傾向・概要
回答全体を俯瞰し、共通するテーマや傾向をまとめてください。

## 2. 主な意見・感想のカテゴリ分け
回答を内容でグループ化し、各カテゴリの代表的な意見を挙げてください。

## 3. 課題・改善が必要な点
ネガティブな意見や要望から、対応が必要な課題を整理してください。

## 4. アクションプラン（具体的な対応策）
課題に対して、実行可能な具体的アクションを優先度付きで提案してください。

## 5. 今後の研修・施策への改善提案
アンケート結果を踏まえ、次回以降の研修や取り組みへの改善提案をしてください。

分析は具体的かつ実用的な内容にし、現場で活用できる提案を心がけてください。
"""


class TextEditor:
    def __init__(self):
        self.api_key = _get_api_key()
        self.endpoint = (
            f"https://generativelanguage.googleapis.com/v1beta/models/"
            f"{MODEL_NAME}:generateContent"
        )

    def apply_editing(self, text: str, mode: str, custom_prompt: str = "") -> str:
        """Apply the selected editing mode to the given text."""
        if mode == "custom":
            if not custom_prompt.strip():
                return text
            prompt = f"{custom_prompt.strip()}\n\n対象テキスト：\n{text}"
        elif mode in EDITING_MODES:
            prompt = EDITING_MODES[mode]["prompt"].format(text=text)
        else:
            return text

        return self._call_gemini(prompt)

    def analyze_survey(self, texts: list[str], filenames: list[str]) -> str:
        """Analyze survey responses and generate an action plan."""
        sections = []
        for fn, t in zip(filenames, texts):
            sections.append(f"【{fn}】\n{t}")
        combined = "\n\n---\n\n".join(sections)
        prompt = _SURVEY_ANALYSIS_PROMPT.format(combined=combined)
        return self._call_gemini(prompt)

    def check_text_quality(self, text: str) -> str:
        """Check pasted survey text for typos, awkward style, and unnatural Japanese.

        Returns a Markdown-formatted report listing detected issues with
        category, location, suggested fix and surrounding context.
        """
        if not text or not text.strip():
            return "## 問題は検出されませんでした\n\n（入力テキストが空です）"
        prompt = _TEXT_QUALITY_CHECK_PROMPT.format(text=text)
        return self._call_gemini(prompt)

    def _call_gemini(self, prompt: str, max_retries: int = 3) -> str:
        """Call Gemini API with exponential backoff retry on 5xx errors."""
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        last_err: Exception = RuntimeError("API呼び出しに失敗しました")

        for attempt in range(max_retries):
            try:
                res = requests.post(
                    f"{self.endpoint}?key={self.api_key}",
                    headers={"Content-Type": "application/json"},
                    json=payload,
                    timeout=120,
                )
                res.raise_for_status()
                return (
                    res.json()
                    .get("candidates", [{}])[0]
                    .get("content", {})
                    .get("parts", [{}])[0]
                    .get("text", "")
                ).strip()
            except requests.HTTPError as e:
                last_err = e
                status = e.response.status_code if e.response is not None else 0
                # Retry on 5xx (server-side) errors
                if status >= 500 and attempt < max_retries - 1:
                    wait = 2 ** attempt  # 1s → 2s → 4s
                    time.sleep(wait)
                    continue
                raise
            except Exception as e:
                last_err = e
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                raise

        raise last_err

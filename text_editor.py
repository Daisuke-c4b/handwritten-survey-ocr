import time
import requests
from ocr_processor import _get_api_key, MODEL_NAME

EDITING_MODES: dict[str, dict] = {
    "matome": {
        "label": "質問ごとにまとめる",
        "short_desc": "質問別に全回答を集約して整理",
        "description": (
            "複数ページ・複数回答の文字起こし結果から、"
            "印字された質問ごとに回答を集約します。"
            "同じ質問への回答がバラバラに並んでいる場合に、"
            "質問単位でまとめて見やすく整理します。"
        ),
        "prompt": (
            "以下の文字起こしテキストを、印字された質問ごとに回答をまとめてください。\n"
            "同じ質問への複数の回答を集約し、以下の形式で出力してください。\n\n"
            "出力形式：\n"
            "質問1（印字された質問文をそのまま記載）\n"
            "回答A\n"
            "回答B\n"
            "回答C\n"
            "\n"
            "質問2（印字された質問文をそのまま記載）\n"
            "回答A\n"
            "回答B\n"
            "\n"
            "注意：\n"
            "- 質問文はそのまま転記してください\n"
            "- 各回答は1行ずつ改行で区切ってください\n"
            "- 回答がない場合は（無回答）と記載してください\n"
            "- 回答の意味・内容は変えないでください\n\n"
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

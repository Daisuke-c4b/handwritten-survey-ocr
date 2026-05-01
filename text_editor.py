import requests
from ocr_processor import _get_api_key, MODEL_NAME

EDITING_MODES: dict[str, dict] = {
    "kebatori": {
        "label": "ケバ取り",
        "short_desc": "フィラー語・言い直しを削除",
        "description": (
            "「えー」「あー」「えっと」「まあ」「なんか」などのフィラー語（不要な言葉）や、"
            "言い直し（「○○、じゃなくて△△」など）、重複表現を自動で削除します。"
            "内容・意味は変えず、読みやすく整理します。"
        ),
        "prompt": (
            "以下の文字起こしテキストから、「えー」「あー」「えっと」「まあ」「なんか」などの"
            "フィラー語、言い直し、重複表現を除去してください。"
            "内容・意味は変えずに、整理したテキストのみ出力してください。\n\n"
            "元のテキスト：\n{text}"
        ),
    },
    "seimon": {
        "label": "整文",
        "short_desc": "句読点補完・文体を整理",
        "description": (
            "句読点の補完・文体の統一・文章の流れを整えます。"
            "話し言葉を書き言葉に整えますが、内容・意味は変えません。"
            "アンケート回答を正式な記録文書として仕上げたいときに適しています。"
        ),
        "prompt": (
            "以下の文字起こしテキストを整文してください。"
            "句読点を適切に補完し、文章の流れを整えて自然な日本語の書き言葉にしてください。"
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

    def _call_gemini(self, prompt: str) -> str:
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
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

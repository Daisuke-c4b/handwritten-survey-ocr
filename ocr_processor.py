import streamlit as st
import requests
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageFilter
import io
import os
import re
import base64
import numpy as np

from cost_tracker import TimedCall
from gemini_api import (
    GeminiApiError,
    parse_generate_content_response,
    raise_for_gemini_response,
    wrap_gemini_exception,
)

APP_VERSION = "1.8.0"
APP_UPDATED = "2026-05-26"

DEFAULT_MODEL_ID = "gemini-3.1-flash-lite"

GEMINI_MODELS: dict[str, dict[str, str]] = {
    "gemini-3.1-flash-lite": {
        "label": "Gemini 3.1 Flash-Lite",
        "description": (
            "Gemini 3 シリーズの安定版 Flash-Lite。"
            "大規模モデルに匹敵するパフォーマンスを、わずかな費用で実現します。"
        ),
    },
    "gemini-3.5-flash": {
        "label": "Gemini 3.5 Flash",
        "description": (
            "エージェントタスクとコーディングタスクで最先端のパフォーマンスを"
            "維持できる、最もインテリジェントなモデル。"
        ),
    },
}

# 後方互換のためデフォルトモデルのエイリアスを維持
MODEL_NAME = DEFAULT_MODEL_ID
MODEL_LABEL = GEMINI_MODELS[DEFAULT_MODEL_ID]["label"]
MODEL_DESCRIPTION = GEMINI_MODELS[DEFAULT_MODEL_ID]["description"]


def get_model_config(model_id: str | None = None) -> dict[str, str]:
    """モデル ID から表示名・説明を取得する."""
    mid = model_id or DEFAULT_MODEL_ID
    if mid not in GEMINI_MODELS:
        mid = DEFAULT_MODEL_ID
    cfg = GEMINI_MODELS[mid]
    return {
        "id": mid,
        "label": cfg["label"],
        "description": cfg["description"],
    }

CHANGELOG: list[dict] = [
    {
        "version": "1.8.0",
        "date": "2026-05-26",
        "changes": [
            "設問テンプレートの事前登録: OCR 時に Q1/Q2 一覧を参照し設問欠落を防止",
            "編集タブに原画像とのインライン校正（ページ単位の修正反映）を追加",
            "前回研修 vs 今回研修の Word アンケート比較機能を追加",
            "README を現行機能に合わせて全面更新",
        ],
    },
    {
        "version": "1.7.3",
        "date": "2026-05-26",
        "changes": [
            "使用モデルを選択可能に: デフォルト gemini-3.1-flash-lite / オプション gemini-3.5-flash",
            "使用モデル情報に Gemini API 公式モデル一覧ドキュメントへのリンクを追加",
        ],
    },
    {
        "version": "1.7.2",
        "date": "2026-05-26",
        "changes": [
            "複数ページPDF: 各ページで印字質問を省略せず Q1/Q2 形式で全文文字起こししてから質問別に集約",
            "2ページ目以降は1ページ目の質問一覧を参照し、設問欠落を防止",
            "全ページで設問が揃っている場合はローカル集約を優先し、回答の取り違えを低減",
        ],
    },
    {
        "version": "1.7.1",
        "date": "2026-05-25",
        "changes": [
            "使用モデルを gemini-3.1-flash-lite-preview から gemini-3.1-flash-lite（安定版）へ変更",
            "Gemini API エラー時に HTTP ステータス・モデル名・API 詳細をユーザーに表示",
            "モデル移行の可能性を案内し、公式モデル一覧ドキュメントへのリンクを表示",
        ],
    },
    {
        "version": "1.7.0",
        "date": "2026-05-25",
        "changes": [
            "複数ページPDFの質問欠落バグを根本対応: 複数ページは Gemini で全ページ一括統合方式に切替",
            "「質問ごとにまとめる」ボタンを 1 箇所（編集タブのステップ1）に集約し、重複を排除",
            "サイドバー折りたたみ・展開時のメイン領域重なりを解消し、自然なレイアウト変動に",
        ],
    },
    {
        "version": "1.6.0",
        "date": "2026-05-25",
        "changes": [
            "画面レイアウトを全面リフォーム: ワイドレイアウト化 + サイドバーに設定集約 + メインタブナビゲーション",
            "ファイル結果を「編集・加工 / 回答者ビュー / 定量分析 / 原画像対比 / ダウンロード」のタブ構成へ刷新",
            "原画像対比ビューを大幅拡大（6:4横並び または 縦並び大画面表示）、ページ画像のダウンロードも可能に",
            "定量サマリーを altair でスタイリッシュ化: センチメントはドーナツチャート、トピック/キーワードは横棒チャート（日本語ラベルが横向きで読みやすい）",
            "回答者単位ビューを 2 カラム カード表示に変更し、メタ情報チップで全体像を可視化",
            "メトリクスをカスタムカード（カラーアクセント＋サブテキスト）で表現",
            "コストダッシュボードに用途別の横棒チャートを追加",
        ],
    },
    {
        "version": "1.5.0",
        "date": "2026-05-25",
        "changes": [
            "複数ページPDF集約のバグ修正: ページごとに印字質問の認識が欠落しても、他ページから取得した canonical な質問文と Gemini による再割り当てで正しい質問の下に回答を集約",
            "回答者単位ビュー: [P.N] 識別子でグルーピングし、誰がどの質問にどう答えたかを縦読み可能に",
            "定量サマリー: センチメント分類・トピック抽出・キーワード頻度を JSON で取得し、可視化表示",
            "設問テンプレート自動生成: アップロード済みファイルから印字テキストを抽出し除外テンプレートを提案",
            "原画像との対比ビュー: 各ページ画像を保持し、文字起こしと並べて確認可能",
            "校正の差分プレビュー: 自動修正後テキストと原文の差分を色付き HTML で表示",
            "Excel 出力: 質問別シート・回答者別マトリクス・定量サマリーを 1 ファイルで",
            "API コスト・処理時間の可視化: 全 Gemini 呼び出しを集計し、目的別・累計サマリーを表示",
        ],
    },
    {
        "version": "1.4.0",
        "date": "2026-05-25",
        "changes": [
            "「質問ごとにまとめる」を強化: 質問境界を厳守し、Q2 の回答が Q1 に混入する事象を解消",
            "OCR 集約結果の各回答にページ識別子 [P.N] を付与し、同一回答者を追跡可能に",
            "新機能: 文字校正チェック（コピペしたテキストの誤字脱字・違和感のある文体・不自然な日本語を検出）",
            "校正レポートを Markdown 形式で表示し、Word ファイルとしてもダウンロード可能に",
        ],
    },
    {
        "version": "1.3.0",
        "date": "2026-05-01",
        "changes": [
            "文字起こしモード選択を追加（正確文字起こし / 校正文字起こし）",
            "校正文字起こし: 二重線訂正・丸印選択・不自然な表現をページ番号と位置付きで検出",
            "テキスト加工機能を追加（整文〈です・ます調〉・要約・カスタムプロンプト）",
            "加工後テキストをエディタブルなテキストエリアに即時反映",
            "アンケート分析機能を追加（傾向・課題・アクションプラン・研修改善提案）",
            "分析レポートを Word 形式でダウンロード可能に（Markdown 装飾を Word 書式に変換）",
            "API 503 エラー時の指数バックオフリトライを実装",
        ],
    },
    {
        "version": "1.2.0",
        "date": "2026-05-01",
        "changes": [
            "OCR 出力形式をシンプル化（書かれた文字をそのまま出力）",
            "マルチページ処理の統合ルール指示を削除",
            "除外テキストの例文を研修アンケート向けに変更",
            "テンプレート名の例を「研修アンケート用」に変更",
        ],
    },
    {
        "version": "1.1.0",
        "date": "2026-04-30",
        "changes": [
            "除外テキスト機能を追加（テキスト入力・スクリーンショット・テンプレート）",
            "画像前処理の強化（コントラスト・シャープネス・エッジ強調）",
            "マルチページ PDF の個別処理とバッチ処理のフォールバックを実装",
            "Word 出力フォーマットの改善",
        ],
    },
    {
        "version": "1.0.0",
        "date": "2026-04-01",
        "changes": [
            "初期リリース",
            "Gemini AI による手書き日本語文字認識",
            "PDF・画像ファイル（PNG / JPG / WebP など）の処理対応",
            "文字起こし結果の Word ファイル出力",
        ],
    },
]


def _get_api_key() -> str:
    """Resolve Gemini API key from Streamlit secrets or environment."""
    api_key = None
    try:
        api_key = st.secrets.get("GEMINI_API_KEY", None)  # type: ignore[attr-defined]
    except Exception:
        api_key = None
    if not api_key:
        api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY environment variable is required")
    return api_key


def extract_texts_from_screenshots(
    image_bytes_list: list[bytes],
    model_name: str | None = None,
) -> list[str]:
    """Extract printed text from screenshot images using Gemini (for exclusion)."""
    api_key = _get_api_key()
    model = model_name or DEFAULT_MODEL_ID
    if model not in GEMINI_MODELS:
        model = DEFAULT_MODEL_ID
    endpoint = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{model}:generateContent"
    )
    prompt = (
        "この画像に含まれる**印刷された文字**をすべて読み取り、"
        "1 行に 1 フレーズずつ出力してください。\n"
        "装飾・説明は不要です。テキストだけを返してください。"
    )

    extracted: list[str] = []
    for img_bytes in image_bytes_list:
        try:
            payload = {
                "contents": [{
                    "parts": [
                        {"text": prompt},
                        {"inline_data": {
                            "mime_type": "image/png",
                            "data": base64.b64encode(img_bytes).decode("utf-8"),
                        }},
                    ]
                }]
            }
            with TimedCall(
                purpose="除外テキスト抽出",
                model=model,
                input_chars=len(prompt),
                image_bytes=len(img_bytes),
            ) as tc:
                try:
                    res = requests.post(
                        f"{endpoint}?key={api_key}",
                        headers={"Content-Type": "application/json"},
                        json=payload,
                        timeout=60,
                    )
                    raise_for_gemini_response(res, model)
                    text = parse_generate_content_response(res.json())
                    tc.set_output(text or "")
                except Exception as _e_cost:
                    tc.mark_error(_e_cost)
                    raise wrap_gemini_exception(_e_cost, model) from _e_cost
            for line in text.strip().splitlines():
                line = line.strip()
                if line:
                    extracted.append(line)
        except GeminiApiError:
            raise
        except Exception as e:
            raise wrap_gemini_exception(e, model) from e
    return extracted


class OCRProcessor:
    def __init__(
        self,
        exclude_texts: list[str] | None = None,
        ocr_mode: str = "accurate",
        model_name: str | None = None,
        template_questions: dict[int, str] | None = None,
    ):
        """Initialize Gemini AI for OCR processing"""
        self.api_key = _get_api_key()
        self.model_name = model_name or DEFAULT_MODEL_ID
        if self.model_name not in GEMINI_MODELS:
            self.model_name = DEFAULT_MODEL_ID
        self.endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{self.model_name}:generateContent"
        self.exclude_texts = [t.strip() for t in (exclude_texts or []) if t.strip()]
        self.ocr_mode = ocr_mode
        self.template_questions = {
            int(k): v.strip()
            for k, v in (template_questions or {}).items()
            if v and str(v).strip()
        }

        self.ocr_prompt = self._build_ocr_prompt()

    def _merged_reference_questions(
        self,
        reference_questions: dict[int, str] | None,
    ) -> dict[int, str]:
        """事前登録テンプレートとページ参照を統合."""
        merged = dict(self.template_questions)
        if reference_questions:
            for q_num, q_text in reference_questions.items():
                existing = merged.get(q_num, "")
                if not existing or len(q_text) > len(existing):
                    merged[q_num] = q_text
        return merged

    def _build_ocr_prompt(self) -> str:
        """Route to mode-specific prompt builder."""
        if self.ocr_mode == "proofread":
            return self._build_proofread_prompt()
        return self._build_accurate_prompt()

    def _build_exclude_section(self) -> str:
        """Build exclude-text instructions section."""
        if not self.exclude_texts:
            return ""
        items = "\n".join(f"  - 「{t}」" for t in self.exclude_texts)
        return f"""
6. **文字起こし対象外テキスト（重要）**
   以下の文字列はアンケート用紙に印刷されたタイトル・ヘッダーなど、文字起こし不要のテキストです。
   これらの文字列が画像内に見つかった場合、出力から**完全に除外**してください。
   部分一致でも該当する場合は除外してください。
   **ただし**、設問（Q1/Q2 等の質問文）でその下に手書き回答欄があるものは、
   上記リストに該当しても**必ず文字起こしに含めてください**（設問の省略は禁止）。
{items}
"""

    def _build_page_ocr_prompt(
        self,
        page_num: int,
        total_pages: int = 1,
        reference_questions: dict[int, str] | None = None,
    ) -> str:
        """ページ単位 OCR 用プロンプト（複数ページ PDF では設問省略を禁止）."""
        base = self.ocr_prompt
        effective_ref = self._merged_reference_questions(reference_questions)

        ref_section = ""
        if effective_ref:
            ref_lines = "\n".join(
                f"Q{q_num}: {q_text}"
                for q_num, q_text in sorted(effective_ref.items())
            )
            source = (
                "事前登録された設問テンプレート"
                if self.template_questions
                else "1ページ目の OCR 結果"
            )
            ref_section = f"""
**参考：このアンケートの質問一覧（{source}）**
{ref_lines}

- 用紙上の印字質問は上記と同一です。
- 見出しは `Q1:`, `Q2:` … の形式で、**すべて省略せず**出力してください。
"""

        if total_pages <= 1 and not effective_ref:
            return base

        if total_pages <= 1 and effective_ref:
            return f"""{base}

**設問テンプレートに基づく追加ルール（厳守）**
- 用紙上に存在する**すべての印字質問**を `Q1: <質問文>`, `Q2: <質問文>` … の形式で出力してください。
- **質問文の省略・スキップ・要約は禁止**です。
- 各質問見出しの直下に、その質問に対する**手書き回答のみ**を記載してください。
{ref_section}
"""

        return f"""{base}

**複数ページPDF — ページ {page_num}/{total_pages} の追加ルール（厳守）**
- 全回答者に同じ印字質問が印刷されたアンケート用紙です。
- 用紙上に存在する**すべての印字質問**を `Q1: <質問文>`, `Q2: <質問文>` … の形式で出力してください。
- **質問文の省略・スキップ・要約は禁止**です。回答だけを並べる形式も禁止です。
- 各質問見出しの直下に、その質問に対する**手書き回答のみ**を記載してください。
- 手書き回答が空欄の質問も、質問見出しは必ず出力し、回答は「（無回答）」と記載してください。
- このページは回答者 {page_num} 人目に相当します（後で集約時に [P.{page_num}] として識別されます）。
{ref_section}
"""

    def _build_accurate_prompt(self) -> str:
        """Build accurate transcription prompt."""
        exclude_section = self._build_exclude_section()
        return f"""
あなたは日本語手書き文字認識の専門家です。画像を非常に注意深く観察し、書かれた文字を一字一句正確に読み取ってください。

**超精密文字認識手順：**

1. **画像全体の観察**
   - 全体レイアウトを把握
   - 質問と回答の位置関係を確認
   - 文字の配置パターンを理解

2. **文字単位での詳細分析**
   - 各文字の形状を慎重に観察
   - 画線の太さ、角度、曲がり具合を分析
   - 文字の一部が欠けていても、見える部分から判断

3. **日本語文字の特別な注意点**
   - ひらがな：「る/ろ」「は/ば/ぱ」「き/さ」「な/た」「わ/れ」「め/ぬ」
   - カタカナ：「ソ/ン」「シ/ツ」「ク/ワ」「ロ/コ」「エ/ユ」
   - 漢字：似た形の字に特に注意
   - 濁点・半濁点の有無を慎重に確認

4. **文字起こしの原則**
   - 書かれている通りに正確に転写
   - 推測や修正は一切しない
   - 誤字があってもそのまま記録
   - 判読不能な文字は「？」で表記

5. **出力形式**
   - 書かれている文字をそのまま出力する
   - 特定のフォーマットは適用しない
{exclude_section}
**絶対に守ること：**
- 見えた文字をそのまま書く
- 意味を考えて修正しない
- 推測で文字を補完しない

この画像の手書き文字を正確に読み取ってください：
"""

    def _build_proofread_prompt(self) -> str:
        """Build proofread mode prompt."""
        exclude_section = self._build_exclude_section()
        return f"""
あなたは日本語手書き文字認識・校正の専門家です。画像を非常に注意深く観察し、手書き文字の認識と校正を同時に行ってください。

**超精密文字認識手順：**

1. **画像全体の観察**
   - 全体レイアウトを把握
   - 質問と回答の位置関係を確認
   - 文字の配置パターンを理解
   - 訂正箇所（二重線・取り消し線・修正テープなど）を特定
   - 丸印・チェックマークなどの選択表示を確認

2. **文字単位での詳細分析**
   - 各文字の形状を慎重に観察
   - 画線の太さ、角度、曲がり具合を分析
   - 文字の一部が欠けていても、見える部分から判断

3. **日本語文字の特別な注意点**
   - ひらがな：「る/ろ」「は/ば/ぱ」「き/さ」「な/た」「わ/れ」「め/ぬ」
   - カタカナ：「ソ/ン」「シ/ツ」「ク/ワ」「ロ/コ」「エ/ユ」
   - 漢字：似た形の字に特に注意
   - 濁点・半濁点の有無を慎重に確認

4. **文字起こしの原則**
   - 訂正がある場合は訂正後の文字を採用して転写
   - 丸印・チェックで選択された選択肢を記録
   - 誤字があってもそのまま記録
   - 判読不能な文字は「？」で表記

5. **出力形式**
   - 書かれている文字（訂正後）をそのまま出力する
   - 特定のフォーマットは適用しない
   - 校正箇所は文字起こし本文の後に【校正注記】としてまとめる

**校正検出指示（重要）：**
以下を検出した場合、文字起こし本文の後に「【校正注記】」セクションを設けて記載してください。
各注記には「ページ番号」と「該当する質問番号または位置（回答欄の何行目など）」を必ず明記してください：
- **二重線・取り消し線による訂正**：「校正 [ページX・質問Y / 位置]: 「元の文字」→「訂正後の文字」」
- **丸印・チェックによる選択回答**：「選択 [ページX・質問Y / 位置]: 「選ばれた選択肢」」
- **文脈から不自然・矛盾している表現**：「要確認 [ページX・質問Y / 位置]: 「該当箇所」（理由を簡潔に）」
校正注記がない場合は【校正注記】セクション自体を省略してください。
{exclude_section}
**絶対に守ること：**
- 訂正後の文字を正として転写する
- 校正箇所は本文ではなく【校正注記】にまとめる
- 推測で文字を補完しない

この画像の手書き文字を正確に読み取り、校正箇所も検出してください：
"""

    def process_pdf(self, pdf_path):
        """Process PDF file and extract handwritten Japanese text using Gemini AI.

        テキストのみ返す。原画像が必要な場合は process_pdf_with_images() を使う。
        """
        text, _images = self.process_pdf_with_images(pdf_path)
        return text

    def process_pdf_with_images(self, pdf_path):
        """PDF 処理を行い、文字起こしテキストと各ページの PIL 画像リストを返す.

        Returns:
            (text, [(pil_image, page_num), ...])
        """
        try:
            pdf_document = fitz.open(pdf_path)
            all_images = []
            original_images = []

            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)

                mat = fitz.Matrix(6.0, 6.0)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")

                raw_image = Image.open(io.BytesIO(img_data))
                # オリジナル画像は対比ビュー用に別途保持（軽い縮小版）
                original_images.append((self._make_preview_image(raw_image), page_num + 1))

                image = self._enhance_image_for_ocr(raw_image.copy())
                image = self._apply_additional_preprocessing(image)

                all_images.append((image, page_num + 1))

            pdf_document.close()

            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)

            final_text = self._post_process_text(final_text)

            return final_text, original_images

        except Exception as e:
            raise Exception(f"PDF処理エラー: {str(e)}")
    
    def process_image(self, image_path):
        """Process image file (PNG, JPG, etc.) and extract handwritten Japanese text using Gemini AI."""
        text, _images = self.process_image_with_images(image_path)
        return text

    def process_image_with_images(self, image_path):
        """単一画像の処理結果（テキスト, 画像リスト）を返す."""
        try:
            image = Image.open(image_path)

            if image.mode in ("RGBA", "P", "LA"):
                background = Image.new("RGB", image.size, (255, 255, 255))
                if image.mode == "P":
                    image = image.convert("RGBA")
                background.paste(
                    image,
                    mask=image.split()[-1] if image.mode in ("RGBA", "LA") else None,
                )
                image = background
            elif image.mode != "RGB":
                image = image.convert("RGB")

            preview = self._make_preview_image(image)

            enhanced = self._enhance_image_for_ocr(image.copy())
            enhanced = self._apply_additional_preprocessing(enhanced)

            all_images = [(enhanced, 1)]
            original_images = [(preview, 1)]

            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)

            final_text = self._post_process_text(final_text)

            return final_text, original_images

        except Exception as e:
            raise Exception(f"画像処理エラー: {str(e)}")

    def _make_preview_image(self, image, max_dim: int = 1400):
        """対比ビュー表示用に元画像の縮小コピーを作る."""
        try:
            img = image.copy()
            if img.mode != "RGB":
                img = img.convert("RGB")
            w, h = img.size
            if max(w, h) > max_dim:
                scale = max_dim / max(w, h)
                img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
            return img
        except Exception:
            return image
    
    def process_images(self, image_paths):
        """Process multiple image files and extract handwritten Japanese text using Gemini AI"""
        try:
            all_images = []
            
            for idx, image_path in enumerate(image_paths):
                # Open image file
                image = Image.open(image_path)
                
                # Convert to RGB if necessary
                if image.mode in ('RGBA', 'P', 'LA'):
                    background = Image.new('RGB', image.size, (255, 255, 255))
                    if image.mode == 'P':
                        image = image.convert('RGBA')
                    background.paste(image, mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None)
                    image = background
                elif image.mode != 'RGB':
                    image = image.convert('RGB')
                
                # Apply image enhancement for OCR
                image = self._enhance_image_for_ocr(image)
                image = self._apply_additional_preprocessing(image)
                
                all_images.append((image, idx + 1))
            
            # Try individual processing first, fallback to batch if needed
            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)
            
            # Post-process the text
            final_text = self._post_process_text(final_text)
            
            return final_text
            
        except Exception as e:
            raise Exception(f"画像処理エラー: {str(e)}")
    
    def process_image_bytes(self, image_bytes, filename="image"):
        """Process image from bytes data"""
        try:
            # Open image from bytes
            image = Image.open(io.BytesIO(image_bytes))
            
            # Convert to RGB if necessary
            if image.mode in ('RGBA', 'P', 'LA'):
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                background.paste(image, mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None)
                image = background
            elif image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Apply image enhancement for OCR
            image = self._enhance_image_for_ocr(image)
            image = self._apply_additional_preprocessing(image)
            
            # Process as single image
            all_images = [(image, 1)]
            
            # Try individual processing
            try:
                final_text = self._process_pages_individually(all_images)
            except Exception as e:
                print(f"Individual processing failed, trying batch processing: {e}")
                final_text = self._extract_text_from_all_images(all_images)
            
            # Post-process the text
            final_text = self._post_process_text(final_text)
            
            return final_text
            
        except Exception as e:
            raise Exception(f"画像処理エラー: {str(e)}")
    
    def _enhance_image_for_ocr(self, image):
        """Enhance image quality for better OCR accuracy"""
        try:
            # Convert to RGB if necessary
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # 1. Increase contrast for better text visibility
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(1.3)
            
            # 2. Increase sharpness to make text edges clearer
            enhancer = ImageEnhance.Sharpness(image)
            image = enhancer.enhance(1.5)
            
            # 3. Adjust brightness if too dark
            enhancer = ImageEnhance.Brightness(image)
            image = enhancer.enhance(1.1)
            
            # 4. Apply slight noise reduction
            image = image.filter(ImageFilter.MedianFilter(size=1))
            
            # 5. Apply unsharp mask for text clarity
            image = image.filter(ImageFilter.UnsharpMask(radius=1, percent=150, threshold=3))
            
            return image
            
        except Exception as e:
            # Return original image if enhancement fails
            return image
    
    def _apply_additional_preprocessing(self, image):
        """Apply advanced preprocessing techniques for maximum OCR accuracy"""
        try:
            # Convert to grayscale for better text recognition
            if image.mode != 'L':
                image = image.convert('L')
            
            # Convert to numpy array for advanced processing
            img_array = np.array(image)
            
            # Apply adaptive threshold to improve text contrast
            from PIL import ImageOps
            image = Image.fromarray(img_array)
            
            # Increase contrast more aggressively
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(2.0)
            
            # Apply edge enhancement
            image = image.filter(ImageFilter.EDGE_ENHANCE_MORE)
            
            # Convert back to RGB for Gemini compatibility
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            return image
            
        except Exception as e:
            # Return original image if preprocessing fails
            return image
    
    def _resize_image(self, image, scale_factor):
        """Resize image by scale factor while maintaining aspect ratio"""
        try:
            original_size = image.size
            new_size = (int(original_size[0] * scale_factor), int(original_size[1] * scale_factor))
            # Use LANCZOS resampling for better quality
            resized_image = image.resize(new_size, Image.Resampling.LANCZOS)
            print(f"画像リサイズ: {original_size} -> {new_size}")
            return resized_image
        except Exception as e:
            print(f"画像リサイズエラー: {e}")
            return image
    
    def _ensure_max_size(self, image, max_dimension=4096):
        """Ensure image doesn't exceed maximum dimension for API"""
        try:
            width, height = image.size
            if width <= max_dimension and height <= max_dimension:
                return image
            
            # Calculate scale factor to fit within max_dimension
            scale = min(max_dimension / width, max_dimension / height)
            new_size = (int(width * scale), int(height * scale))
            resized = image.resize(new_size, Image.Resampling.LANCZOS)
            print(f"画像サイズ制限適用: {image.size} -> {new_size}")
            return resized
        except Exception as e:
            print(f"画像サイズ制限エラー: {e}")
            return image
    
    def _is_text_valid(self, text):
        """Check if extracted text is valid (not empty or error message)"""
        if not text or not text.strip():
            return False
        
        # Check for error messages
        error_indicators = [
            "テキストが検出されませんでした",
            "OCRエラー",
            "文字起こしに失敗しました",
            "文字起こし結果がありません"
        ]
        
        text_lower = text.lower()
        for indicator in error_indicators:
            if indicator in text:
                return False
        
        # Check if text is too short (likely invalid)
        if len(text.strip()) < 10:
            return False
        
        return True
    
    def _extract_text_from_image(
        self,
        image,
        page_num,
        resize_on_failure=True,
        ocr_prompt: str | None = None,
    ):
        """Extract text from image using Gemini AI with automatic resize retry on failure"""
        prompt = ocr_prompt if ocr_prompt is not None else self.ocr_prompt
        # First, ensure image doesn't exceed API limits
        image = self._ensure_max_size(image, max_dimension=4096)
        
        # Try with original image first, then progressively smaller
        scale_factors = [1.0, 0.75, 0.5, 0.35, 0.25] if resize_on_failure else [1.0]
        
        print(f"ページ {page_num}: OCR処理開始 (元画像サイズ: {image.size})")
        
        for scale_factor in scale_factors:
            try:
                # Resize image if needed
                current_image = image
                if scale_factor < 1.0:
                    current_image = self._resize_image(image, scale_factor)
                    print(f"ページ {page_num}: 画像を{scale_factor*100:.0f}%に縮小して再試行中...")
                
                # Convert PIL Image to bytes
                img_byte_arr = io.BytesIO()
                current_image.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                
                # Check image size in bytes
                img_size_mb = len(img_byte_arr) / (1024 * 1024)
                print(f"ページ {page_num}: 画像サイズ {img_size_mb:.2f}MB, 解像度 {current_image.size}")
                
                # If image is too large, skip this scale and try smaller
                if img_size_mb > 20:
                    print(f"ページ {page_num}: 画像サイズが大きすぎます ({img_size_mb:.2f}MB)。縮小します...")
                    continue
                
                # Generate content with Gemini
                headers = {
                    "Content-Type": "application/json"
                }
                payload = {
                    "contents": [
                        {
                            "parts": [
                                {"text": prompt},
                                {
                                    "inline_data": {
                                        "mime_type": "image/png",
                                        "data": base64.b64encode(img_byte_arr).decode("utf-8")
                                    }
                                }
                            ]
                        }
                    ]
                }
                
                print(f"ページ {page_num}: Gemini API呼び出し中...")
                with TimedCall(
                    purpose=f"OCR: 個別ページ (P.{page_num})",
                    model=self.model_name,
                    input_chars=len(prompt),
                    image_bytes=len(img_byte_arr),
                ) as tc:
                    try:
                        res = requests.post(
                            f"{self.endpoint}?key={self.api_key}",
                            headers=headers,
                            json=payload,
                            timeout=180
                        )
                        print(f"ページ {page_num}: APIレスポンス status={res.status_code}")
                        res.raise_for_status()
                        data = res.json()
                        _resp_text_for_cost = (
                            data.get("candidates", [{}])[0]
                            .get("content", {})
                            .get("parts", [{}])[0]
                            .get("text", "")
                        )
                        tc.set_output(_resp_text_for_cost)
                    except Exception as _e_cost:
                        tc.mark_error(_e_cost)
                        raise
                
                # Debug: print response structure
                if "error" in data:
                    print(f"ページ {page_num}: APIエラー: {data['error']}")
                    continue
                
                text = (
                    data.get("candidates", [{}])[0]
                    .get("content", {})
                    .get("parts", [{}])[0]
                    .get("text", "")
                )
                
                if text:
                    text = text.strip()
                    print(f"ページ {page_num}: テキスト取得成功 ({len(text)}文字)")
                    # Check if text is valid
                    if self._is_text_valid(text):
                        if scale_factor < 1.0:
                            print(f"ページ {page_num}: 縮小画像({scale_factor*100:.0f}%)で文字読み取り成功")
                        return text
                    # If text is invalid and we have more scale factors to try, continue
                    elif scale_factor != scale_factors[-1]:
                        print(f"ページ {page_num}: テキストが無効 (短すぎるか、エラーメッセージ)。次の縮小率を試行します...")
                        continue
                    else:
                        # Last attempt, return what we got
                        return text if text else f"ページ {page_num}: テキストが検出されませんでした"
                else:
                    print(f"ページ {page_num}: テキストが空です")
                    # Empty response, try next scale factor if available
                    if scale_factor != scale_factors[-1]:
                        continue
                    return f"ページ {page_num}: テキストが検出されませんでした"
                    
            except Exception as e:
                error_msg = str(e)
                print(f"ページ {page_num}: エラー発生 - {error_msg}")
                api_err = wrap_gemini_exception(e, self.model_name)
                # モデル未提供・認可エラー等は即座に上位へ伝播（ユーザーに表示させる）
                if api_err.status_code in (400, 403, 404):
                    raise api_err from e
                # If quota exceeded, don't retry
                if "429" in error_msg or "quota" in error_msg.lower():
                    return f"ページ {page_num}: OCRエラー - {str(e)}"
                
                # If this is the last attempt, return error
                if scale_factor == scale_factors[-1]:
                    return f"ページ {page_num}: OCRエラー - {str(e)}"
                
                # Otherwise, try next scale factor
                print(f"ページ {page_num}: エラー発生。次の縮小率を試行します: {str(e)}")
                continue
        
        return f"ページ {page_num}: すべての試行が失敗しました"
    
    def _extract_text_from_all_images(self, images_with_pages):
        """Extract text from multiple images with question consolidation"""
        try:
            multi_prompt = self.ocr_prompt
            
            # Create content parts list
            content_parts = [multi_prompt]
            
            # Add all images
            for image, page_num in images_with_pages:
                # Convert PIL Image to bytes
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                
                # Create Gemini compatible image
                gemini_image = {
                    "mime_type": "image/png",
                    "data": img_byte_arr
                }
                
                content_parts.append(f"\n--- ページ {page_num} ---")
                content_parts.append(gemini_image)
            
            # Build payload for multi-page
            parts = []
            for part in content_parts:
                if isinstance(part, str):
                    parts.append({"text": part})
                else:
                    # image dict with bytes
                    parts.append({
                        "inline_data": {
                            "mime_type": part.get("mime_type", "image/png"),
                            "data": base64.b64encode(part["data"]).decode("utf-8")
                        }
                    })
            payload = {"contents": [{"parts": parts}]}
            total_img_bytes = sum(
                len(p.get("inline_data", {}).get("data", "")) for p in parts
                if isinstance(p, dict) and "inline_data" in p
            )
            with TimedCall(
                purpose="OCR: 複数ページ一括",
                model=self.model_name,
                input_chars=len(multi_prompt),
                image_bytes=total_img_bytes,
            ) as tc:
                try:
                    res = requests.post(
                        f"{self.endpoint}?key={self.api_key}",
                        headers={"Content-Type": "application/json"},
                        json=payload,
                        timeout=300
                    )
                    res.raise_for_status()
                    data = res.json()
                    text = (
                        data.get("candidates", [{}])[0]
                        .get("content", {})
                        .get("parts", [{}])[0]
                        .get("text", "")
                    )
                    tc.set_output(text or "")
                except Exception as _e_cost:
                    tc.mark_error(_e_cost)
                    raise
            if text:
                return text.strip()
            else:
                return "テキストが検出されませんでした"
                
        except Exception as e:
            # Fallback to individual page processing if multi-image fails
            all_text = []
            for image, page_num in images_with_pages:
                page_text = self._extract_text_from_image(image, page_num)
                if page_text and page_text.strip():
                    all_text.append(f"--- ページ {page_num} ---\n{page_text}")
            
            return "\n\n".join(all_text) if all_text else "文字起こし結果がありません。"
    
    def _process_pages_individually(self, images_with_pages):
        """Process each page individually with retry logic for maximum accuracy"""
        total_pages = len(images_with_pages)
        all_page_results = []
        reference_questions: dict[int, str] = {}

        for idx, (image, page_num) in enumerate(images_with_pages):
            ref = reference_questions if idx > 0 and reference_questions else None
            page_text = self._extract_with_retry(
                image,
                page_num,
                total_pages=total_pages,
                reference_questions=ref,
            )
            if page_text and page_text.strip():
                all_page_results.append({
                    'page_num': page_num,
                    'text': page_text.strip(),
                })
                for q_num, q_text, _ in self._parse_questions_from_text(page_text):
                    existing = reference_questions.get(q_num, "")
                    if len(q_text) > len(existing):
                        reference_questions[q_num] = q_text
        
        # If no results, return empty string
        if not all_page_results:
            return ""
        
        # Consolidate questions across pages
        result = self._consolidate_questions_from_pages(all_page_results)
        
        # If consolidation returned empty, return raw texts
        if not result or not result.strip():
            raw_texts = []
            for page_result in all_page_results:
                raw_texts.append(f"--- ページ {page_result['page_num']} ---\n{page_result['text']}")
            return "\n\n".join(raw_texts)
        
        return result
    
    def _extract_with_retry(
        self,
        image,
        page_num,
        max_retries=1,
        total_pages: int = 1,
        reference_questions: dict[int, str] | None = None,
    ):
        """Extract text with multiple attempts and automatic resize on failure"""
        page_prompt = self._build_page_ocr_prompt(
            page_num, total_pages, reference_questions
        )
        result = self._extract_text_from_image(
            image,
            page_num,
            resize_on_failure=True,
            ocr_prompt=page_prompt,
        )
        
        # If result is valid, return it
        if self._is_text_valid(result):
            return result
        
        # If still invalid after resize attempts, try with enhanced prompt
        scale_factors = [1.0, 0.8, 0.6, 0.4]
        best_result = ""
        
        for scale_factor in scale_factors:
            try:
                # Resize image if needed
                current_image = image
                if scale_factor < 1.0:
                    current_image = self._resize_image(image, scale_factor)
                    print(f"ページ {page_num}: 拡張プロンプトで画像を{scale_factor*100:.0f}%に縮小して再試行中...")
                
                # Convert PIL Image to bytes with high quality
                img_byte_arr = io.BytesIO()
                current_image.save(img_byte_arr, format='PNG', quality=100, optimize=False)
                img_byte_arr = img_byte_arr.getvalue()
                
                # Enhanced prompt for individual page processing
                proofread_hint = (
                    f"\n- 【校正注記】の各項目には必ず「ページ{page_num}」と"
                    f"該当する質問番号・位置（例: ページ{page_num}・Q1回答欄1行目）を記載してください"
                    if self.ocr_mode == "proofread" else ""
                )
                individual_prompt = f"""
{page_prompt}

**ページ {page_num} の再試行（より厳密に）：**
- 用紙上の**すべての印字質問**を `Q1:`, `Q2:` … 形式で省略なく出力してください。
- 各質問の直下に手書き回答を配置してください（質問と回答の対応を絶対に取り違えない）。
- 質問を1つにまとめたり、回答だけを並べたりしないでください。{proofread_hint}

画像を慎重に分析し、印刷質問文と手書き回答を正確に文字起こししてください：
"""
                
                # REST call
                payload = {
                    "contents": [
                        {
                            "parts": [
                                {"text": individual_prompt},
                                {
                                    "inline_data": {
                                        "mime_type": "image/png",
                                        "data": base64.b64encode(img_byte_arr).decode("utf-8")
                                    }
                                }
                            ]
                        }
                    ]
                }
                with TimedCall(
                    purpose=f"OCR: 拡張プロンプト (P.{page_num})",
                    model=self.model_name,
                    input_chars=len(individual_prompt),
                    image_bytes=len(img_byte_arr),
                ) as tc:
                    try:
                        res = requests.post(
                            f"{self.endpoint}?key={self.api_key}",
                            headers={"Content-Type": "application/json"},
                            json=payload,
                            timeout=120
                        )
                        res.raise_for_status()
                        data = res.json()
                        current_result = (
                            data.get("candidates", [{}])[0]
                            .get("content", {})
                            .get("parts", [{}])[0]
                            .get("text", "")
                        ).strip()
                        tc.set_output(current_result)
                    except Exception as _e_cost:
                        tc.mark_error(_e_cost)
                        raise
                
                if current_result and self._is_text_valid(current_result):
                    # Found valid result
                    if scale_factor < 1.0:
                        print(f"ページ {page_num}: 拡張プロンプトで縮小画像({scale_factor*100:.0f}%)で文字読み取り成功")
                    return current_result
                
                # Keep the longest/most detailed result
                if len(current_result) > len(best_result):
                    best_result = current_result
                
            except Exception as e:
                error_message = str(e)
                print(f"ページ {page_num}: 拡張プロンプト試行失敗 (縮小率{scale_factor*100:.0f}%): {e}")
                
                # If quota exceeded, don't retry
                if "429" in error_message or "quota" in error_message.lower():
                    print(f"Quota exceeded for page {page_num}, skipping retries")
                    break
                
                # Continue to next scale factor
                continue
        
        # Return best result if valid, otherwise return error message
        if self._is_text_valid(best_result):
            return best_result
        
        return best_result if best_result else f"ページ {page_num}: 文字起こしに失敗しました"
    
    def _consolidate_questions_from_pages(self, page_results):
        """Consolidate questions and answers from multiple pages.

        改良ポイント (v1.7.2):
        - 各ページ OCR で Q1/Q2 形式の設問を省略せず取得することを前提とする。
        - 全ページで設問が揃っている場合はローカル集約のみ（回答の取り違えを防止）。
        - 設問欠落がある場合のみ Gemini 統合 → ローカル集約の順でフォールバック。
        """
        if not page_results:
            return ""

        if len(page_results) == 1:
            return self._local_consolidate(page_results)

        if self._pages_have_full_question_coverage(page_results):
            return self._local_consolidate(page_results)

        try:
            unified = self._unified_consolidate_via_gemini(page_results)
            if unified and unified.strip() and "Q1" in unified.upper().replace("Ｑ", "Q"):
                return unified
        except Exception as e:
            print(f"統合集約が失敗、ローカル集約にフォールバック: {e}")

        return self._local_consolidate(page_results)

    def _build_canonical_questions(self, page_results: list[dict]) -> dict[int, str]:
        """全ページから canonical な質問文を集約（最長表記を採用）."""
        canonical: dict[int, str] = {}
        for pr in page_results:
            for q_num, q_text, _ in self._parse_questions_from_text(pr.get('text', '')):
                existing = canonical.get(q_num, "")
                if not existing or len(q_text) > len(existing):
                    canonical[q_num] = q_text
        return canonical

    def _pages_have_full_question_coverage(self, page_results: list[dict]) -> bool:
        """全ページで canonical 設問が OCR 結果に含まれているか."""
        canonical = self._build_canonical_questions(page_results)
        if len(canonical) < 1:
            return False
        required = set(canonical.keys())
        for pr in page_results:
            parsed = self._parse_questions_from_text(pr.get('text', ''))
            page_q_nums = {q_num for q_num, _, _ in parsed}
            if not required.issubset(page_q_nums):
                return False
        return True

    def _local_consolidate(self, page_results):
        """各ページをローカルパースし、不足分は Gemini に再割り当てさせる従来パス."""
        if not page_results:
            return ""

        # --- 1) 各ページをローカルパース ---
        parsed_pages: list[dict] = []
        for page_result in page_results:
            page_num = page_result['page_num']
            text = page_result['text'] or ""
            questions = self._parse_questions_from_text(text)
            parsed_pages.append({
                'page_num': page_num,
                'raw_text': text,
                'questions': questions,
            })

        # --- 2) canonical 質問文を集約（最長表記を採用）---
        canonical_questions = self._build_canonical_questions(page_results)

        # canonical が無ければ最低限のフォールバック
        if not canonical_questions:
            raw_texts = [
                f"--- ページ {pp['page_num']} ---\n{pp['raw_text']}"
                for pp in parsed_pages
                if pp['raw_text'].strip()
            ]
            return "\n\n".join(raw_texts)

        # --- 3) 質問辞書を canonical で初期化 ---
        questions_dict: dict[int, dict] = {
            q_num: {'question': q_text, 'answers': []}
            for q_num, q_text in canonical_questions.items()
        }

        # --- 4) ローカルパースで取得済みの回答を投入 ---
        pages_needing_reassignment: list[tuple[int, str, list[int]]] = []
        for pp in parsed_pages:
            page_num = pp['page_num']
            raw_text = pp['raw_text']
            page_q_nums = set()
            for q_num, _q_text, answer in pp['questions']:
                page_q_nums.add(q_num)
                if answer and answer.strip() and answer.strip() != '（無回答）':
                    questions_dict[q_num]['answers'].append({
                        'page': page_num,
                        'text': answer.strip(),
                    })

            # このページが canonical 全質問をカバーしていない & テキストが残っていれば
            # 再割り当て対象に積む（小さな PDF や OCR が全質問を取れていれば不要）
            missing = [
                q_num
                for q_num in canonical_questions
                if q_num not in page_q_nums
            ]
            if missing and raw_text.strip() and self._has_answer_like_content(raw_text):
                pages_needing_reassignment.append((page_num, raw_text, missing))

        # --- 5) Gemini に再割り当てを依頼 ---
        if len(canonical_questions) > 1 and pages_needing_reassignment:
            for page_num, raw_text, missing in pages_needing_reassignment:
                assignments = self._reassign_page_to_canonical(
                    page_text=raw_text,
                    canonical_questions=canonical_questions,
                    target_q_nums=missing,
                    page_num=page_num,
                )
                for q_num, answer in assignments:
                    if not answer or not answer.strip():
                        continue
                    if answer.strip() == '（無回答）':
                        continue
                    # 既に同一テキストが入っていれば重複を避ける
                    existing_texts = {
                        a['text'] for a in questions_dict[q_num]['answers']
                        if a['page'] == page_num
                    }
                    if answer.strip() in existing_texts:
                        continue
                    questions_dict[q_num]['answers'].append({
                        'page': page_num,
                        'text': answer.strip(),
                    })

        # --- 6) 出力組み立て ---
        result_lines: list[str] = []
        for q_num in sorted(questions_dict.keys()):
            q_data = questions_dict[q_num]
            result_lines.append(f"Q{q_num}: {q_data['question']}")
            if q_data['answers']:
                for ans in sorted(q_data['answers'], key=lambda a: a['page']):
                    result_lines.append(f"・[P.{ans['page']}] {ans['text']}")
            else:
                result_lines.append("・（無回答）")
            result_lines.append("")

        return "\n".join(result_lines)

    def _unified_consolidate_via_gemini(self, page_results: list[dict]) -> str:
        """全ページの OCR テキストを 1 回の Gemini 呼び出しで統合集約する.

        各ページの OCR で印字質問が一部欠落していても、ページ全体の文脈から
        canonical 質問の特定と回答の分類を Gemini に任せる。
        """
        if not page_results:
            return ""

        pages_section = "\n\n".join(
            f"--- ページ {pr['page_num']} ---\n{pr.get('text', '').strip()}"
            for pr in page_results
            if pr.get('text', '').strip()
        )
        if not pages_section:
            return ""

        prompt = f"""\
あなたはアンケート集約の専門家です。以下は同一のアンケート用紙を複数の回答者が記入した
複数ページの OCR 結果です。各ページは 1 名の回答者に相当し、ページ番号 = 回答者番号です。
各ページは `Q1:`, `Q2:` … 形式で印字質問と手書き回答が記載されている想定ですが、
一部ページで設問見出しの OCR 欠落がある可能性があります。

【ページごとの OCR テキスト】
{pages_section}

【あなたのタスク】
1. 各ページ内で **回答はそのページ内の質問見出し（Q1/Q2…）の直下に書かれていたもの** として扱い、
   質問と回答の対応関係をページ内の並び順から厳密に復元してください。
2. このアンケート用紙に印字されている **canonical な質問一覧** を確定してください。
   - 欠落している設問見出しは他ページから補完し、最も完全な表記を採用してください。
   - 質問番号は Q1, Q2, ... の昇順で番号付けしてください。
3. 各回答（手書きの内容）を **正しい質問の下** に分類してください。
   - **絶対に質問の境界を超えて回答を混在させない**でください。
     Q2 の回答が Q1 に入る・Q1 の回答が Q2 に入るような取り違えは禁止です。
   - ページ {{N}} に書かれていた回答には、必ず先頭に `[P.{{N}}]` を付けてください。
   - 質問1だけ・質問2だけしか回答していない回答者がいても、その回答は該当質問の下にだけ配置してください。
4. 該当する手書き回答がない質問の下には `・（無回答）` と 1 行だけ書いてください。
5. **回答の取りこぼし禁止**: 各ページに書かれている手書きと思しき内容は、原則として
   いずれかの質問の下に必ず配置してください（印字された質問文や見出しを回答として
   出力してはいけません。手書きと判別できないものは除外してください）。

【出力形式（厳守）】
Q1: <印字された質問文をそのまま転記>
・[P.1] 回答内容
・[P.2] 回答内容

Q2: <印字された質問文をそのまま転記>
・[P.1] 回答内容
・[P.3] 回答内容

【出力に関する追加ルール】
- 余計な見出し・前置き・注釈・締めの文章は一切付けないでください。出力は上記の構造のみとしてください。
- 質問文や印字テキストを回答として出力してはいけません（手書きの内容のみ）。
- Markdown コードフェンスや装飾は使わないでください。
"""
        return self._call_gemini_text(
            prompt=prompt,
            purpose="OCR: 全ページ統合集約",
            timeout=240,
        )

    def _has_answer_like_content(self, text: str) -> bool:
        """質問ヘッダ以外に回答らしき行が含まれているか判定."""
        for line in text.splitlines():
            s = line.strip()
            if not s:
                continue
            if re.match(r'^Q\d+\s*[:：]', s):
                continue
            if s.startswith('---'):
                continue
            if s.startswith('【'):
                continue
            return True
        return False

    def _reassign_page_to_canonical(
        self,
        page_text: str,
        canonical_questions: dict[int, str],
        target_q_nums: list[int],
        page_num: int,
    ) -> list[tuple[int, str]]:
        """canonical 質問リストを基に、ページのテキストから対象質問への回答を Gemini に抽出させる."""
        if not target_q_nums:
            return []

        all_q_text = "\n".join(
            f"Q{q_num}: {canonical_questions[q_num]}"
            for q_num in sorted(canonical_questions.keys())
        )
        target_q_text = "\n".join(
            f"Q{q_num}: {canonical_questions[q_num]}"
            for q_num in sorted(target_q_nums)
        )

        prompt = f"""\
あなたはアンケート集約の専門家です。
以下のページ（ページ{page_num}）のOCRテキストから、各質問への回答を抽出してください。
このページでは印字された質問文の OCR 認識が一部欠落しているため、
他ページから取得した canonical な質問文一覧を参照して回答を正しく分類してください。

【canonical 質問一覧（参考用）】
{all_q_text}

【このページで抽出してほしい質問】
{target_q_text}

【出力形式（厳守）】
Q<番号>: <該当する回答テキスト1行のみ>
Q<番号>: <該当する回答テキスト1行のみ>
...

- 1 つの質問に対する複数行の回答は半角スペース区切りで 1 行にまとめてください。
- 該当する回答が見つからない質問は出力しないでください（行をスキップ）。
- 回答が手書きで存在しない・空欄である場合も出力しないでください。
- 余計な前置き・注釈・Markdown 記号は禁止です。
- 質問文や印字テキストを回答として出力してはいけません（手書きの内容のみ）。

【ページ{page_num} の OCR テキスト】
{page_text}
"""
        try:
            response_text = self._call_gemini_text(
                prompt=prompt,
                purpose=f"OCR: 再割当て (P.{page_num})",
                timeout=90,
            )
        except Exception as e:
            print(f"_reassign_page_to_canonical page {page_num} failed: {e}")
            return []

        return self._parse_qnum_assignments(response_text, target_q_nums)

    def _parse_qnum_assignments(
        self,
        text: str,
        target_q_nums: list[int],
    ) -> list[tuple[int, str]]:
        """`Q1: 回答` 形式の応答を (q_num, answer) のリストに変換."""
        results: list[tuple[int, str]] = []
        target_set = set(target_q_nums)
        for line in (text or "").splitlines():
            s = line.strip()
            if not s:
                continue
            m = re.match(r'^Q\s*(\d+)\s*[:：]\s*(.+)$', s)
            if not m:
                continue
            q_num = int(m.group(1))
            ans = m.group(2).strip()
            if q_num in target_set and ans:
                results.append((q_num, ans))
        return results

    def _call_gemini_text(
        self,
        prompt: str,
        purpose: str = "OCR補助",
        timeout: int = 120,
    ) -> str:
        """テキストのみの軽量 Gemini 呼び出し（コスト記録あり）."""
        payload = {"contents": [{"parts": [{"text": prompt}]}]}
        with TimedCall(
            purpose=purpose,
            model=self.model_name,
            input_chars=len(prompt),
        ) as tc:
            try:
                res = requests.post(
                    f"{self.endpoint}?key={self.api_key}",
                    headers={"Content-Type": "application/json"},
                    json=payload,
                    timeout=timeout,
                )
                raise_for_gemini_response(res, self.model_name)
                text = parse_generate_content_response(res.json())
                tc.set_output(text)
                return text
            except Exception as e:
                tc.mark_error(e)
                raise wrap_gemini_exception(e, self.model_name) from e
    
    def _parse_questions_from_text(self, text):
        """Parse questions and answers from text"""
        questions = []
        lines = text.split('\n')
        current_question = None
        current_q_num = None
        current_answer = ""
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Check if line is a question
            q_match = re.match(r'^(?:Q|質問)\s*(\d+)\s*[:：]\s*(.+)', line, re.IGNORECASE)
            if q_match:
                # Save previous question if exists
                if current_question and current_q_num:
                    questions.append((current_q_num, current_question, current_answer.strip()))
                
                # Start new question
                current_q_num = int(q_match.group(1))
                current_question = q_match.group(2).strip()
                current_answer = ""
            else:
                # This is likely an answer
                if current_question:
                    if current_answer:
                        current_answer += " " + line
                    else:
                        current_answer = line
        
        # Add the last question
        if current_question and current_q_num:
            questions.append((current_q_num, current_question, current_answer.strip()))
        
        return questions
    
    def _post_process_text(self, text):
        """Post-process extracted text for better readability"""
        if not text:
            return "文字起こし結果がありません。"
        
        # Clean up common OCR artifacts
        lines = text.split('\n')
        cleaned_lines = []
        
        for line in lines:
            line = line.strip()
            if line:  # Skip empty lines
                # Remove "回答：" prefix if present
                if line.startswith('回答：'):
                    line = line[3:].strip()
                elif line.startswith('回答:'):
                    line = line[3:].strip()
                
                # Normalize spaces - remove irregular half-width/full-width spaces
                import re
                # Replace multiple spaces with single space
                line = re.sub(r'\s+', ' ', line)
                # Remove spaces around Japanese punctuation
                line = re.sub(r'\s*([。、！？：；])\s*', r'\1', line)
                # Remove unnecessary spaces in Japanese text
                line = re.sub(r'([あ-ん])\s+([あ-ん])', r'\1\2', line)
                line = re.sub(r'([ア-ン])\s+([ア-ン])', r'\1\2', line)
                line = re.sub(r'([一-龯])\s+([一-龯])', r'\1\2', line)
                
                cleaned_lines.append(line)
        
        # Join lines with proper spacing
        result = '\n'.join(cleaned_lines)
        
        # Add final processing note
        if result:
            result += "\n\n--- 注意 ---\n手書き文字の自動認識結果です。不正確な部分がある可能性があります。"
        
        return result if result else "文字起こし結果がありません。"

# 手書きアンケート OCR・文字起こしアプリ

手書きのアンケート（PDF / 画像）を Google Gemini AI で OCR 処理し、文字起こし・編集・分析・出力までを一括で行う Streamlit アプリケーションです。

**現在のバージョン: v1.8.0**

## 主な機能

### 文字起こし

- PDF・画像（PNG / JPG / WebP など）の複数アップロード
- Gemini AI による手書き日本語 OCR
- **正確文字起こし** / **校正文字起こし**（訂正線・丸印検出）の 2 モード
- 複数ページ PDF の設問別集約（各回答に `[P.N]` 識別子付与）
- **設問テンプレートの事前登録** — OCR 前に Q1/Q2 を登録し、設問欠落を防止
- 除外テキスト（タイトル・固定文）の指定とテンプレート管理
- 使用モデル選択: `gemini-3.1-flash-lite`（デフォルト）/ `gemini-3.5-flash`

### 編集・分析

- **質問ごとにまとめる** — 質問単位で回答を集約
- **原画像とのインライン校正** — 編集タブで原画像と並べてページ単位修正
- 原画像対比ビュー（横並び / 縦並び）
- テキスト加工（整文・要約・翻訳・敬語化・カスタムプロンプト）
- 文字校正チェック（コピペ用・差分プレビュー付き）
- 回答者単位ビュー
- 定量サマリー（センチメント・トピック・キーワード）
- アンケート分析レポート（傾向・改善提案）
- **前回研修 vs 今回研修の比較** — Word ファイル 2 つをアップロードして比較レポート生成

### 出力

- Word（.docx）ダウンロード
- Excel（質問別・回答者別・定量サマリー）
- 設問テンプレート自動生成

### その他

- API 利用状況・処理時間・概算コストの可視化
- Gemini API エラー時の詳細表示と公式ドキュメントリンク

## 必要要件

- Python 3.11 以上
- [Google Gemini API Key](https://ai.google.dev/)

## インストール

```bash
uv sync
```

または:

```bash
pip install -r requirements.txt
```

## 使い方

### 1. API キーの設定

`.streamlit/secrets.toml` を作成:

```toml
GEMINI_API_KEY = "your-api-key-here"
```

または環境変数:

```bash
export GEMINI_API_KEY="your-api-key-here"
```

### 2. アプリの起動

```bash
uv run streamlit run app.py
```

ブラウザで `http://localhost:8501` を開きます。

### 3. 基本的なワークフロー

1. **サイドバー**で文字起こしモード・除外テキスト・**設問テンプレート**・使用モデルを設定
2. **アップロード & 結果**タブで PDF / 画像をアップロードし文字起こし
3. 各ファイルの **編集・加工**タブで「質問ごとにまとめる」を実行
4. 必要に応じて **原画像と並べて校正** でページ単位修正
5. **回答者ビュー** / **定量分析** / **ダウンロード** で確認・出力
6. **研修比較**タブで前回・今回の Word アンケートを比較

### 設問テンプレート（事前登録）

サイドバー「📋 設問テンプレート」で:

```
Q1: 本講座で印象に残ったことは？
Q2: 改善してほしい点は？
```

の形式で設問を登録し、OCR 時に参照させます。処理済みファイルから設問を取り込むこともできます。

### 前回研修との比較

**📈 研修比較**タブで:

- 前回研修のアンケート Word（.docx）
- 今回研修のアンケート Word（.docx）

をアップロードし、比較レポートを生成します。精度向上のため、「質問ごとにまとめた」後の Word 出力を使うことを推奨します。

## Streamlit Cloud でデプロイ

1. GitHub にリポジトリをプッシュ
2. [Streamlit Cloud](https://streamlit.io/cloud) で New app
3. リポジトリと `app.py` を指定
4. Secrets に `GEMINI_API_KEY` を設定
5. Deploy

## プロジェクト構成

```
.
├── app.py                      # メイン UI（Streamlit）
├── ocr_processor.py            # OCR・複数ページ集約・モデル設定
├── text_editor.py              # テキスト加工・校正・分析 API
├── survey_analyzer.py          # 回答者ビュー・定量サマリー・パーサ
├── question_template_manager.py # 設問テンプレート（事前登録）
├── training_comparison.py      # 前回 vs 今回研修比較
├── word_parser.py              # Word テキスト抽出
├── document_generator.py       # Word 出力
├── excel_exporter.py           # Excel 出力
├── diff_viewer.py              # 校正差分 HTML
├── cost_tracker.py             # API コスト計測
├── gemini_api.py               # Gemini API 共通エラー処理
├── template_manager.py         # 除外テキストテンプレート
├── utils.py                    # ファイル検証など
├── requirements.txt
├── pyproject.toml
└── .streamlit/
    ├── config.toml
    └── secrets.toml.example
```

## 使用モデル

| モデル | API 名 | 用途 |
|--------|--------|------|
| Gemini 3.1 Flash-Lite（デフォルト） | `gemini-3.1-flash-lite` | 費用対効果重視 |
| Gemini 3.5 Flash | `gemini-3.5-flash` | 高精度が必要な場合 |

モデル一覧: [Gemini API 公式ドキュメント](https://ai.google.dev/gemini-api/docs/models?hl=ja)

## 依存パッケージ

- streamlit, pillow, pymupdf, python-docx, openpyxl, numpy, requests, altair, pandas

## セキュリティ

- API キーは `.streamlit/secrets.toml` に保存し、Git にコミットしないでください
- `.gitignore` で `secrets.toml` が除外されていることを確認してください

## ライセンス

MIT License

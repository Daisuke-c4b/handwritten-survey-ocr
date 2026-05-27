# 手書きアンケート OCR・文字起こしアプリ

手書きのアンケート（PDF）を Google Gemini AI で OCR 処理し、テキストを抽出して Word ファイルに変換する Streamlit アプリケーションです。

## 機能

- 📄 PDF ファイルの複数アップロード対応
- 🤖 Gemini AI による高精度な手書き文字認識
- 📝 テキストの自動抽出と整形
- 📥 Word ファイル（.docx）での出力
- 🎨 使いやすい Web インターフェース

## デモ

Streamlit Cloud でホストされています：
[デプロイ後に URL を追加]

## 必要要件

- Python 3.11 以上
- Google Gemini API Key

## インストール

### 依存関係のインストール

```bash
pip install -r requirements.txt
```

または、uv を使用する場合：

```bash
uv sync
```

## 使い方

### ローカルで実行

1. Gemini API Key を設定します：

`.streamlit/secrets.toml`ファイルを作成し、以下を追加：

```toml
GEMINI_API_KEY = "your-api-key-here"
```

または、環境変数として設定：

```bash
export GEMINI_API_KEY="your-api-key-here"
```

2. アプリを起動：

```bash
streamlit run app.py
```

または、uv を使用する場合：

```bash
uv run streamlit run app.py
```

3. ブラウザで `http://localhost:8501` にアクセス

### Streamlit Cloud でデプロイ

1. GitHub にリポジトリをプッシュ
2. [Streamlit Cloud](https://streamlit.io/cloud)にログイン
3. "New app"をクリック
4. リポジトリを選択し、`app.py`を指定
5. "Advanced settings"でシークレット（GEMINI_API_KEY）を設定
6. "Deploy!"をクリック

## プロジェクト構成

```
.
├── app.py                  # メインアプリケーション
├── ocr_processor.py        # OCR処理（Gemini API連携）
├── document_generator.py   # Word文書生成
├── utils.py               # ユーティリティ関数
├── requirements.txt       # 依存パッケージ
├── pyproject.toml        # プロジェクト設定（uv用）
└── .streamlit/
    └── config.toml       # Streamlit設定
```

## 依存パッケージ

- streamlit: Web アプリケーションフレームワーク
- pillow: 画像処理
- pymupdf: PDF ファイル処理
- python-docx: Word 文書生成
- numpy: 数値計算
- requests: HTTP 通信（Gemini API）

## セキュリティ

- API Key は`.streamlit/secrets.toml`に保存し、Git にはコミットしないでください
- `.gitignore`で`secrets.toml`が除外されていることを確認してください

## ライセンス

MIT License

## 作成者

[あなたの名前]

## 謝辞

- Google Gemini API for OCR capabilities
- Streamlit for the web framework


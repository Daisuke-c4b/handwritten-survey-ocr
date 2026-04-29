import streamlit as st
import os
import tempfile
from pathlib import Path

from ocr_processor import (
    OCRProcessor,
    extract_texts_from_screenshots,
    MODEL_NAME,
    MODEL_LABEL,
    MODEL_DESCRIPTION,
    APP_VERSION,
    APP_UPDATED,
)
from document_generator import DocumentGenerator
from utils import (
    validate_pdf,
    validate_image,
    get_file_type,
    extract_filename,
    get_supported_extensions_list,
)

# ---------------------------------------------------------------------------
# Session state helpers
# ---------------------------------------------------------------------------

_STATE_DEFAULTS: dict = {
    "processed_files": [],
    "ocr_processor": None,
    "exclude_texts": "",
    "exclude_screenshots": [],
    "extracted_screenshot_texts": [],
}


def _init_session_state() -> None:
    for key, default in _STATE_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = default


def _invalidate_processor() -> None:
    st.session_state.ocr_processor = None


# ---------------------------------------------------------------------------
# Exclude-text section (tabs)
# ---------------------------------------------------------------------------

def _render_exclude_section() -> None:
    st.subheader("🚫 文字起こし除外設定")
    st.caption(
        "アンケート用紙に印刷されているタイトルや質問文など、"
        "文字起こし不要のテキストを指定できます。"
    )

    tab_text, tab_screenshot = st.tabs(["✏️ テキスト入力", "📸 スクリーンショット"])

    # --- Tab 1: manual text input ---
    with tab_text:
        st.markdown("除外したいテキストを **1 行に 1 つずつ** 入力してください。")
        exclude_input = st.text_area(
            "除外テキスト",
            value=st.session_state.exclude_texts,
            height=140,
            placeholder=(
                "例:\n"
                "顧客満足度アンケート\n"
                "Q1. 本日のセミナーの満足度を教えてください\n"
                "Q2. 理解度はいかがでしたか"
            ),
            label_visibility="collapsed",
        )
        if exclude_input != st.session_state.exclude_texts:
            st.session_state.exclude_texts = exclude_input
            _invalidate_processor()

    # --- Tab 2: screenshot upload ---
    with tab_screenshot:
        st.markdown(
            "除外したい箇所（タイトル・質問文など）の **スクリーンショット** をアップロードしてください。  \n"
            "複数画像を一度にアップロードできます。画像から印刷テキストを自動で読み取ります。"
        )
        screenshots = st.file_uploader(
            "スクリーンショットをアップロード",
            type=["png", "jpg", "jpeg", "gif", "webp", "bmp", "tiff"],
            accept_multiple_files=True,
            key="exclude_screenshot_uploader",
            label_visibility="collapsed",
        )

        if screenshots:
            cols = st.columns(min(len(screenshots), 4))
            for i, shot in enumerate(screenshots):
                with cols[i % len(cols)]:
                    st.image(shot, caption=shot.name, use_container_width=True)

            if st.button("🔍 テキストを読み取る", key="extract_btn", use_container_width=True):
                with st.spinner("スクリーンショットからテキストを抽出中..."):
                    img_bytes_list = [s.getvalue() for s in screenshots]
                    extracted = extract_texts_from_screenshots(img_bytes_list)
                    st.session_state.extracted_screenshot_texts = extracted
                    _invalidate_processor()

        if st.session_state.extracted_screenshot_texts:
            st.success(
                f"{len(st.session_state.extracted_screenshot_texts)} 件のテキストを読み取りました"
            )
            for t in st.session_state.extracted_screenshot_texts:
                st.markdown(f"- {t}")

            if st.button("🗑️ 読み取り結果をクリア", key="clear_extracted"):
                st.session_state.extracted_screenshot_texts = []
                _invalidate_processor()
                st.rerun()


# ---------------------------------------------------------------------------
# Collect all exclude texts
# ---------------------------------------------------------------------------

def _collect_exclude_texts() -> list[str]:
    texts: list[str] = []
    for line in st.session_state.exclude_texts.splitlines():
        line = line.strip()
        if line:
            texts.append(line)
    texts.extend(st.session_state.extracted_screenshot_texts)
    return texts


# ---------------------------------------------------------------------------
# File processing
# ---------------------------------------------------------------------------

def _process_files(uploaded_files: list) -> None:
    """Process uploaded PDF/image files through OCR and generate Word documents."""
    ocr_processor = st.session_state.ocr_processor
    if ocr_processor is None:
        try:
            ocr_processor = OCRProcessor(exclude_texts=_collect_exclude_texts())
            st.session_state.ocr_processor = ocr_processor
        except Exception as e:
            st.error(f"Gemini API の初期化に失敗しました: {str(e)}")
            st.error(
                "GEMINI_API_KEY が正しく設定されているか、"
                "または .streamlit/secrets.toml に設定されているか確認してください。"
            )
            return
    doc_generator = DocumentGenerator()

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files):
        try:
            progress_bar.progress((i + 1) / len(uploaded_files))
            status_text.text(
                f"処理中: {uploaded_file.name} ({i + 1}/{len(uploaded_files)})"
            )

            file_type = get_file_type(uploaded_file.name)

            if file_type == "pdf":
                if not validate_pdf(uploaded_file):
                    st.error(f"❌ {uploaded_file.name}: 有効な PDF ファイルではありません")
                    continue
            elif file_type == "image":
                if not validate_image(uploaded_file):
                    st.error(f"❌ {uploaded_file.name}: 有効な画像ファイルではありません")
                    continue
            else:
                st.error(f"❌ {uploaded_file.name}: サポートされていないファイル形式です")
                continue

            file_ext = Path(uploaded_file.name).suffix.lower()

            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name

            try:
                with st.spinner(f"Gemini AI で文字起こし中: {uploaded_file.name}"):
                    if file_type == "pdf":
                        transcription = ocr_processor.process_pdf(tmp_path)
                    else:
                        transcription = ocr_processor.process_image(tmp_path)

                if not transcription or transcription.strip() == "":
                    st.warning(f"⚠️ {uploaded_file.name}: 文字起こし結果が空です")
                    continue

                with st.spinner(f"Word ファイル生成中: {uploaded_file.name}"):
                    word_content = doc_generator.create_document(
                        transcription, uploaded_file.name
                    )

                st.session_state.processed_files.append(
                    {
                        "filename": uploaded_file.name,
                        "transcription": transcription,
                        "word_content": word_content,
                    }
                )
                st.success(f"✅ {uploaded_file.name}: 処理完了")

            finally:
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

        except Exception as e:
            st.error(f"❌ {uploaded_file.name}: エラーが発生しました - {str(e)}")

    progress_bar.progress(1.0)
    status_text.text("すべての処理が完了しました！")
    st.rerun()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="手書きアンケートOCR",
        page_icon="📝",
        layout="centered",
    )

    _init_session_state()

    # ---- Header ----
    st.title("📝 手書きアンケート OCR・文字起こしアプリ")
    st.markdown(
        "手書きのアンケート（PDF・画像）をアップロードし、"
        "**Gemini AI** で文字起こしして **Word ファイル**としてダウンロードできます。"
    )

    col_model, col_ver = st.columns([3, 1])
    with col_model:
        with st.expander("ℹ️ 使用モデル情報", expanded=False):
            st.markdown(
                f"**{MODEL_LABEL}** &nbsp; `{MODEL_NAME}`  \n"
                f"{MODEL_DESCRIPTION}"
            )
    with col_ver:
        st.markdown(
            f"<div style='text-align:right; color:gray; font-size:0.85em;'>"
            f"Ver. {APP_VERSION}<br>{APP_UPDATED} 更新</div>",
            unsafe_allow_html=True,
        )

    st.divider()

    # ---- Exclude section ----
    _render_exclude_section()

    st.divider()

    # ---- File upload ----
    st.subheader("📁 ファイルのアップロード")

    supported_types = [ext.lstrip(".") for ext in get_supported_extensions_list()]

    uploaded_files = st.file_uploader(
        "手書きアンケートのファイルを選択してください（複数選択可能）",
        type=supported_types,
        accept_multiple_files=True,
        help="対応形式: PDF, PNG, JPG, JPEG, GIF, WebP, BMP, TIFF",
    )

    if uploaded_files:
        st.success(f"{len(uploaded_files)} 個のファイルがアップロードされました")

        with st.expander("アップロードされたファイル一覧", expanded=True):
            for i, f in enumerate(uploaded_files):
                st.write(f"{i + 1}. **{f.name}** ({f.size:,} bytes)")

        if st.button("📝 文字起こしを開始", type="primary", use_container_width=True):
            _process_files(uploaded_files)

    # ---- Results ----
    if st.session_state.processed_files:
        st.divider()
        st.subheader("📄 処理結果")

        for idx, pf in enumerate(st.session_state.processed_files):
            with st.expander(f"📋 {pf['filename']}", expanded=False):
                st.text_area(
                    "文字起こし結果:",
                    pf["transcription"],
                    height=200,
                    disabled=True,
                    key=f"transcription_text_{idx}",
                )
                st.download_button(
                    label="📥 Word ファイルをダウンロード",
                    data=pf["word_content"],
                    file_name=f"{extract_filename(pf['filename'])}_transcription.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        if st.button("🗑️ すべてクリア", use_container_width=True):
            st.session_state.processed_files = []
            st.rerun()


if __name__ == "__main__":
    main()

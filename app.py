import streamlit as st
import os
import tempfile
from io import BytesIO
import base64
from ocr_processor import OCRProcessor, MODEL_NAME, MODEL_LABEL, MODEL_DESCRIPTION
from document_generator import DocumentGenerator
from utils import validate_pdf, validate_image, validate_file, get_file_type, extract_filename, get_supported_extensions_list


def main():
    st.set_page_config(page_title="手書きアンケートOCR", page_icon="📝", layout="centered")

    st.title("📝 手書きアンケート OCR・文字起こしアプリ")
    st.markdown(
        "手書きのアンケート（PDF・画像）をアップロードし、"
        "**Gemini AI** で文字起こしして **Word ファイル**としてダウンロードできます。"
    )

    with st.expander("ℹ️ 使用モデル情報", expanded=False):
        st.markdown(
            f"**{MODEL_LABEL}** &nbsp; `{MODEL_NAME}`  \n"
            f"{MODEL_DESCRIPTION}"
        )

    st.divider()

    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = []
    if 'ocr_processor' not in st.session_state:
        st.session_state.ocr_processor = None
    if 'exclude_texts' not in st.session_state:
        st.session_state.exclude_texts = ""

    # --- Exclude texts ---
    st.subheader("🚫 文字起こし除外テキスト")
    st.caption(
        "アンケート用紙に印刷されているタイトルや質問文など、"
        "文字起こし不要のテキストがあれば入力してください。"
        "該当する文字列は OCR 結果から除外されます。"
    )

    exclude_input = st.text_area(
        "除外テキスト（1 行に 1 つずつ入力）",
        value=st.session_state.exclude_texts,
        height=120,
        placeholder="例:\n顧客満足度アンケート\nQ1. 本日のセミナーの満足度を教えてください\nQ2. 理解度はいかがでしたか",
    )

    if exclude_input != st.session_state.exclude_texts:
        st.session_state.exclude_texts = exclude_input
        st.session_state.ocr_processor = None

    st.divider()

    # --- File upload ---
    st.subheader("📁 ファイルのアップロード")

    supported_types = [ext.lstrip('.') for ext in get_supported_extensions_list()]

    uploaded_files = st.file_uploader(
        "手書きアンケートのファイルを選択してください（複数選択可能）",
        type=supported_types,
        accept_multiple_files=True,
        help="対応形式: PDF, PNG, JPG, JPEG, GIF, WebP, BMP, TIFF",
    )

    if uploaded_files:
        st.success(f"{len(uploaded_files)} 個のファイルがアップロードされました")

        with st.expander("アップロードされたファイル一覧", expanded=True):
            for i, file in enumerate(uploaded_files):
                st.write(f"{i+1}. **{file.name}** ({file.size:,} bytes)")

        if st.button("📝 文字起こしを開始", type="primary", use_container_width=True):
            process_files(uploaded_files)

    # --- Results ---
    if st.session_state.processed_files:
        st.divider()
        st.subheader("📄 処理結果")

        for idx, processed_file in enumerate(st.session_state.processed_files):
            with st.expander(f"📋 {processed_file['filename']}", expanded=False):
                st.text_area(
                    "文字起こし結果:",
                    processed_file['transcription'],
                    height=200,
                    disabled=True,
                    key=f"transcription_text_{idx}",
                )

                st.download_button(
                    label="📥 Word ファイルをダウンロード",
                    data=processed_file['word_content'],
                    file_name=f"{extract_filename(processed_file['filename'])}_transcription.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        if st.button("🗑️ すべてクリア", use_container_width=True):
            st.session_state.processed_files = []
            st.rerun()


def process_files(uploaded_files):
    """Process uploaded PDF/image files through OCR and generate Word documents"""
    ocr_processor = st.session_state.ocr_processor
    if ocr_processor is None:
        try:
            exclude_list = [
                line for line in st.session_state.exclude_texts.splitlines()
                if line.strip()
            ]
            ocr_processor = OCRProcessor(exclude_texts=exclude_list)
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
            progress = (i + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            status_text.text(f"処理中: {uploaded_file.name} ({i+1}/{len(uploaded_files)})")

            file_type = get_file_type(uploaded_file.name)

            if file_type == 'pdf':
                if not validate_pdf(uploaded_file):
                    st.error(f"❌ {uploaded_file.name}: 有効な PDF ファイルではありません")
                    continue
            elif file_type == 'image':
                if not validate_image(uploaded_file):
                    st.error(f"❌ {uploaded_file.name}: 有効な画像ファイルではありません")
                    continue
            else:
                st.error(f"❌ {uploaded_file.name}: サポートされていないファイル形式です")
                continue

            from pathlib import Path
            file_ext = Path(uploaded_file.name).suffix.lower()

            with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name

            try:
                with st.spinner(f"Gemini AI で文字起こし中: {uploaded_file.name}"):
                    if file_type == 'pdf':
                        transcription = ocr_processor.process_pdf(tmp_file_path)
                    else:
                        transcription = ocr_processor.process_image(tmp_file_path)

                if not transcription or transcription.strip() == "":
                    st.warning(f"⚠️ {uploaded_file.name}: 文字起こし結果が空です")
                    continue

                with st.spinner(f"Word ファイル生成中: {uploaded_file.name}"):
                    word_content = doc_generator.create_document(transcription, uploaded_file.name)

                st.session_state.processed_files.append({
                    'filename': uploaded_file.name,
                    'transcription': transcription,
                    'word_content': word_content,
                })

                st.success(f"✅ {uploaded_file.name}: 処理完了")

            finally:
                if os.path.exists(tmp_file_path):
                    os.unlink(tmp_file_path)

        except Exception as e:
            st.error(f"❌ {uploaded_file.name}: エラーが発生しました - {str(e)}")

    progress_bar.progress(1.0)
    status_text.text("すべての処理が完了しました！")

    st.rerun()


if __name__ == "__main__":
    main()

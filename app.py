import streamlit as st
import os
import tempfile
from io import BytesIO
import base64
from ocr_processor import OCRProcessor
from document_generator import DocumentGenerator
from utils import validate_pdf, extract_filename


def main():
    st.title("手書きアンケートOCR・文字起こしアプリ")
    st.markdown(
        "手書きのアンケート（PDF）をアップロードして、Gemini AIで文字起こしし、Wordファイルでダウンロードできます。")

    # Initialize session state
    if 'processed_files' not in st.session_state:
        st.session_state.processed_files = []
    if 'ocr_processor' not in st.session_state:
        # Lazy initialization: create only when needed
        st.session_state.ocr_processor = None

    # File upload section
    st.header("📁 PDFファイルのアップロード")
    uploaded_files = st.file_uploader("手書きアンケートのPDFファイルを選択してください（複数選択可能）",
                                      type=['pdf'],
                                      accept_multiple_files=True,
                                      help="複数のPDFファイルを一度にアップロードできます")

    if uploaded_files:
        st.success(f"{len(uploaded_files)}個のファイルがアップロードされました")

        # Display uploaded files
        st.subheader("アップロードされたファイル:")
        for i, file in enumerate(uploaded_files):
            st.write(f"{i+1}. {file.name} ({file.size:,} bytes)")

        # Process files button
        if st.button("📝 文字起こしを開始", type="primary"):
            process_files(uploaded_files)

    # Display processed files
    if st.session_state.processed_files:
        st.header("📄 処理完了したファイル")

        for idx, processed_file in enumerate(st.session_state.processed_files):
            with st.expander(f"📋 {processed_file['filename']}",
                             expanded=False):
                st.text_area("文字起こし結果:",
                             processed_file['transcription'],
                             height=200,
                             disabled=True,
                             key=f"transcription_text_{idx}")

                # Download button
                if st.download_button(
                        label="📥 Wordファイルをダウンロード",
                        data=processed_file['word_content'],
                        file_name=
                        f"{extract_filename(processed_file['filename'])}_transcription.docx",
                        mime=
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                ):
                    st.success("ダウンロードが開始されました！")

        # Clear all button
        if st.button("🗑️ すべてクリア"):
            st.session_state.processed_files = []
            st.rerun()


def process_files(uploaded_files):
    """Process uploaded PDF files through OCR and generate Word documents"""
    # Ensure OCRProcessor is available (lazy init)
    ocr_processor = st.session_state.ocr_processor
    if ocr_processor is None:
        try:
            ocr_processor = OCRProcessor()
            st.session_state.ocr_processor = ocr_processor
        except Exception as e:
            st.error(f"Gemini APIの初期化に失敗しました: {str(e)}")
            st.error("GEMINI_API_KEYが正しく設定されているか、または .streamlit/secrets.toml に設定されているか確認してください。")
            return
    doc_generator = DocumentGenerator()

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files):
        try:
            # Update progress
            progress = (i + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            status_text.text(
                f"処理中: {uploaded_file.name} ({i+1}/{len(uploaded_files)})")

            # Validate PDF
            if not validate_pdf(uploaded_file):
                st.error(f"❌ {uploaded_file.name}: 有効なPDFファイルではありません")
                continue

            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False,
                                             suffix='.pdf') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name

            try:
                # Process OCR
                with st.spinner(f"Gemini AIで文字起こし中: {uploaded_file.name}"):
                    transcription = ocr_processor.process_pdf(tmp_file_path)

                if not transcription or transcription.strip() == "":
                    st.warning(f"⚠️ {uploaded_file.name}: 文字起こし結果が空です")
                    continue

                # Generate Word document
                with st.spinner(f"Wordファイル生成中: {uploaded_file.name}"):
                    word_content = doc_generator.create_document(
                        transcription, uploaded_file.name)

                # Store processed file
                st.session_state.processed_files.append({
                    'filename':
                    uploaded_file.name,
                    'transcription':
                    transcription,
                    'word_content':
                    word_content
                })

                st.success(f"✅ {uploaded_file.name}: 処理完了")

            finally:
                # Clean up temporary file
                if os.path.exists(tmp_file_path):
                    os.unlink(tmp_file_path)

        except Exception as e:
            st.error(f"❌ {uploaded_file.name}: エラーが発生しました - {str(e)}")

    progress_bar.progress(1.0)
    status_text.text("すべての処理が完了しました！")

    # Auto-rerun to show results
    st.rerun()


if __name__ == "__main__":
    main()

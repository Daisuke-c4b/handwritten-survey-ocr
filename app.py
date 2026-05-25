import streamlit as st
import os
import io
import re
import tempfile
from pathlib import Path

import altair as alt
import pandas as pd

from ocr_processor import (
    OCRProcessor,
    extract_texts_from_screenshots,
    MODEL_NAME,
    MODEL_LABEL,
    MODEL_DESCRIPTION,
    APP_VERSION,
    APP_UPDATED,
    CHANGELOG,
)
from document_generator import DocumentGenerator
from template_manager import (
    list_templates,
    get_template,
    add_template,
    update_template,
    delete_template,
)
from text_editor import TextEditor, EDITING_MODES
from utils import (
    validate_pdf,
    validate_image,
    get_file_type,
    extract_filename,
    get_supported_extensions_list,
)
import cost_tracker
import survey_analyzer
import excel_exporter
from diff_viewer import generate_diff_html

# ---------------------------------------------------------------------------
# Theme / visual constants
# ---------------------------------------------------------------------------

_COLOR_POSITIVE = "#22c55e"
_COLOR_NEGATIVE = "#ef4444"
_COLOR_NEUTRAL = "#94a3b8"
_COLOR_PRIMARY = "#3b82f6"
_COLOR_ACCENT = "#10b981"
_COLOR_WARNING = "#f59e0b"

_GLOBAL_CSS = """
<style>
/* メインコンテナをビューポート幅いっぱいまで広げる */
.block-container {
    padding-top: 1.6rem;
    padding-bottom: 4rem;
    padding-left: 2.2rem;
    padding-right: 2.2rem;
    max-width: none;
}

/* サイドバー: 背景と内側パディングのみ調整し、幅・折りたたみは Streamlit の挙動に完全に任せる */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f8fafc 0%, #ffffff 100%);
}
section[data-testid="stSidebar"] .block-container {
    padding: 1.5rem 1rem 2rem 1rem;
}
/* タブを少し大きめに */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 2px solid #e5e7eb;
}
.stTabs [data-baseweb="tab"] {
    height: 44px;
    padding: 0 18px;
    font-weight: 500;
    border-radius: 8px 8px 0 0;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(180deg, #eff6ff 0%, #ffffff 100%);
    color: #1d4ed8;
    border-bottom: 3px solid #3b82f6;
}
/* エキスパンダ */
[data-testid="stExpander"] {
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    background: #ffffff;
}
/* メトリクス */
[data-testid="stMetric"] {
    background: linear-gradient(135deg, #f8fafc 0%, #ffffff 100%);
    padding: 14px 18px;
    border-radius: 10px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04);
}
/* ボタンの強調 */
.stButton button[kind="primary"] {
    background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
    border: none;
    box-shadow: 0 2px 4px rgba(59,130,246,0.25);
}
.stButton button[kind="primary"]:hover {
    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
}
/* ファイル結果のヘッダ */
.file-card-header {
    background: linear-gradient(135deg, #f8fafc 0%, #ffffff 100%);
    padding: 12px 16px;
    border-radius: 8px;
    border-left: 4px solid #3b82f6;
    margin-bottom: 12px;
    font-weight: 600;
}
/* メトリクスカード（カスタム） */
.metric-card {
    padding: 16px 18px;
    border-radius: 10px;
    border: 1px solid #e5e7eb;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    background: #ffffff;
}
.metric-card .label {
    font-size: 0.8rem;
    color: #64748b;
    margin-bottom: 4px;
    text-transform: uppercase;
    letter-spacing: 0.02em;
}
.metric-card .value {
    font-size: 1.9rem;
    font-weight: 700;
    line-height: 1.1;
}
.metric-card .sub {
    font-size: 0.85rem;
    color: #94a3b8;
    margin-top: 4px;
}
</style>
"""


def _inject_global_css() -> None:
    st.markdown(_GLOBAL_CSS, unsafe_allow_html=True)


def _metric_card(label: str, value, sub: str = "", color: str = _COLOR_PRIMARY) -> str:
    """カラーアクセント付きのメトリクスカード HTML を返す."""
    return (
        f'<div class="metric-card" style="border-left: 4px solid {color};">'
        f'<div class="label">{label}</div>'
        f'<div class="value" style="color: {color};">{value}</div>'
        f'<div class="sub">{sub}</div>'
        f'</div>'
    )


# ---------------------------------------------------------------------------
# Respondent identifier helpers
# ---------------------------------------------------------------------------

_PN_PATTERN = re.compile(r"\[(?:P\.\d+|回答者\d+|F\.[^\]]+-P\.\d+)\]")


def has_respondent_id(text: str) -> bool:
    """テキストに [P.N] 等の回答者識別子が含まれているか."""
    if not text:
        return False
    return bool(_PN_PATTERN.search(text))


def _apply_matome(idx: int, pf: dict, rerun: bool = True) -> bool:
    """「質問ごとにまとめる」を pf['current_text'] に適用する."""
    try:
        editor = TextEditor()
        with st.spinner("質問ごとにまとめています..."):
            result = editor.apply_editing(pf["current_text"], "matome", "")
        pf["current_text"] = result
        ver_key = f"text_ver_{idx}"
        st.session_state[ver_key] = st.session_state.get(ver_key, 0) + 1
        if rerun:
            st.rerun()
        return True
    except Exception as e:
        st.error(f"質問ごとにまとめる処理でエラーが発生しました: {str(e)}")
        return False


def _render_inline_matome_cta(
    idx: int,
    pf: dict,
    message: str,
    button_key_suffix: str,
) -> None:
    """各タブ内で「質問ごとにまとめる」の実行を促すインラインバナー（ボタンなし）.

    実行ボタンは「✏️ 編集・加工」タブのステップ 1 に一本化している。
    """
    st.markdown(
        f'<div style="background:{_COLOR_WARNING}11;border:1px solid {_COLOR_WARNING}55;'
        f'border-left:4px solid {_COLOR_WARNING};padding:12px 16px;border-radius:8px;'
        f'margin:8px 0 12px 0;">'
        f'<div style="font-weight:600;color:#b45309;margin-bottom:4px;">'
        f'🎯 まず「質問ごとにまとめる」を実行してください</div>'
        f'<div style="color:#475569;font-size:0.92rem;line-height:1.55;">{message}'
        f'<br><strong>「✏️ 編集・加工」タブのステップ 1</strong> から実行できます。</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def _render_matome_banner(idx: int, pf: dict) -> None:
    """ファイルカード直下に状態バナーを表示する（ボタンは置かない）.

    実行ボタンは「✏️ 編集・加工」タブのステップ1に一本化している。
    """
    if has_respondent_id(pf["current_text"]):
        st.markdown(
            f'<div style="background:{_COLOR_POSITIVE}11;border:1px solid {_COLOR_POSITIVE}55;'
            f'border-left:4px solid {_COLOR_POSITIVE};padding:10px 16px;border-radius:8px;'
            f'margin-bottom:14px;">'
            f'<strong style="color:#15803d;">✅ 回答識別子が付与されています</strong>'
            f'<span style="color:#475569;margin-left:10px;font-size:0.9rem;">'
            f'回答者ビュー / 定量分析 / Excel 出力など全ての機能をご利用いただけます。</span>'
            f'</div>',
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f'<div style="background:{_COLOR_WARNING}11;border:1px solid {_COLOR_WARNING}55;'
        f'border-left:4px solid {_COLOR_WARNING};padding:14px 18px;border-radius:8px;'
        f'margin-bottom:14px;">'
        f'<div style="font-weight:600;color:#b45309;font-size:1.02rem;margin-bottom:4px;">'
        f'🎯 まだ「質問ごとにまとめる」が実行されていません</div>'
        f'<div style="color:#475569;font-size:0.92rem;line-height:1.55;">'
        f'回答者単位ビュー・定量分析・Excel 出力など多くの機能で <strong>[P.N] 形式の回答者識別子</strong>が必要です。'
        f'下の <strong>「✏️ 編集・加工」タブ → ステップ 1</strong> から実行してください。'
        f'</div></div>',
        unsafe_allow_html=True,
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
    "ocr_mode": "accurate",
    "survey_analysis": "",
    "proofread_input": "",
    "proofread_result": "",
    "proofread_fixed_text": "",
    "auto_template_candidates": [],
}


def _init_session_state() -> None:
    for key, default in _STATE_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = default


def _invalidate_processor() -> None:
    st.session_state.ocr_processor = None


# ---------------------------------------------------------------------------
# OCR mode section
# ---------------------------------------------------------------------------

def _render_ocr_mode_section() -> None:
    st.caption("画像の読み取り方を選択します。")

    mode = st.radio(
        "文字起こしモード",
        options=["accurate", "proofread"],
        format_func=lambda x: {
            "accurate": "✍️ 正確文字起こし",
            "proofread": "🔍 校正文字起こし",
        }[x],
        index=0 if st.session_state.ocr_mode == "accurate" else 1,
        horizontal=False,
        key="ocr_mode_radio",
        label_visibility="collapsed",
    )

    if mode != st.session_state.ocr_mode:
        st.session_state.ocr_mode = mode
        _invalidate_processor()

    if mode == "accurate":
        st.caption("書かれている文字をそのまま一字一句正確に転写します。")
    else:
        st.caption(
            "二重線の訂正・丸印による選択など校正箇所を検出し、"
            "不自然な表現も指摘します。結果に【校正注記】が付加されます。"
        )


# ---------------------------------------------------------------------------
# Exclude-text section (tabs)
# ---------------------------------------------------------------------------

def _render_exclude_section() -> None:
    st.caption(
        "アンケート用紙に印刷されているタイトルや質問文など、"
        "文字起こし不要のテキストを指定できます。"
    )

    tab_text, tab_screenshot, tab_template = st.tabs(
        ["✏️ テキスト入力", "📸 スクリーンショット", "📋 テンプレート"]
    )

    # --- Tab 1: manual text input ---
    with tab_text:
        st.markdown("除外したいテキストを **1 行に 1 つずつ** 入力してください。")
        exclude_input = st.text_area(
            "除外テキスト",
            value=st.session_state.exclude_texts,
            height=140,
            placeholder=(
                "例:\n"
                "1.本講座の中で、興味を持った・参考になった等、感想や心に残ったことをお聞かせください\n"
                "2.その他、ご意見、ご要望等お気づきの点がありましたらご記入ください"
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

    # --- Tab 3: template management ---
    with tab_template:
        _render_template_tab()


# ---------------------------------------------------------------------------
# Template tab
# ---------------------------------------------------------------------------

def _render_template_tab() -> None:
    templates = list_templates()

    # ---- Load template ----
    if templates:
        st.markdown("##### テンプレートを読み込む")
        tmpl_names = [t["name"] for t in templates]
        selected = st.selectbox(
            "テンプレートを選択",
            options=tmpl_names,
            index=0,
            key="tmpl_select",
            label_visibility="collapsed",
        )
        col_load, col_del = st.columns(2)
        with col_load:
            if st.button("📥 読み込む", key="tmpl_load", use_container_width=True):
                tmpl = get_template(selected)
                if tmpl:
                    st.session_state.exclude_texts = "\n".join(tmpl["texts"])
                    _invalidate_processor()
                    st.rerun()
        with col_del:
            if st.button("🗑️ 削除", key="tmpl_delete", use_container_width=True):
                delete_template(selected)
                st.success(f"テンプレート「{selected}」を削除しました")
                st.rerun()
    else:
        st.info("保存済みテンプレートはありません。")

    st.divider()

    # ---- Save current as template ----
    st.markdown("##### 現在の除外テキストをテンプレートとして保存")
    new_name = st.text_input(
        "テンプレート名",
        key="tmpl_new_name",
        placeholder="例: 研修アンケート用",
    )
    if st.button("💾 保存", key="tmpl_save", use_container_width=True):
        if not new_name.strip():
            st.warning("テンプレート名を入力してください。")
        else:
            current_texts = [
                line.strip()
                for line in st.session_state.exclude_texts.splitlines()
                if line.strip()
            ]
            current_texts.extend(st.session_state.extracted_screenshot_texts)
            if not current_texts:
                st.warning("除外テキストが空のため保存できません。")
            else:
                try:
                    add_template(new_name.strip(), current_texts)
                    st.success(f"テンプレート「{new_name.strip()}」を保存しました")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))

    # ---- Edit existing template ----
    if templates:
        st.divider()
        st.markdown("##### テンプレートを編集")
        edit_target = st.selectbox(
            "編集するテンプレート",
            options=[t["name"] for t in templates],
            index=0,
            key="tmpl_edit_select",
        )
        tmpl_data = get_template(edit_target)
        edit_name = st.text_input(
            "テンプレート名",
            value=edit_target,
            key="tmpl_edit_name",
        )
        edit_texts = st.text_area(
            "除外テキスト（1 行に 1 つ）",
            value="\n".join(tmpl_data["texts"]) if tmpl_data else "",
            height=140,
            key="tmpl_edit_texts",
        )
        if st.button("✏️ 更新", key="tmpl_update", use_container_width=True):
            if not edit_name.strip():
                st.warning("テンプレート名を入力してください。")
            else:
                new_texts = [l.strip() for l in edit_texts.splitlines() if l.strip()]
                if not new_texts:
                    st.warning("除外テキストが空のため更新できません。")
                else:
                    try:
                        update_template(edit_target, edit_name.strip(), new_texts)
                        st.success(f"テンプレート「{edit_name.strip()}」を更新しました")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))


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
# Standalone proofread (paste-and-check) section
# ---------------------------------------------------------------------------

def _render_proofread_section() -> None:
    """Render the standalone copy-paste proofread / quality-check feature.

    OCR を介さず、ユーザーがアンケートのテキストデータを貼り付けて
    誤字脱字・違和感のある文体・不自然な日本語をチェックできる。
    """
    st.caption(
        "アンケートの文字データを貼り付けて、**誤字脱字・違和感のある文体・"
        "不自然な日本語**をチェックできます。OCR を介さず既存テキストの校正に使えます。"
    )

    proof_input = st.text_area(
        "校正対象のテキスト",
        value=st.session_state.proofread_input,
        height=240,
        key="proofread_input_area",
        placeholder=(
            "ここにアンケートの文字データを貼り付けてください。\n"
            "（例）昨日の研修はとても勉強なりました。とくに事例紹介の部分が良かったです。"
        ),
    )
    if proof_input != st.session_state.proofread_input:
        st.session_state.proofread_input = proof_input

    col_run, col_clear = st.columns([3, 1])
    with col_run:
        run_clicked = st.button(
            "🔍 校正チェックを実行",
            key="proofread_run",
            use_container_width=True,
            type="primary",
        )
    with col_clear:
        clear_clicked = st.button(
            "🗑️ クリア",
            key="proofread_clear",
            use_container_width=True,
        )

    if clear_clicked:
        st.session_state.proofread_input = ""
        st.session_state.proofread_result = ""
        st.session_state.proofread_fixed_text = ""
        st.rerun()

    if run_clicked:
        if not st.session_state.proofread_input.strip():
            st.warning("校正対象のテキストを入力してください。")
        else:
            try:
                editor = TextEditor()
                with st.spinner("テキストを校正中... しばらくお待ちください"):
                    result = editor.check_text_quality(
                        st.session_state.proofread_input
                    )
                    fixed = editor.fix_text_quality(
                        st.session_state.proofread_input
                    )
                st.session_state.proofread_result = result
                st.session_state.proofread_fixed_text = fixed
                st.rerun()
            except Exception as e:
                st.error(f"校正エラー: {str(e)}")

    if st.session_state.proofread_result:
        st.markdown("**校正結果（指摘レポート）**")
        st.markdown(st.session_state.proofread_result)

        if st.session_state.proofread_fixed_text:
            st.divider()
            st.markdown("**🔄 修正前/修正後 の差分プレビュー**")
            st.caption(
                "Gemini が自動修正したテキストと原文の差分を表示します。"
                "緑色が追加・赤色（取り消し線）が削除です。"
            )
            diff_html = generate_diff_html(
                st.session_state.proofread_input,
                st.session_state.proofread_fixed_text,
            )
            st.markdown(diff_html, unsafe_allow_html=True)

            with st.expander("修正後テキスト（コピー用）", expanded=False):
                st.text_area(
                    "修正後テキスト",
                    value=st.session_state.proofread_fixed_text,
                    height=240,
                    key="proofread_fixed_area",
                    label_visibility="collapsed",
                )

        st.divider()
        st.caption("結果テキストをコピーしたい場合は以下のエリアから取得できます。")
        st.text_area(
            "校正結果（コピー用）",
            value=st.session_state.proofread_result,
            height=180,
            key="proofread_result_area",
            label_visibility="collapsed",
        )

        try:
            doc_generator = DocumentGenerator()
            proof_word = doc_generator.create_analysis_document(
                st.session_state.proofread_result
            )
            st.download_button(
                label="📥 校正結果を Word でダウンロード",
                data=proof_word,
                file_name="校正チェック結果.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                use_container_width=True,
                key="proofread_download",
            )
        except Exception as e:
            st.error(f"Word 生成エラー: {str(e)}")


# ---------------------------------------------------------------------------
# File processing
# ---------------------------------------------------------------------------

def _process_files(uploaded_files: list) -> None:
    """Process uploaded PDF/image files through OCR and generate Word documents."""
    ocr_processor = st.session_state.ocr_processor
    if ocr_processor is None:
        try:
            ocr_processor = OCRProcessor(
                exclude_texts=_collect_exclude_texts(),
                ocr_mode=st.session_state.ocr_mode,
            )
            st.session_state.ocr_processor = ocr_processor
        except Exception as e:
            st.error(f"Gemini API の初期化に失敗しました: {str(e)}")
            st.error(
                "GEMINI_API_KEY が正しく設定されているか、"
                "または .streamlit/secrets.toml に設定されているか確認してください。"
            )
            return

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
                        transcription, page_images = (
                            ocr_processor.process_pdf_with_images(tmp_path)
                        )
                    else:
                        transcription, page_images = (
                            ocr_processor.process_image_with_images(tmp_path)
                        )

                if not transcription or transcription.strip() == "":
                    st.warning(f"⚠️ {uploaded_file.name}: 文字起こし結果が空です")
                    continue

                # ページ画像を PNG バイト列にしてセッションへ格納（再実行に耐える）
                page_image_bytes: list[tuple[bytes, int]] = []
                for img, pnum in page_images:
                    try:
                        buf = io.BytesIO()
                        img.save(buf, format="PNG")
                        page_image_bytes.append((buf.getvalue(), pnum))
                    except Exception:
                        continue

                st.session_state.processed_files.append(
                    {
                        "filename": uploaded_file.name,
                        "transcription": transcription,
                        "current_text": transcription,
                        "page_images": page_image_bytes,
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
# Respondent view / Quant summary / Image comparison
# ---------------------------------------------------------------------------

def _render_respondent_view(idx: int, pf: dict) -> None:
    """[P.N] でグルーピングした回答者単位ビュー（カード型・複数カラム）."""
    st.markdown("**👥 回答者単位ビュー**")
    st.caption(
        "[P.N] や [回答者N] の識別子を基準にグルーピングして、回答者ごとに縦読みします。"
        "「質問ごとにまとめる」加工を先に実行しておくと識別子が確実に付与されます。"
    )
    parsed = survey_analyzer.parse_consolidated_text(pf["current_text"])
    respondents = survey_analyzer.to_respondent_view(parsed)
    if not respondents:
        _render_inline_matome_cta(
            idx,
            pf,
            message=(
                "このビューを利用するには、まず回答者識別子の付与が必要です。"
                "「質問ごとにまとめる」を実行すると、各回答に [P.N] が付き、回答者ごとに縦読みできるようになります。"
            ),
            button_key_suffix="resp",
        )
        return

    # サマリーチップ
    total_answers = sum(len(b.answers) for b in respondents)
    st.markdown(
        f'<div style="margin: 8px 0 16px 0;">'
        f'<span style="background:{_COLOR_PRIMARY}15;color:{_COLOR_PRIMARY};'
        f'padding:6px 12px;border-radius:14px;font-size:0.85rem;margin-right:6px;">'
        f'👥 {len(respondents)} 名の回答者</span>'
        f'<span style="background:{_COLOR_ACCENT}15;color:{_COLOR_ACCENT};'
        f'padding:6px 12px;border-radius:14px;font-size:0.85rem;">'
        f'💬 {total_answers} 件の回答</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ワイドレイアウトを活かして 2 列でカード表示
    cols = st.columns(2, gap="large")
    for i, block in enumerate(respondents):
        with cols[i % 2]:
            with st.container(border=True):
                st.markdown(
                    f'<div style="font-weight:600;color:{_COLOR_PRIMARY};margin-bottom:6px;">'
                    f'[{block.respondent_id}]'
                    f'<span style="color:#94a3b8;font-weight:400;font-size:0.85rem;margin-left:8px;">'
                    f'回答数 {len(block.answers)}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
                for q_num, q_text, ans in block.answers:
                    st.markdown(
                        f'<div style="margin: 8px 0 4px 0;">'
                        f'<span style="background:#eff6ff;color:{_COLOR_PRIMARY};'
                        f'padding:2px 8px;border-radius:6px;font-size:0.8rem;font-weight:600;'
                        f'margin-right:6px;">Q{q_num}</span>'
                        f'<span style="font-size:0.85rem;color:#475569;">{q_text}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                    st.markdown(f"&nbsp;&nbsp;&nbsp;&nbsp;{ans}", unsafe_allow_html=True)


def _render_quant_summary(idx: int, pf: dict) -> None:
    """センチメント分類・トピック・キーワードの定量サマリー（altair でスタイリッシュ化）."""
    summary_key = "quant_summary"

    st.markdown("**📊 定量サマリー**")
    st.caption(
        "現在のテキストを Gemini に解析させ、センチメント・トピック・キーワード・"
        "代表回答を JSON で取得し、視覚化します。"
    )

    if not has_respondent_id(pf["current_text"]):
        _render_inline_matome_cta(
            idx,
            pf,
            message=(
                "より精度の高い定量サマリーを得るためには、先に「質問ごとにまとめる」"
                "を実行して回答者識別子を付与しておくことを推奨します（質問単位の集約とセンチメント評価が安定します）。"
            ),
            button_key_suffix="quant",
        )

    col_run, col_clear, _ = st.columns([2, 1, 3])
    with col_run:
        if st.button(
            "📈 定量サマリーを生成",
            key=f"quant_btn_{idx}",
            use_container_width=True,
            type="primary",
        ):
            try:
                editor = TextEditor()
                with st.spinner("定量サマリーを生成中..."):
                    summary = survey_analyzer.generate_quant_summary(
                        pf["current_text"],
                        caller=lambda p: editor.call_with_purpose(p, "定量サマリー"),
                    )
                pf[summary_key] = summary
                st.rerun()
            except Exception as e:
                st.error(f"定量サマリー生成エラー: {str(e)}")
    with col_clear:
        if pf.get(summary_key):
            if st.button("🗑️ クリア", key=f"quant_clear_{idx}", use_container_width=True):
                pf[summary_key] = None
                st.rerun()

    summary = pf.get(summary_key)
    if not summary:
        st.info("「定量サマリーを生成」を押すと、ここに結果が表示されます。")
        return

    if "_raw" in summary:
        st.warning("JSON 解析に失敗しました。元応答を表示します。")
        st.code(summary.get("_raw", ""))
        return

    sent = summary.get("sentiment", {}) or {}
    pos = int(sent.get("positive", 0))
    neg = int(sent.get("negative", 0))
    neu = int(sent.get("neutral", 0))
    total = max(1, pos + neg + neu)
    answer_count = summary.get("answer_count", pos + neg + neu)

    # ---- メトリクスカード ----
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(_metric_card("回答数", f"{answer_count}", color="#0f172a"), unsafe_allow_html=True)
    with c2:
        st.markdown(
            _metric_card("ポジティブ", f"{pos}", sub=f"{pos / total * 100:.0f}%", color=_COLOR_POSITIVE),
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            _metric_card("ネガティブ", f"{neg}", sub=f"{neg / total * 100:.0f}%", color=_COLOR_NEGATIVE),
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            _metric_card("中立", f"{neu}", sub=f"{neu / total * 100:.0f}%", color=_COLOR_NEUTRAL),
            unsafe_allow_html=True,
        )

    st.markdown('<div style="margin-top: 16px;"></div>', unsafe_allow_html=True)

    # ---- 2列構成: 左=ドーナツチャート、右=トピック横棒 ----
    col_sent, col_topic = st.columns([2, 3], gap="large")

    with col_sent:
        st.markdown("**センチメント分布**")
        if pos + neg + neu > 0:
            sent_df = pd.DataFrame(
                [
                    {"label": "ポジティブ", "count": pos},
                    {"label": "ネガティブ", "count": neg},
                    {"label": "中立", "count": neu},
                ]
            )
            sent_chart = (
                alt.Chart(sent_df)
                .mark_arc(innerRadius=58, outerRadius=110, stroke="#fff", strokeWidth=2)
                .encode(
                    theta=alt.Theta("count:Q"),
                    color=alt.Color(
                        "label:N",
                        scale=alt.Scale(
                            domain=["ポジティブ", "ネガティブ", "中立"],
                            range=[_COLOR_POSITIVE, _COLOR_NEGATIVE, _COLOR_NEUTRAL],
                        ),
                        legend=alt.Legend(title=None, orient="bottom", labelFontSize=12),
                    ),
                    tooltip=[
                        alt.Tooltip("label:N", title="分類"),
                        alt.Tooltip("count:Q", title="件数"),
                    ],
                )
                .properties(height=320)
                .configure_view(strokeWidth=0)
            )
            st.altair_chart(sent_chart, use_container_width=True)
        else:
            st.info("センチメント情報がありません。")

    with col_topic:
        topics = summary.get("topics", []) or []
        st.markdown("**トピック（件数の多い順）**")
        if topics:
            topic_df = pd.DataFrame(
                [
                    {
                        "topic": t.get("label", ""),
                        "count": int(t.get("count", 0)),
                        "examples": " / ".join(t.get("examples", []) or []),
                    }
                    for t in topics
                ]
            )
            topic_chart = (
                alt.Chart(topic_df)
                .mark_bar(cornerRadiusEnd=4)
                .encode(
                    x=alt.X("count:Q", title="件数", axis=alt.Axis(grid=True, tickMinStep=1)),
                    y=alt.Y(
                        "topic:N",
                        sort="-x",
                        title=None,
                        axis=alt.Axis(labelLimit=240, labelFontSize=12),
                    ),
                    color=alt.Color(
                        "count:Q",
                        scale=alt.Scale(scheme="blues"),
                        legend=None,
                    ),
                    tooltip=[
                        alt.Tooltip("topic:N", title="トピック"),
                        alt.Tooltip("count:Q", title="件数"),
                        alt.Tooltip("examples:N", title="代表表現"),
                    ],
                )
                .properties(height=max(220, 36 * len(topic_df)))
                .configure_view(strokeWidth=0)
                .configure_axis(labelColor="#475569", titleColor="#334155")
            )
            st.altair_chart(topic_chart, use_container_width=True)

            with st.expander("トピックの代表表現を見る", expanded=False):
                st.dataframe(
                    topic_df.rename(
                        columns={"topic": "トピック", "count": "件数", "examples": "代表表現"}
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
        else:
            st.info("トピックは検出されませんでした。")

    # ---- キーワードと代表回答 ----
    col_kw, col_hl = st.columns([3, 2], gap="large")

    keywords = summary.get("keywords", []) or []
    with col_kw:
        st.markdown("**頻出キーワード**")
        if keywords:
            kw_df = pd.DataFrame(
                [{"word": k.get("word", ""), "count": int(k.get("count", 0))} for k in keywords]
            )
            kw_chart = (
                alt.Chart(kw_df)
                .mark_bar(cornerRadiusEnd=4, color=_COLOR_ACCENT)
                .encode(
                    x=alt.X("count:Q", title="出現数", axis=alt.Axis(tickMinStep=1)),
                    y=alt.Y(
                        "word:N",
                        sort="-x",
                        title=None,
                        axis=alt.Axis(labelLimit=200, labelFontSize=12),
                    ),
                    tooltip=[
                        alt.Tooltip("word:N", title="キーワード"),
                        alt.Tooltip("count:Q", title="出現数"),
                    ],
                )
                .properties(height=max(180, 28 * len(kw_df)))
                .configure_view(strokeWidth=0)
                .configure_axis(labelColor="#475569", titleColor="#334155")
            )
            st.altair_chart(kw_chart, use_container_width=True)
        else:
            st.info("キーワードは検出されませんでした。")

    highlights = summary.get("highlights", {}) or {}
    pos_h = highlights.get("positive", []) or []
    neg_h = highlights.get("negative", []) or []
    with col_hl:
        st.markdown("**代表的な回答**")
        if pos_h:
            st.markdown(
                f'<div style="background:{_COLOR_POSITIVE}11;border-left:4px solid {_COLOR_POSITIVE};'
                'padding:10px 14px;border-radius:6px;margin-bottom:8px;">'
                '<strong style="color:#16a34a;">ポジティブ</strong></div>',
                unsafe_allow_html=True,
            )
            for s in pos_h:
                st.markdown(f"- {s}")
        if neg_h:
            st.markdown(
                f'<div style="background:{_COLOR_NEGATIVE}11;border-left:4px solid {_COLOR_NEGATIVE};'
                'padding:10px 14px;border-radius:6px;margin:12px 0 8px 0;">'
                '<strong style="color:#dc2626;">ネガティブ</strong></div>',
                unsafe_allow_html=True,
            )
            for s in neg_h:
                st.markdown(f"- {s}")
        if not pos_h and not neg_h:
            st.info("代表回答は抽出されませんでした。")


def _render_image_comparison(idx: int, pf: dict) -> None:
    """原画像との対比ビュー（大きく見やすく）."""
    st.markdown("**🖼️ 原画像との対比**")
    page_images = pf.get("page_images") or []
    if not page_images:
        st.info("原画像が保持されていません（旧フォーマットの結果）。再アップロードで利用可能になります。")
        return

    page_options = [pnum for _b, pnum in page_images]

    col_page, col_layout = st.columns([2, 3])
    with col_page:
        selected = st.selectbox(
            "表示するページ",
            options=page_options,
            index=0,
            key=f"img_page_{idx}",
            help="該当ページを選んで原画像と文字起こしテキストを比較できます。",
        )
    with col_layout:
        layout_mode = st.radio(
            "表示レイアウト",
            options=["横並び（対比）", "縦並び（画像を大きく表示）"],
            horizontal=True,
            key=f"img_layout_{idx}",
            label_visibility="collapsed",
        )

    img_bytes = None
    for b, pnum in page_images:
        if pnum == selected:
            img_bytes = b
            break

    if img_bytes is None:
        st.warning("画像が見つかりませんでした。")
        return

    snippet = _extract_page_snippet(pf["current_text"], selected)

    if layout_mode == "横並び（対比）":
        # ワイドレイアウトでは画像側を広めに（6:4）し、画像をしっかり大きく表示
        col_img, col_txt = st.columns([6, 4], gap="large")
        with col_img:
            st.markdown(f"**📄 ページ {selected} の原画像**")
            st.image(img_bytes, use_container_width=True)
        with col_txt:
            st.markdown(f"**✏️ ページ {selected} の文字起こし抜粋**")
            st.caption(f"識別子 [P.{selected}] を含む行を抽出しています。")
            st.text_area(
                f"ページ {selected} 抜粋",
                value=snippet,
                height=620,
                key=f"img_text_{idx}_{selected}",
                label_visibility="collapsed",
            )
    else:
        st.markdown(f"**📄 ページ {selected} の原画像**")
        # 中央寄せの大きな画像
        col_lp, col_img, col_rp = st.columns([1, 6, 1])
        with col_img:
            st.image(img_bytes, use_container_width=True)
        st.markdown(f"**✏️ ページ {selected} の文字起こし抜粋**")
        st.caption(f"識別子 [P.{selected}] を含む行を抽出しています。")
        st.text_area(
            f"ページ {selected} 抜粋",
            value=snippet,
            height=320,
            key=f"img_text_{idx}_{selected}",
            label_visibility="collapsed",
        )

    # ダウンロードボタン（高解像度画像をそのまま）
    st.download_button(
        label="🖼️ このページの原画像をダウンロード",
        data=img_bytes,
        file_name=f"{extract_filename(pf['filename'])}_page{selected}.png",
        mime="image/png",
        use_container_width=False,
        key=f"img_dl_{idx}_{selected}",
    )


def _extract_page_snippet(text: str, page_num: int) -> str:
    """集約済みテキストから [P.<page_num>] を含む行と質問見出しを抜粋."""
    if not text:
        return ""
    needle = f"[P.{page_num}]"
    lines = text.splitlines()
    out: list[str] = []
    current_q: str = ""
    last_q_emitted = ""
    for line in lines:
        s = line.strip()
        if not s:
            continue
        if s.startswith("Q") and ":" in s.split(" ", 1)[0] + ":":
            current_q = s
            continue
        # Q ヘッダ簡易判定（質問/Q）
        if s.startswith("質問") or (s.startswith("Q") and any(ch.isdigit() for ch in s[:4])):
            current_q = s
            continue
        if needle in s:
            if current_q and current_q != last_q_emitted:
                if out:
                    out.append("")
                out.append(current_q)
                last_q_emitted = current_q
            out.append(s)
    if not out:
        return f"（ページ {page_num} に対応する [P.{page_num}] 行が見つかりませんでした。"  \
               f"「質問ごとにまとめる」加工を実行してください。）"
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Cost / usage dashboard
# ---------------------------------------------------------------------------

def _render_cost_dashboard() -> None:
    summary = cost_tracker.summary()
    if summary["total_calls"] == 0:
        st.caption("まだ API 呼び出しは行われていません。文字起こしや加工を実行すると統計が表示されます。")
        return

    # ---- メトリクスカード ----
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            _metric_card("総呼び出し回数", f"{summary['total_calls']}", color=_COLOR_PRIMARY),
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            _metric_card("合計処理時間", f"{summary['total_duration_sec']:.1f} 秒", color=_COLOR_ACCENT),
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            _metric_card(
                "失敗回数",
                f"{summary['failed_calls']}",
                color=_COLOR_NEGATIVE if summary["failed_calls"] > 0 else _COLOR_NEUTRAL,
            ),
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            _metric_card("推定コスト", f"${summary['estimated_cost_usd']:.6f}", color=_COLOR_WARNING),
            unsafe_allow_html=True,
        )

    st.markdown('<div style="margin-top: 16px;"></div>', unsafe_allow_html=True)

    c5, c6, c7 = st.columns(3)
    with c5:
        st.markdown(
            _metric_card("入力文字数 合計", f"{summary['total_input_chars']:,}", color="#0f172a"),
            unsafe_allow_html=True,
        )
    with c6:
        st.markdown(
            _metric_card("出力文字数 合計", f"{summary['total_output_chars']:,}", color="#0f172a"),
            unsafe_allow_html=True,
        )
    with c7:
        st.markdown(
            _metric_card("1回あたり平均", f"{summary['average_duration_sec']:.1f} 秒", color="#0f172a"),
            unsafe_allow_html=True,
        )

    st.markdown('<div style="margin-top: 18px;"></div>', unsafe_allow_html=True)

    # ---- 目的別サマリー (横棒チャート + テーブル) ----
    per_purpose = cost_tracker.per_purpose_summary()
    if per_purpose:
        col_chart, col_table = st.columns([3, 4], gap="large")
        with col_chart:
            st.markdown("**用途別の呼び出し回数**")
            pp_df = pd.DataFrame(
                [
                    {"purpose": p["purpose"], "calls": p["calls"], "duration": p["duration_sec"]}
                    for p in per_purpose
                ]
            )
            pp_chart = (
                alt.Chart(pp_df)
                .mark_bar(cornerRadiusEnd=4)
                .encode(
                    x=alt.X("calls:Q", title="回数", axis=alt.Axis(tickMinStep=1)),
                    y=alt.Y(
                        "purpose:N",
                        sort="-x",
                        title=None,
                        axis=alt.Axis(labelLimit=240, labelFontSize=12),
                    ),
                    color=alt.Color("calls:Q", scale=alt.Scale(scheme="blues"), legend=None),
                    tooltip=[
                        alt.Tooltip("purpose:N", title="用途"),
                        alt.Tooltip("calls:Q", title="回数"),
                        alt.Tooltip("duration:Q", title="合計秒"),
                    ],
                )
                .properties(height=max(200, 32 * len(pp_df)))
                .configure_view(strokeWidth=0)
                .configure_axis(labelColor="#475569", titleColor="#334155")
            )
            st.altair_chart(pp_chart, use_container_width=True)
        with col_table:
            st.markdown("**用途別サマリー（詳細）**")
            rows = [
                {
                    "用途": p["purpose"],
                    "回数": p["calls"],
                    "合計秒": round(p["duration_sec"], 1),
                    "入力chars": p["input_chars"],
                    "出力chars": p["output_chars"],
                    "失敗": p["errors"],
                    "推定$": f"${p['estimated_cost_usd']:.6f}",
                }
                for p in per_purpose
            ]
            st.dataframe(rows, use_container_width=True, hide_index=True)

    with st.expander("直近の呼び出し履歴（最大50件）", expanded=False):
        records = cost_tracker.recent_records(50)
        if records:
            display_rows = [
                {
                    "purpose": r["purpose"],
                    "duration(s)": round(r["duration_sec"], 2),
                    "input_chars": r["input_chars"],
                    "output_chars": r["output_chars"],
                    "image_bytes": r["image_bytes"],
                    "status": r["status"],
                }
                for r in records
            ]
            st.dataframe(display_rows, use_container_width=True, hide_index=True)

    st.caption(
        f"※ コスト見積もりは入力 ${summary['input_price_per_1m_usd']}/1M tokens, "
        f"出力 ${summary['output_price_per_1m_usd']}/1M tokens, "
        f"日本語約 {summary['chars_per_token']} 文字/トークンの仮定で算出した目安値です。"
    )

    if st.button("🗑️ 集計をリセット", key="cost_reset"):
        cost_tracker.reset()
        st.rerun()


# ---------------------------------------------------------------------------
# Auto-generate exclude template from uploaded files
# ---------------------------------------------------------------------------

def _render_auto_template_section() -> None:
    """アップロード済みファイルの文字起こし結果から、除外テンプレート候補を自動生成."""
    if not st.session_state.processed_files:
        st.caption(
            "ファイルをアップロードして文字起こしを実行すると、"
            "印字テキストから除外テンプレート候補を自動抽出できます。"
        )
        return

    st.caption(
        "現在の処理結果から、印字された質問文・タイトル・案内文などを抽出して"
        "「除外テキスト」候補として提示します。確認後、テンプレートとして保存できます。"
    )

    col_run, col_clear = st.columns([3, 1])
    with col_run:
        if st.button(
            "🪄 アップロード結果から候補を抽出",
            key="auto_tmpl_run",
            use_container_width=True,
            type="primary",
        ):
            try:
                editor = TextEditor()
                # 全ファイルの transcription を結合（current_text ではなく原文を優先）
                combined = "\n\n---\n\n".join(
                    f"【{pf['filename']}】\n{pf.get('transcription', '')}"
                    for pf in st.session_state.processed_files
                )
                with st.spinner("テンプレート候補を抽出中..."):
                    items = survey_analyzer.generate_exclude_template(
                        combined,
                        caller=lambda p: editor.call_with_purpose(p, "テンプレート自動生成"),
                    )
                st.session_state.auto_template_candidates = items
                st.rerun()
            except Exception as e:
                st.error(f"抽出エラー: {str(e)}")
    with col_clear:
        if st.session_state.auto_template_candidates:
            if st.button("🗑️ クリア", key="auto_tmpl_clear", use_container_width=True):
                st.session_state.auto_template_candidates = []
                st.rerun()

    if not st.session_state.auto_template_candidates:
        return

    st.markdown(f"**抽出された候補（{len(st.session_state.auto_template_candidates)} 件）**")
    edited_lines = st.text_area(
        "抽出された候補（必要に応じて編集）",
        value="\n".join(st.session_state.auto_template_candidates),
        height=240,
        key="auto_tmpl_edit",
    )

    col_apply, col_save = st.columns(2)
    with col_apply:
        if st.button(
            "📝 現在の除外テキストに反映",
            key="auto_tmpl_apply",
            use_container_width=True,
        ):
            new_items = [l.strip() for l in edited_lines.splitlines() if l.strip()]
            existing = {
                l.strip() for l in st.session_state.exclude_texts.splitlines() if l.strip()
            }
            merged = list(existing) + [it for it in new_items if it not in existing]
            st.session_state.exclude_texts = "\n".join(merged)
            _invalidate_processor()
            st.success(f"{len(new_items)} 件を除外テキストに反映しました。")

    with col_save:
        save_name = st.text_input(
            "テンプレート名（任意）",
            key="auto_tmpl_name",
            placeholder="例: 自動生成_研修アンケート",
        )
        if st.button(
            "💾 テンプレートとして保存",
            key="auto_tmpl_save",
            use_container_width=True,
        ):
            if not save_name.strip():
                st.warning("テンプレート名を入力してください。")
            else:
                new_items = [l.strip() for l in edited_lines.splitlines() if l.strip()]
                if not new_items:
                    st.warning("候補が空です。")
                else:
                    try:
                        add_template(save_name.strip(), new_items)
                        st.success(f"テンプレート「{save_name.strip()}」を保存しました。")
                    except ValueError as e:
                        st.error(str(e))


# ---------------------------------------------------------------------------
# Results section
# ---------------------------------------------------------------------------

def _render_results() -> None:
    if not st.session_state.processed_files:
        return

    st.divider()
    st.subheader("📄 処理結果")
    st.caption(
        f"{len(st.session_state.processed_files)} 件のファイルを処理しました。"
        "各ファイルのタブを切り替えて、編集・分析・対比表示・ダウンロードを行えます。"
    )

    doc_generator = DocumentGenerator()

    for idx, pf in enumerate(st.session_state.processed_files):
        status_chip = (
            f'<span style="background:{_COLOR_POSITIVE}15;color:{_COLOR_POSITIVE};'
            f'padding:3px 10px;border-radius:10px;font-size:0.75rem;font-weight:600;margin-left:10px;">'
            f'識別子付与済</span>'
            if has_respondent_id(pf["current_text"])
            else f'<span style="background:{_COLOR_WARNING}15;color:#b45309;'
                 f'padding:3px 10px;border-radius:10px;font-size:0.75rem;font-weight:600;margin-left:10px;">'
                 f'未集約</span>'
        )
        st.markdown(
            f'<div class="file-card-header">📋 {pf["filename"]}'
            f'{status_chip}'
            f'<span style="color:#94a3b8; font-weight:400; margin-left:12px;">'
            f'（ページ数: {len(pf.get("page_images") or [])} / 文字数: {len(pf.get("current_text", "")):,}）'
            f'</span></div>',
            unsafe_allow_html=True,
        )
        _render_matome_banner(idx, pf)
        _render_file_tabs(idx, pf, doc_generator)
        st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)


def _render_file_tabs(idx: int, pf: dict, doc_generator: DocumentGenerator) -> None:
    """1ファイル分の処理結果をタブで構成して表示する."""
    tab_edit, tab_resp, tab_quant, tab_image, tab_dl = st.tabs(
        [
            "✏️ 編集・加工",
            "👥 回答者ビュー",
            "📊 定量分析",
            "🖼️ 原画像対比",
            "📥 ダウンロード",
        ]
    )

    with tab_edit:
        _render_edit_tab(idx, pf)

    with tab_resp:
        _render_respondent_view(idx, pf)

    with tab_quant:
        _render_quant_summary(idx, pf)

    with tab_image:
        _render_image_comparison(idx, pf)

    with tab_dl:
        _render_download_tab(idx, pf, doc_generator)


def _render_edit_tab(idx: int, pf: dict) -> None:
    """文字起こしテキストの編集・加工 UI.

    「質問ごとにまとめる」をプライマリ操作として目立たせ、
    他の加工モード（整文・要約・カスタム）はセカンダリ操作として提示する。
    """
    ver_key = f"text_ver_{idx}"
    if ver_key not in st.session_state:
        st.session_state[ver_key] = 0
    ver = st.session_state[ver_key]

    col_text, col_tools = st.columns([3, 2], gap="large")

    with col_text:
        st.markdown("**✏️ 文字起こし結果**")
        st.caption(
            "テキストエリアを直接クリックして編集できます。"
            "加工後の結果もここに反映されます。"
        )
        edited = st.text_area(
            "文字起こし結果",
            value=pf["current_text"],
            height=520,
            key=f"text_area_{idx}_v{ver}",
            label_visibility="collapsed",
        )
        if edited != pf["current_text"]:
            pf["current_text"] = edited

    with col_tools:
        # ---- プライマリ: 質問ごとにまとめる ----
        st.markdown("**🎯 ステップ 1: 質問ごとにまとめる（推奨）**")
        st.caption(
            "印字された質問ごとに回答を集約し、各回答に [P.N] の回答者識別子を付与します。"
            "回答者ビュー・定量分析・Excel 出力など多くの機能で利用されます。"
        )
        matome_info = EDITING_MODES["matome"]
        with st.expander("この処理の詳細を見る", expanded=False):
            st.markdown(matome_info["description"])
        already_done = has_respondent_id(pf["current_text"])
        if st.button(
            "🔁 もう一度「質問ごとにまとめる」を適用" if already_done else "✨ 質問ごとにまとめる を実行",
            key=f"apply_matome_{idx}",
            use_container_width=True,
            type="primary",
        ):
            _apply_matome(idx, pf)

        st.markdown('<div style="margin-top: 18px;"></div>', unsafe_allow_html=True)
        st.markdown("---")

        # ---- セカンダリ: その他の加工 ----
        st.markdown("**🛠️ ステップ 2: 追加のテキスト加工（任意）**")
        st.caption(
            "必要に応じて文体を整えたり、要約・翻訳・敬語化などを行えます。"
        )

        secondary_modes = {k: v for k, v in EDITING_MODES.items() if k != "matome"}

        with st.expander("各モードの説明を見る", expanded=False):
            for k, info in secondary_modes.items():
                st.markdown(f"**{info['label']}**：{info['description']}")

        edit_mode = st.radio(
            "加工モード",
            options=list(secondary_modes.keys()),
            format_func=lambda x: f"{secondary_modes[x]['label']} — {secondary_modes[x]['short_desc']}",
            horizontal=False,
            key=f"edit_mode_{idx}",
            label_visibility="collapsed",
        )

        custom_prompt = ""
        if edit_mode == "custom":
            custom_prompt = st.text_area(
                "カスタムプロンプト",
                placeholder=(
                    "例: 敬語に変換してください\n"
                    "例: 重要な箇所を箇条書きにまとめてください\n"
                    "例: 英語に翻訳してください"
                ),
                height=100,
                key=f"custom_prompt_{idx}",
            )

        col_apply, col_reset = st.columns([2, 1])
        with col_apply:
            if st.button(
                f"✨ {secondary_modes[edit_mode]['label']}を適用",
                key=f"apply_{idx}",
                use_container_width=True,
            ):
                if edit_mode == "custom" and not custom_prompt.strip():
                    st.warning("カスタムプロンプトを入力してください。")
                else:
                    try:
                        editor = TextEditor()
                        with st.spinner("テキストを加工中..."):
                            result = editor.apply_editing(
                                pf["current_text"], edit_mode, custom_prompt
                            )
                        pf["current_text"] = result
                        st.session_state[ver_key] += 1
                        st.rerun()
                    except Exception as e:
                        st.error(f"加工エラー: {str(e)}")
        with col_reset:
            if st.button(
                "↩️ 元に戻す",
                key=f"reset_{idx}",
                use_container_width=True,
            ):
                pf["current_text"] = pf["transcription"]
                st.session_state[ver_key] += 1
                st.rerun()


def _render_download_tab(idx: int, pf: dict, doc_generator: DocumentGenerator) -> None:
    """ダウンロード（Word / Excel）タブ."""
    st.markdown("**📥 ダウンロード**")
    st.caption(
        "現在のテキスト（加工済みの場合は加工後）でファイルを生成します。"
        "Word と Excel から選択できます。"
    )
    col_dl_w, col_dl_x = st.columns(2, gap="large")
    with col_dl_w:
        st.markdown("**📄 Word**")
        st.caption("文字起こし結果を Word ファイルとして書き出します。")
        try:
            word_content = doc_generator.create_document(
                pf["current_text"], pf["filename"]
            )
            st.download_button(
                label="📥 Word ファイルをダウンロード",
                data=word_content,
                file_name=f"{extract_filename(pf['filename'])}_transcription.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                use_container_width=True,
                key=f"download_{idx}",
                type="primary",
            )
        except Exception as e:
            st.error(f"Word 生成エラー: {str(e)}")
    with col_dl_x:
        st.markdown("**📊 Excel**")
        st.caption(
            "質問別シート・回答者別マトリクス・定量サマリーを 1 ファイルで書き出します。"
        )
        try:
            parsed_q = survey_analyzer.parse_consolidated_text(pf["current_text"])
            respondents = survey_analyzer.to_respondent_view(parsed_q)
            quant = pf.get("quant_summary")
            excel_bytes = excel_exporter.build_workbook(
                questions=parsed_q,
                respondents=respondents,
                quant_summary=quant,
                source_filename=pf["filename"],
            )
            st.download_button(
                label="📥 Excel ファイルをダウンロード",
                data=excel_bytes,
                file_name=f"{extract_filename(pf['filename'])}_transcription.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
                key=f"download_xlsx_{idx}",
                type="primary",
            )
        except Exception as e:
            st.warning(
                "Excel を生成するには、まず「質問ごとにまとめる」加工で"
                "[P.N] 識別子付きテキストにしてください。"
            )
            st.caption(f"内部エラー: {e}")

    # ---- Survey analysis ----
    st.divider()
    st.markdown("### 📊 アンケート分析")
    st.caption(
        "文字起こし結果（加工済みの場合はその内容）をもとに、"
        "全体の傾向・課題・アクションプラン・研修改善案を自動生成します。"
    )

    if st.button(
        "📊 アンケートをまとめて分析",
        use_container_width=True,
        type="primary",
        key="analyze_btn",
    ):
        try:
            editor = TextEditor()
            texts = [pf["current_text"] for pf in st.session_state.processed_files]
            filenames = [pf["filename"] for pf in st.session_state.processed_files]
            with st.spinner("アンケートを分析中... しばらくお待ちください"):
                analysis = editor.analyze_survey(texts, filenames)
            st.session_state.survey_analysis = analysis
            st.rerun()
        except Exception as e:
            st.error(f"分析エラー: {str(e)}")

    if st.session_state.survey_analysis:
        st.markdown("**分析結果**")
        st.caption("テキストエリアを直接編集できます。")
        st.text_area(
            "分析結果",
            value=st.session_state.survey_analysis,
            height=500,
            key="analysis_area",
            label_visibility="collapsed",
        )
        try:
            analysis_word = doc_generator.create_analysis_document(
                st.session_state.survey_analysis
            )
            col_dl, col_clear = st.columns([3, 1])
            with col_dl:
                st.download_button(
                    label="📥 分析結果を Word でダウンロード",
                    data=analysis_word,
                    file_name="アンケート分析結果.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="download_analysis",
                )
            with col_clear:
                if st.button("🗑️ 分析をクリア", use_container_width=True, key="clear_analysis"):
                    st.session_state.survey_analysis = ""
                    st.rerun()
        except Exception as e:
            st.error(f"Word 生成エラー: {str(e)}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="手書きアンケートOCR",
        page_icon="📝",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    _init_session_state()
    _inject_global_css()

    # ---- Sidebar: 設定パネル ----
    _render_sidebar()

    # ---- Header (メイン) ----
    col_title, col_meta = st.columns([5, 2])
    with col_title:
        st.markdown(
            f'<h1 style="margin-bottom:0;">📝 手書きアンケート OCR・文字起こしアプリ</h1>'
            f'<div style="color:#64748b;margin-top:4px;">'
            f'手書きアンケートを Gemini AI で文字起こしし、編集・分析・出力までを 1 つの画面で行えます。'
            f'</div>',
            unsafe_allow_html=True,
        )
    with col_meta:
        st.markdown(
            f'<div style="text-align:right;padding-top:12px;">'
            f'<span style="background:{_COLOR_PRIMARY}15;color:{_COLOR_PRIMARY};'
            f'padding:6px 12px;border-radius:14px;font-size:0.85rem;font-weight:600;">'
            f'Ver. {APP_VERSION}</span>'
            f'<div style="color:#94a3b8;font-size:0.8rem;margin-top:4px;">'
            f'{APP_UPDATED} 更新</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    st.markdown('<div style="margin-bottom: 20px;"></div>', unsafe_allow_html=True)

    # ---- Main tabs ----
    tab_upload, tab_proof, tab_template, tab_cost, tab_about = st.tabs(
        [
            "📁 アップロード & 結果",
            "📝 文字校正チェック",
            "🪄 設問テンプレート生成",
            "💰 API 利用状況",
            "ℹ️ アプリ情報",
        ]
    )

    with tab_upload:
        _render_upload_tab()

    with tab_proof:
        _render_proof_tab()

    with tab_template:
        _render_template_section_tab()

    with tab_cost:
        _render_cost_tab()

    with tab_about:
        _render_about_tab()


def _render_sidebar() -> None:
    """サイドバーに OCR 設定・除外設定・更新履歴をまとめる."""
    with st.sidebar:
        st.markdown("### ⚙️ 文字起こし設定")
        _render_ocr_mode_section()

        st.markdown("---")
        st.markdown("### 🚫 文字起こし除外設定")
        st.caption(
            "アンケート用紙に印刷されているタイトルや質問文など、"
            "文字起こし不要のテキストを指定できます。"
        )
        _render_exclude_section()

        st.markdown("---")
        with st.expander("ℹ️ 使用モデル情報", expanded=False):
            st.markdown(
                f"**{MODEL_LABEL}** &nbsp; `{MODEL_NAME}`  \n"
                f"{MODEL_DESCRIPTION}"
            )


def _render_upload_tab() -> None:
    st.subheader("📁 ファイルのアップロード")
    st.caption(
        "手書きアンケートのファイルを選択してください（複数選択可能）。"
        "対応形式: PDF, PNG, JPG, JPEG, GIF, WebP, BMP, TIFF"
    )

    supported_types = [ext.lstrip(".") for ext in get_supported_extensions_list()]

    uploaded_files = st.file_uploader(
        "ファイルを選択",
        type=supported_types,
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        st.success(f"{len(uploaded_files)} 個のファイルがアップロードされました")
        with st.expander("アップロードされたファイル一覧", expanded=False):
            for i, f in enumerate(uploaded_files):
                st.write(f"{i + 1}. **{f.name}** ({f.size:,} bytes)")
        if st.button(
            "📝 文字起こしを開始",
            type="primary",
            use_container_width=True,
        ):
            _process_files(uploaded_files)

    _render_results()

    if st.session_state.processed_files:
        st.divider()
        if st.button(
            "🗑️ 全ての処理結果をクリア",
            use_container_width=False,
            key="clear_all_results",
        ):
            st.session_state.processed_files = []
            st.session_state.survey_analysis = ""
            st.rerun()


def _render_proof_tab() -> None:
    st.subheader("📝 文字校正チェック（コピペ）")
    _render_proofread_section()


def _render_template_section_tab() -> None:
    st.subheader("🪄 設問テンプレートの自動生成")
    _render_auto_template_section()


def _render_cost_tab() -> None:
    st.subheader("💰 API 利用状況・処理時間")
    _render_cost_dashboard()


def _render_about_tab() -> None:
    st.subheader("ℹ️ アプリ情報")
    st.markdown(
        f"**{MODEL_LABEL}** &nbsp; `{MODEL_NAME}`  \n"
        f"{MODEL_DESCRIPTION}"
    )
    st.markdown("---")
    st.markdown("### 📋 更新履歴")
    for entry in CHANGELOG:
        with st.expander(f"v{entry['version']} — {entry['date']}", expanded=False):
            for change in entry["changes"]:
                st.markdown(f"- {change}")


if __name__ == "__main__":
    main()

import streamlit as st
import os
import io
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
    st.subheader("⚙️ 文字起こしモード")

    mode = st.radio(
        "文字起こしモード",
        options=["accurate", "proofread"],
        format_func=lambda x: {
            "accurate": "✍️ 正確文字起こし",
            "proofread": "🔍 校正文字起こし",
        }[x],
        index=0 if st.session_state.ocr_mode == "accurate" else 1,
        horizontal=True,
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
    """[P.N] でグルーピングした回答者単位ビュー."""
    st.markdown("**👥 回答者単位ビュー**")
    st.caption(
        "[P.N] や [回答者N] の識別子を基準にグルーピングして、回答者ごとに縦読みします。"
        "「質問ごとにまとめる」加工を先に実行しておくと識別子が確実に付与されます。"
    )
    parsed = survey_analyzer.parse_consolidated_text(pf["current_text"])
    respondents = survey_analyzer.to_respondent_view(parsed)
    if not respondents:
        st.info(
            "回答者識別子が検出できませんでした。"
            "「質問ごとにまとめる」加工を実行してから再度開いてください。"
        )
        return

    for block in respondents:
        with st.expander(f"[{block.respondent_id}] の回答（{len(block.answers)} 件）", expanded=False):
            for q_num, q_text, ans in block.answers:
                st.markdown(f"- **Q{q_num}: {q_text}**")
                st.markdown(f"  - {ans}")


def _render_quant_summary(idx: int, pf: dict) -> None:
    """センチメント分類・トピック・キーワードの定量サマリー."""
    st.markdown("**📊 定量サマリー（センチメント・トピック・キーワード）**")
    st.caption(
        "現在のテキストを Gemini に解析させ、ポジ/ネガ/中立の件数、"
        "トピック・キーワード・代表回答を JSON で返してもらいます。"
    )

    summary_key = "quant_summary"
    col_run, col_clear = st.columns([3, 1])
    with col_run:
        if st.button(
            "📈 定量サマリーを生成",
            key=f"quant_btn_{idx}",
            use_container_width=True,
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
        return

    if "_raw" in summary:
        st.warning("JSON 解析に失敗しました。元応答を表示します。")
        st.code(summary.get("_raw", ""))
        return

    sent = summary.get("sentiment", {}) or {}
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("回答数", summary.get("answer_count", 0))
    col2.metric("ポジティブ", sent.get("positive", 0))
    col3.metric("ネガティブ", sent.get("negative", 0))
    col4.metric("中立", sent.get("neutral", 0))

    topics = summary.get("topics", []) or []
    if topics:
        st.markdown("**トピック（多い順）**")
        topic_rows = [
            {
                "トピック": t.get("label", ""),
                "件数": t.get("count", 0),
                "代表表現": " / ".join(t.get("examples", []) or []),
            }
            for t in topics
        ]
        st.dataframe(topic_rows, use_container_width=True, hide_index=True)
        try:
            chart_data = {t.get("label", f"トピック{i}"): t.get("count", 0) for i, t in enumerate(topics)}
            st.bar_chart(chart_data)
        except Exception:
            pass

    keywords = summary.get("keywords", []) or []
    if keywords:
        st.markdown("**頻出キーワード**")
        st.dataframe(
            [{"キーワード": k.get("word", ""), "出現数": k.get("count", 0)} for k in keywords],
            use_container_width=True,
            hide_index=True,
        )

    highlights = summary.get("highlights", {}) or {}
    pos = highlights.get("positive", []) or []
    neg = highlights.get("negative", []) or []
    if pos or neg:
        st.markdown("**代表的な回答**")
        if pos:
            st.markdown("- ポジティブ:")
            for s in pos:
                st.markdown(f"  - {s}")
        if neg:
            st.markdown("- ネガティブ:")
            for s in neg:
                st.markdown(f"  - {s}")


def _render_image_comparison(idx: int, pf: dict) -> None:
    """原画像との対比ビュー."""
    st.markdown("**🖼️ 原画像との対比**")
    page_images = pf.get("page_images") or []
    if not page_images:
        st.info("原画像が保持されていません（旧フォーマットの結果）。再アップロードで利用可能になります。")
        return
    st.caption("該当ページを選んで、OCRテキストと並べて確認できます。")

    page_options = [pnum for _b, pnum in page_images]
    selected = st.selectbox(
        "表示するページ",
        options=page_options,
        index=0,
        key=f"img_page_{idx}",
    )

    img_bytes = None
    for b, pnum in page_images:
        if pnum == selected:
            img_bytes = b
            break

    if img_bytes is None:
        st.warning("画像が見つかりませんでした。")
        return

    col_img, col_txt = st.columns([1, 1])
    with col_img:
        st.image(img_bytes, caption=f"ページ {selected}", use_container_width=True)
    with col_txt:
        st.markdown(f"**ページ {selected} の文字起こしテキスト**")
        st.caption("[P.{selected}] 識別子を含む行を抽出して表示しています。".replace("{selected}", str(selected)))
        snippet = _extract_page_snippet(pf["current_text"], selected)
        st.text_area(
            f"ページ {selected} 抜粋",
            value=snippet,
            height=320,
            key=f"img_text_{idx}_{selected}",
            label_visibility="collapsed",
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

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("総呼び出し回数", summary["total_calls"])
    col2.metric("処理時間 合計", f"{summary['total_duration_sec']:.1f} 秒")
    col3.metric("失敗", summary["failed_calls"])
    col4.metric("推定コスト (USD)", f"${summary['estimated_cost_usd']:.6f}")

    col5, col6, col7 = st.columns(3)
    col5.metric("入力文字数 合計", f"{summary['total_input_chars']:,}")
    col6.metric("出力文字数 合計", f"{summary['total_output_chars']:,}")
    col7.metric("1回あたり平均時間", f"{summary['average_duration_sec']:.1f} 秒")

    per_purpose = cost_tracker.per_purpose_summary()
    if per_purpose:
        st.markdown("**目的別サマリー**")
        rows = [
            {
                "用途": p["purpose"],
                "回数": p["calls"],
                "合計秒": f"{p['duration_sec']:.1f}",
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

    doc_generator = DocumentGenerator()

    for idx, pf in enumerate(st.session_state.processed_files):
        with st.expander(f"📋 {pf['filename']}", expanded=True):

            # Version counter: incremented on apply/reset to force text_area re-render
            ver_key = f"text_ver_{idx}"
            if ver_key not in st.session_state:
                st.session_state[ver_key] = 0
            ver = st.session_state[ver_key]

            # ---- Editable transcription ----
            st.markdown("**✏️ 文字起こし結果**")
            st.caption("テキストエリアを直接クリックして編集できます。加工後の結果もここに反映されます。")
            edited = st.text_area(
                "文字起こし結果",
                value=pf["current_text"],
                height=300,
                key=f"text_area_{idx}_v{ver}",
                label_visibility="collapsed",
            )
            # Capture manual edits
            if edited != pf["current_text"]:
                pf["current_text"] = edited

            st.divider()

            # ---- Text editing tools ----
            st.markdown("**🛠️ テキスト加工**")
            st.caption("加工を適用すると、上のテキストエリアに結果が反映されます。そのまま追加編集も可能です。")

            with st.expander("各モードの説明を見る", expanded=False):
                for k, info in EDITING_MODES.items():
                    st.markdown(f"**{info['label']}**：{info['description']}")

            edit_mode = st.radio(
                "加工モード",
                options=list(EDITING_MODES.keys()),
                format_func=lambda x: f"{EDITING_MODES[x]['label']}　— {EDITING_MODES[x]['short_desc']}",
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
                    f"✨ {EDITING_MODES[edit_mode]['label']}を適用",
                    key=f"apply_{idx}",
                    use_container_width=True,
                    type="primary",
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
                            st.session_state[ver_key] += 1  # Force text_area re-render
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
                    st.session_state[ver_key] += 1  # Force text_area re-render
                    st.rerun()

            st.divider()

            # ---- Respondent view (per-respondent grouping) ----
            _render_respondent_view(idx, pf)

            st.divider()

            # ---- Quantitative summary ----
            _render_quant_summary(idx, pf)

            st.divider()

            # ---- Original image comparison ----
            _render_image_comparison(idx, pf)

            st.divider()

            # ---- Download ----
            st.markdown("**📥 ダウンロード**")
            st.caption(
                "現在のテキスト（加工済みの場合は加工後）でファイルを生成します。"
                "Word と Excel から選択できます。"
            )
            col_dl_w, col_dl_x = st.columns(2)
            with col_dl_w:
                try:
                    word_content = doc_generator.create_document(
                        pf["current_text"], pf["filename"]
                    )
                    st.download_button(
                        label="📥 Word",
                        data=word_content,
                        file_name=f"{extract_filename(pf['filename'])}_transcription.docx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        ),
                        use_container_width=True,
                        key=f"download_{idx}",
                    )
                except Exception as e:
                    st.error(f"Word 生成エラー: {str(e)}")
            with col_dl_x:
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
                        label="📥 Excel",
                        data=excel_bytes,
                        file_name=f"{extract_filename(pf['filename'])}_transcription.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        use_container_width=True,
                        key=f"download_xlsx_{idx}",
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

    st.divider()
    if st.button("🗑️ すべてクリア", use_container_width=True):
        st.session_state.processed_files = []
        st.session_state.survey_analysis = ""
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

    with st.expander("📋 更新履歴", expanded=False):
        for entry in CHANGELOG:
            st.markdown(
                f"**v{entry['version']}** &nbsp; "
                f"<span style='color:gray; font-size:0.9em;'>{entry['date']}</span>",
                unsafe_allow_html=True,
            )
            for change in entry["changes"]:
                st.markdown(f"&nbsp;&nbsp;• {change}")

    st.divider()

    # ---- OCR mode ----
    _render_ocr_mode_section()

    st.divider()

    # ---- Exclude section (collapsed by default) ----
    with st.expander("🚫 文字起こし除外設定", expanded=False):
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
    _render_results()

    # ---- Standalone proofread section ----
    st.divider()
    st.subheader("📝 文字校正チェック（コピペ）")
    with st.expander("テキストを貼り付けて誤字脱字・違和感をチェックする", expanded=False):
        _render_proofread_section()

    # ---- Auto-generate exclude template ----
    st.divider()
    st.subheader("🪄 設問テンプレートの自動生成")
    with st.expander("アップロード結果から除外テンプレートを抽出する", expanded=False):
        _render_auto_template_section()

    # ---- API cost / time dashboard ----
    st.divider()
    st.subheader("💰 API 利用状況・処理時間")
    with st.expander("呼び出し回数・処理時間・推定コストを見る", expanded=False):
        _render_cost_dashboard()


if __name__ == "__main__":
    main()

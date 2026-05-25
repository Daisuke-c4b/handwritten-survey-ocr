"""校正前後の差分プレビュー HTML を生成する."""
from __future__ import annotations

import difflib
import html


def _esc(s: str) -> str:
    return html.escape(s, quote=False).replace("\n", "<br>")


def generate_diff_html(original: str, corrected: str) -> str:
    """文字単位の差分を HTML で返す。
    - 追加: 緑 (#d4edda)
    - 削除: 赤 (#f8d7da) 取り消し線
    - 共通: 標準色
    Streamlit 側で unsafe_allow_html=True で表示する。
    """
    if original == corrected:
        return (
            "<div style='padding:8px;border:1px solid #ccc;background:#f5f5f5;'>"
            "<em>差分はありません（修正提案なし）。</em></div>"
        )

    matcher = difflib.SequenceMatcher(a=original, b=corrected, autojunk=False)
    parts: list[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            parts.append(
                f"<span style='color:#333;'>{_esc(original[i1:i2])}</span>"
            )
        elif tag == "delete":
            parts.append(
                "<span style='background:#f8d7da;color:#721c24;"
                "text-decoration:line-through;'>"
                f"{_esc(original[i1:i2])}</span>"
            )
        elif tag == "insert":
            parts.append(
                "<span style='background:#d4edda;color:#155724;'>"
                f"{_esc(corrected[j1:j2])}</span>"
            )
        elif tag == "replace":
            parts.append(
                "<span style='background:#f8d7da;color:#721c24;"
                "text-decoration:line-through;'>"
                f"{_esc(original[i1:i2])}</span>"
            )
            parts.append(
                "<span style='background:#d4edda;color:#155724;'>"
                f"{_esc(corrected[j1:j2])}</span>"
            )

    body = "".join(parts)
    legend = (
        "<div style='font-size:0.9em;color:#555;margin-bottom:6px;'>"
        "<span style='background:#d4edda;color:#155724;padding:1px 4px;'>追加</span>"
        " &nbsp; "
        "<span style='background:#f8d7da;color:#721c24;text-decoration:line-through;"
        "padding:1px 4px;'>削除</span>"
        "</div>"
    )
    return (
        f"<div>{legend}"
        "<div style='padding:10px;border:1px solid #ddd;border-radius:6px;"
        "background:#fff;line-height:1.7;font-family:\"Hiragino Sans\","
        "\"Yu Gothic\",sans-serif;'>"
        f"{body}"
        "</div></div>"
    )

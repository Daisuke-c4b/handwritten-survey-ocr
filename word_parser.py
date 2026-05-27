"""Word (.docx) ファイルからテキストを抽出."""

from __future__ import annotations

from io import BytesIO

from docx import Document


def extract_text_from_docx(file_bytes: bytes) -> str:
    """Word 文書の段落テキストを改行区切りで返す."""
    if not file_bytes:
        return ""
    doc = Document(BytesIO(file_bytes))
    lines: list[str] = []
    for para in doc.paragraphs:
        text = (para.text or "").strip()
        if text:
            lines.append(text)
    return "\n".join(lines)

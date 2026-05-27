"""設問テンプレート（事前登録）の永続化とパース."""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import TypedDict

_TEMPLATE_FILE = Path(__file__).parent / "question_templates.json"

_Q_LINE_RE = re.compile(
    r"^(?:Q|質問|Ｑ)\s*(\d+)\s*[:：]\s*(.+)$",
    re.IGNORECASE,
)
_NUM_LINE_RE = re.compile(r"^(\d+)\s*[\.．、]\s*(.+)$")


class QuestionTemplate(TypedDict):
    name: str
    questions: dict[str, str]  # "1" -> 質問文


def _load_all() -> list[QuestionTemplate]:
    if not _TEMPLATE_FILE.exists():
        return []
    try:
        data = json.loads(_TEMPLATE_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except (json.JSONDecodeError, OSError):
        pass
    return []


def _save_all(templates: list[QuestionTemplate]) -> None:
    _TEMPLATE_FILE.write_text(
        json.dumps(templates, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def list_question_templates() -> list[QuestionTemplate]:
    return _load_all()


def get_question_template(name: str) -> QuestionTemplate | None:
    for t in _load_all():
        if t["name"] == name:
            return t
    return None


def questions_as_dict(template: QuestionTemplate | None) -> dict[int, str]:
    """テンプレートを {1: '質問文', ...} 形式に変換."""
    if not template:
        return {}
    out: dict[int, str] = {}
    for k, v in (template.get("questions") or {}).items():
        try:
            num = int(str(k).strip())
        except ValueError:
            continue
        text = (v or "").strip()
        if text:
            out[num] = text
    return out


def parse_questions_from_input(text: str) -> dict[int, str]:
    """テキスト入力から設問一覧をパース（Q1:/質問1:/1. 形式）."""
    out: dict[int, str] = {}
    for line in (text or "").splitlines():
        s = line.strip()
        if not s:
            continue
        m = _Q_LINE_RE.match(s)
        if m:
            out[int(m.group(1))] = m.group(2).strip()
            continue
        m2 = _NUM_LINE_RE.match(s)
        if m2:
            out[int(m2.group(1))] = m2.group(2).strip()
    return out


def format_questions_for_editor(questions: dict[int, str]) -> str:
    return "\n".join(
        f"Q{num}: {text}"
        for num, text in sorted(questions.items())
    )


def add_question_template(name: str, questions: dict[int, str]) -> None:
    if not name.strip():
        raise ValueError("テンプレート名を入力してください")
    if not questions:
        raise ValueError("設問が1つ以上必要です")
    templates = _load_all()
    for t in templates:
        if t["name"] == name:
            raise ValueError(f"テンプレート「{name}」は既に存在します")
    templates.append({
        "name": name.strip(),
        "questions": {str(k): v for k, v in sorted(questions.items())},
    })
    _save_all(templates)


def update_question_template(
    old_name: str,
    new_name: str,
    questions: dict[int, str],
) -> None:
    if not questions:
        raise ValueError("設問が1つ以上必要です")
    templates = _load_all()
    if old_name != new_name:
        for t in templates:
            if t["name"] == new_name:
                raise ValueError(f"テンプレート「{new_name}」は既に存在します")
    for t in templates:
        if t["name"] == old_name:
            t["name"] = new_name.strip()
            t["questions"] = {str(k): v for k, v in sorted(questions.items())}
            _save_all(templates)
            return
    raise ValueError(f"テンプレート「{old_name}」が見つかりません")


def delete_question_template(name: str) -> None:
    templates = [t for t in _load_all() if t["name"] != name]
    _save_all(templates)

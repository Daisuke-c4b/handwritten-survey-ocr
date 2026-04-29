"""Manage exclude-text templates persisted as a local JSON file."""

from __future__ import annotations

import json
from pathlib import Path
from typing import TypedDict

_TEMPLATE_FILE = Path(__file__).parent / "templates.json"


class Template(TypedDict):
    name: str
    texts: list[str]


def _load_all() -> list[Template]:
    if not _TEMPLATE_FILE.exists():
        return []
    try:
        data = json.loads(_TEMPLATE_FILE.read_text(encoding="utf-8"))
        if isinstance(data, list):
            return data
    except (json.JSONDecodeError, OSError):
        pass
    return []


def _save_all(templates: list[Template]) -> None:
    _TEMPLATE_FILE.write_text(
        json.dumps(templates, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def list_templates() -> list[Template]:
    return _load_all()


def get_template(name: str) -> Template | None:
    for t in _load_all():
        if t["name"] == name:
            return t
    return None


def add_template(name: str, texts: list[str]) -> None:
    templates = _load_all()
    for t in templates:
        if t["name"] == name:
            raise ValueError(f"テンプレート「{name}」は既に存在します")
    templates.append({"name": name, "texts": texts})
    _save_all(templates)


def update_template(old_name: str, new_name: str, texts: list[str]) -> None:
    templates = _load_all()
    if old_name != new_name:
        for t in templates:
            if t["name"] == new_name:
                raise ValueError(f"テンプレート「{new_name}」は既に存在します")
    for t in templates:
        if t["name"] == old_name:
            t["name"] = new_name
            t["texts"] = texts
            _save_all(templates)
            return
    raise ValueError(f"テンプレート「{old_name}」が見つかりません")


def delete_template(name: str) -> None:
    templates = _load_all()
    templates = [t for t in templates if t["name"] != name]
    _save_all(templates)

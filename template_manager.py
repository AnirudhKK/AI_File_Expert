"""
template_manager.py
Handles all template file operations: load, save, delete, append.

Security measures:
  - All templates are confined to TEMPLATES_DIR (no path traversal).
  - Template names are sanitised to [a-z0-9_] only before use as filenames.
  - resolve() + is_relative_to() enforces the directory boundary even if
    a crafted name somehow contains path separators after sanitisation.
  - Template files are read/written with explicit utf-8 encoding.
"""
import os
import re
from pathlib import Path

# All template files live in this directory — never outside it.
TEMPLATES_DIR = Path("templates").resolve()
TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)

# Maximum template file size (bytes) — prevents writing huge files to disk.
MAX_TEMPLATE_BYTES = 1024 * 512   # 512 KB


def _safe_path(template_file: str) -> Path:
    """
    Resolve `template_file` relative to TEMPLATES_DIR and verify it stays
    inside that directory.  Raises ValueError on path-traversal attempts.
    """
    candidate = (TEMPLATES_DIR / template_file).resolve()
    if not candidate.is_relative_to(TEMPLATES_DIR):
        raise ValueError(
            f"Invalid template path '{template_file}': "
            "must stay within the templates directory."
        )
    return candidate


def get_template_filename(template_name: str) -> str:
    """
    Convert a user-supplied name to a safe filename.
    Keeps only lowercase alphanumerics and underscores; collapses runs of
    non-safe chars to a single underscore; strips leading/trailing underscores.
    """
    sanitised = re.sub(r'[^a-z0-9]+', '_', template_name.strip().lower()).strip('_')
    if not sanitised:
        sanitised = "default_template"
    return sanitised + ".txt"


def load_template(template_file: str) -> list[str]:
    path = _safe_path(template_file)
    if not path.exists():
        return []
    return [
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]


def save_template(template_file: str, commands: list[str]) -> None:
    path = _safe_path(template_file)
    content = "\n".join(commands) + "\n"
    if len(content.encode("utf-8")) > MAX_TEMPLATE_BYTES:
        raise ValueError("Template content exceeds maximum allowed size (512 KB).")
    path.write_text(content, encoding="utf-8")


def append_to_template(template_file: str, code: str) -> None:
    path = _safe_path(template_file)
    # Check combined size before appending
    existing = path.read_text(encoding="utf-8") if path.exists() else ""
    combined = existing + code + "\n"
    if len(combined.encode("utf-8")) > MAX_TEMPLATE_BYTES:
        raise ValueError("Appending this code would exceed the template size limit (512 KB).")
    path.write_text(combined, encoding="utf-8")


def delete_template_lines(
    template_file: str, commands: list[str], indices_to_delete: list[int]
) -> list[str]:
    updated = [cmd for i, cmd in enumerate(commands) if i not in indices_to_delete]
    save_template(template_file, updated)
    return updated


def ensure_template_exists(template_file: str) -> bool:
    """Returns True if the template already existed, False if newly created."""
    path = _safe_path(template_file)
    if path.exists():
        return True
    path.touch()
    return False

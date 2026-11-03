"""Runtime helpers shared by the desktop entry points."""

from __future__ import annotations

import importlib
import os
import sys
from pathlib import Path
from typing import Iterable


def relaunch_in_project_venv_if_needed(
    required_module_groups: Iterable[tuple[str, ...]],
) -> None:
    """Relaunch a script with the local virtual environment when needed.

    Each tuple is a group of alternative import names. The relaunch only
    happens when at least one group is entirely unavailable in the current
    interpreter and a project ``.venv`` interpreter exists.
    """

    def can_import_any(alternatives: tuple[str, ...]) -> bool:
        for name in alternatives:
            try:
                importlib.import_module(name)
            except Exception:
                continue
            return True
        return False

    missing_dependency = any(
        not can_import_any(alternatives) for alternatives in required_module_groups
    )
    if not missing_dependency:
        return

    project_dir = Path(__file__).resolve().parent
    candidates = (
        project_dir / ".venv" / "bin" / "python",
        project_dir / ".venv" / "Scripts" / "python.exe",
    )
    project_python = next((path for path in candidates if path.is_file()), None)
    if project_python is None:
        return

    try:
        already_using_project_python = Path(sys.executable).resolve() == project_python.resolve()
    except OSError:
        already_using_project_python = os.path.abspath(sys.executable) == os.path.abspath(project_python)
    if already_using_project_python:
        return

    original_arguments = getattr(sys, "orig_argv", None)
    if original_arguments:
        arguments = original_arguments[1:]
    else:
        arguments = sys.argv
    os.execv(str(project_python), [str(project_python), *arguments])

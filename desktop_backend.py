"""PyInstaller entrypoint for the desktop backend."""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path


def install_builtin_templates() -> None:
    """Copy shipped templates only when the user has not created a local copy."""
    if not getattr(sys, "frozen", False):
        return
    data_dir = Path(os.environ["PAYROLL_DATA_DIR"])
    bundled_templates = Path(sys._MEIPASS) / "templates"  # type: ignore[attr-defined]
    target_templates = data_dir / "templates"
    target_templates.mkdir(parents=True, exist_ok=True)
    for template in bundled_templates.glob("*.xlsx"):
        target = target_templates / template.name
        if not target.exists():
            shutil.copy2(template, target)


def main() -> None:
    install_builtin_templates()
    import uvicorn
    from backend.main import app

    uvicorn.run(app, host="127.0.0.1", port=18000, log_level="warning")


if __name__ == "__main__":
    main()

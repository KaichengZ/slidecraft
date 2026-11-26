#!/usr/bin/env python3
"""Cross-platform installer for the SlideCraft Antigravity plugin."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

IGNORE_NAMES = {".venv", "__pycache__", ".DS_Store", ".git", ".data"}


def _venv_python(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def main() -> None:
    src = Path(__file__).resolve().parent
    home = Path.home()
    plugin_dst = home / ".gemini" / "config" / "plugins" / "slidecraft"
    data_dir = home / ".gemini" / "antigravity" / "plugin_data" / "slidecraft"
    venv_dir = data_dir / ".venv"

    print(f"[SlideCraft] Installing plugin to: {plugin_dst}")
    plugin_dst.parent.mkdir(parents=True, exist_ok=True)
    if plugin_dst.exists():
        shutil.rmtree(plugin_dst)

    def _ignore(_dir, names):
        return [n for n in names if n in IGNORE_NAMES or n.endswith(".pyc")]

    shutil.copytree(src, plugin_dst, ignore=_ignore)

    print(f"[SlideCraft] Setting up Python virtual environment in: {venv_dir}")
    data_dir.mkdir(parents=True, exist_ok=True)
    py_exe = _venv_python(venv_dir)
    if not py_exe.is_file():
        subprocess.check_call([sys.executable, "-m", "venv", str(venv_dir)])

    req_file = plugin_dst / "requirements.txt"
    subprocess.check_call([str(py_exe), "-m", "pip", "install", "--quiet", "--upgrade", "pip"])
    subprocess.check_call([str(py_exe), "-m", "pip", "install", "--quiet", "-r", str(req_file)])

    print("[SlideCraft] Installation complete!")
    print("Antigravity will discover SlideCraft automatically — open it from the right-pane '+' menu or ask Antigravity to create/open a .pptx deck.")


if __name__ == "__main__":
    main()

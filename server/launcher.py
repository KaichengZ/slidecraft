#!/usr/bin/env python3
"""Cross-platform self-bootstrapping launcher for SlideCraft MCP Server.

Ensures python-pptx and Pillow are available in an isolated venv under
--data_dir/.venv (or falls back to the active interpreter if already installed),
then hands off stdio directly to slidecraft_server.py.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

SERVER_DIR = Path(__file__).resolve().parent
PLUGIN_ROOT = SERVER_DIR.parent
SERVER_SCRIPT = SERVER_DIR / "slidecraft_server.py"
REQ_FILE = PLUGIN_ROOT / "requirements.txt"


def _has_deps(python_exe: str) -> bool:
    try:
        res = subprocess.run(
            [python_exe, "-c", "import pptx, PIL"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=15,
        )
        return res.returncode == 0
    except Exception:
        return False


def _venv_python(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def _ensure_python(data_dir: Path) -> str:
    venv_dir = data_dir / ".venv"
    venv_py = _venv_python(venv_dir)
    if venv_py.is_file() and _has_deps(str(venv_py)):
        return str(venv_py)

    # Use current Python immediately if dependencies are already available
    if _has_deps(sys.executable):
        return sys.executable

    # First-run bootstrap: create venv inside data_dir (stderr only, never touch stdout!)
    data_dir.mkdir(parents=True, exist_ok=True)
    if not venv_py.is_file():
        subprocess.run(
            [sys.executable, "-m", "venv", str(venv_dir)],
            stdout=sys.stderr,
            stderr=sys.stderr,
            check=True,
        )
    if REQ_FILE.is_file():
        subprocess.run(
            [str(venv_py), "-m", "pip", "install", "--quiet", "-r", str(REQ_FILE)],
            stdout=sys.stderr,
            stderr=sys.stderr,
            check=True,
        )
    return str(venv_py)


def main() -> None:
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("--data_dir", default=str(Path.home() / ".gemini" / "antigravity" / "plugin_data" / "slidecraft"))
    known, _unknown = ap.parse_known_args()
    data_dir = Path(os.path.expanduser(known.data_dir))

    py_exe = _ensure_python(data_dir)
    cmd = [py_exe, str(SERVER_SCRIPT), *sys.argv[1:]]
    if os.name == "nt":
        sys.exit(subprocess.call(cmd))
    else:
        os.execv(py_exe, cmd)


if __name__ == "__main__":
    main()

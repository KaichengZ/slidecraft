#!/usr/bin/env python3
"""SlideCraft MCP Server (JSON-RPC 2.0 over stdio, MCP Apps SEP-1865).

Cross-platform MCP server powering a lightweight, 60fps, lossless in-place
.pptx slide editor in Antigravity's auxiliary pane and sidebar.
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

SERVER_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SERVER_DIR))

import slidecraft_engine as engine

APP_URI = "ui://slidecraft/editor.html"
SERVER_INFO = {"name": "slidecraft", "version": "1.0.0"}
_DATA_DIR = str(Path.home() / ".gemini" / "antigravity" / "plugin_data" / "slidecraft")


def _recent_path() -> Path:
    return Path(_DATA_DIR) / "recent.json"


def _load_recent() -> List[Dict[str, Any]]:
    p = _recent_path()
    if not p.is_file():
        return []
    try:
        items = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(items, list):
            return [x for x in items if isinstance(x, dict) and os.path.isfile(x.get("file_path", ""))]
    except Exception:
        pass
    return []


def _record_recent(file_path: str, slide_count: int) -> List[Dict[str, Any]]:
    real = os.path.realpath(file_path)
    items = [x for x in _load_recent() if x.get("file_path") != real]
    items.insert(0, {
        "file_path": real,
        "file_name": os.path.basename(real),
        "slide_count": slide_count,
        "updated_at": int(time.time()),
    })
    items = items[:20]
    p = _recent_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    return items


def _discover_pptx_files(directory: Optional[str] = None) -> List[Dict[str, Any]]:
    seen = set()
    results: List[Dict[str, Any]] = []

    for item in _load_recent():
        fp = item["file_path"]
        if fp not in seen and os.path.isfile(fp):
            seen.add(fp)
            results.append(item)

    home = Path.home()
    roots: List[str] = []
    if directory and os.path.isdir(os.path.expanduser(directory)):
        roots.append(os.path.expanduser(directory))
    else:
        for cand in [
            os.getcwd(),
            str(home / ".gemini" / "antigravity" / "scratch"),
            str(home / "Desktop"),
            str(home / "Downloads"),
            str(home / "Documents"),
        ]:
            if os.path.isdir(cand):
                roots.append(cand)

    found: List[Tuple[float, str]] = []
    for root in roots:
        try:
            for pat in ("*.pptx", "*/*.pptx", "*/*/*.pptx"):
                for match in glob.glob(os.path.join(root, pat)):
                    if os.path.basename(match).startswith("~$"):
                        continue
                    real = os.path.realpath(match)
                    if real not in seen and os.path.isfile(real):
                        seen.add(real)
                        found.append((os.path.getmtime(real), real))
        except Exception:
            continue

    found.sort(key=lambda x: x[0], reverse=True)
    for mt, fp in found[:25]:
        results.append({"file_path": fp, "file_name": os.path.basename(fp), "updated_at": int(mt)})
    return results


def _open_or_reveal_file(file_path: str, action: str) -> None:
    real = os.path.realpath(os.path.expanduser(file_path))
    if not os.path.isfile(real):
        raise FileNotFoundError(f"File not found: {real}")
    if sys.platform == "darwin":
        cmd = ["open", "-R", real] if action == "reveal" else ["open", real]
        subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    elif os.name == "nt":
        if action == "reveal":
            subprocess.Popen(["explorer", f"/select,{real}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            os.startfile(real)  # type: ignore[attr-defined]
    else:
        target = os.path.dirname(real) if action == "reveal" else real
        subprocess.Popen(["xdg-open", target], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _adapt_sdk_for_inline(code: str) -> str:
    m = re.search(r"\bexport\s*\{([^}]+)\}\s*;?\s*(?://#[^\n]*)?\s*$", code)
    if not m:
        return f"(() => {{\n{code}\n}})();\n"
    pairs = []
    for part in m.group(1).split(","):
        part = part.strip()
        if not part:
            continue
        if " as " in part:
            loc, exp = [x.strip() for x in part.split(" as ", 1)]
            pairs.append(f"{exp}:{loc}")
        else:
            pairs.append(f"{part}:{part}")
    inner = code[: m.start()] + f"\nwindow.__McpExtApps={{{','.join(pairs)}}};\n"
    return f"(() => {{\n{inner}\n}})();\n"


def _build_editor_html() -> str:
    html_path = os.environ.get("SLIDECRAFT_HTML") or str(SERVER_DIR / "editor.html")
    sdk_path = SERVER_DIR / "ext-apps.mjs"
    text = Path(html_path).read_text(encoding="utf-8")
    if "/*__INLINE_EXT_APPS_SDK__*/" in text and sdk_path.is_file():
        text = text.replace("/*__INLINE_EXT_APPS_SDK__*/", _adapt_sdk_for_inline(sdk_path.read_text(encoding="utf-8")), 1)
    return text


TOOLS: List[Dict[str, Any]] = [
    {
        "name": "open_slidecraft",
        "title": "SlideCraft",
        "description": (
            "Open a .pptx presentation in SlideCraft's lightweight lossless in-place side-pane editor "
            "(edit text, typography, colors, table cells, layout, and speaker notes and save directly back to the source .pptx). "
            "All arguments are optional."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Absolute path to the .pptx deck."},
                "slide_index": {"type": "integer", "description": "1-based slide number to focus (default 1).", "default": 1},
            },
        },
        "_meta": {
            "ui": {"resourceUri": APP_URI},
            "com.google.antigravity/ui": {
                "entrypoints": [
                    {"type": "aux_pane", "title": "SlideCraft"},
                    {"type": "global", "title": "SlideCraft"},
                ]
            },
        },
    },
    # App-only tools
    {
        "name": "app_load_deck",
        "title": "Load Deck Overview",
        "description": "Load .pptx overview metadata for SlideCraft.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "active_slide": {"type": "integer", "default": 1},
            },
            "required": ["file_path"],
        },
        "_meta": {"ui": {"visibility": ["app"]}},
    },
    {
        "name": "app_get_slide_scene",
        "title": "Get Editable Slide Scene",
        "description": "Return structured editable scene graph for a single slide.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "slide_index": {"type": "integer"},
            },
            "required": ["file_path", "slide_index"],
        },
        "_meta": {"ui": {"visibility": ["app"]}},
    },
    {
        "name": "app_get_thumbnails",
        "title": "Get Slide Thumbnails",
        "description": "Return lightweight scene thumbnails for a batch of slides.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "start": {"type": "integer", "default": 1},
                "count": {"type": "integer", "default": 4},
            },
            "required": ["file_path"],
        },
        "_meta": {"ui": {"visibility": ["app"]}},
    },
    {
        "name": "app_save_slide",
        "title": "Save Slide Edits Losslessly",
        "description": "Apply surgical edits to the target slide and atomically write back to the source .pptx.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "slide_index": {"type": "integer"},
                "ops": {"type": "array", "items": {"type": "object"}},
                "notes": {"type": ["string", "null"]},
            },
            "required": ["file_path", "slide_index", "ops"],
        },
        "_meta": {"ui": {"visibility": ["app"]}},
    },
    {
        "name": "app_slide_op",
        "title": "Reorder or Delete Slide",
        "description": "Move or delete a slide inside the source .pptx file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "action": {"type": "string", "enum": ["move", "delete"]},
                "slide_index": {"type": "integer"},
                "target_index": {"type": "integer"},
            },
            "required": ["file_path", "action", "slide_index"],
        },
        "_meta": {"ui": {"visibility": ["app"]}},
    },
    {
        "name": "app_list_files",
        "title": "List Deck Files",
        "description": "List recent and local .pptx files.",
        "inputSchema": {
            "type": "object",
            "properties": {"directory": {"type": "string"}},
        },
        "_meta": {"ui": {"visibility": ["app"]}},
    },
    {
        "name": "app_open_external",
        "title": "Open Externally",
        "description": "Open the .pptx in the default desktop app or reveal it in the system file manager.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "action": {"type": "string", "enum": ["open", "reveal"], "default": "open"},
            },
            "required": ["file_path"],
        },
        "_meta": {"ui": {"visibility": ["app"]}},
    },
]


def _handle_tool_call(name: str, args: Dict[str, Any]) -> Dict[str, Any]:
    if name == "open_slidecraft":
        fp = (args.get("file_path") or "").strip()
        slide_index = int(args.get("slide_index") or 1)
        if not fp:
            recent = _discover_pptx_files()
            if recent:
                fp = recent[0]["file_path"]
        if not fp or not os.path.isfile(os.path.expanduser(fp)):
            return {
                "content": [{"type": "text", "text": "SlideCraft opened. Select a .pptx deck to edit."}],
                "structuredContent": {"has_deck": False, "recent_files": _discover_pptx_files()[:15]},
            }
        ov = engine.get_deck_overview(fp)
        _record_recent(ov["file_path"], ov["slide_count"])
        ov["active_slide"] = max(1, min(ov["slide_count"], slide_index)) if ov["slide_count"] > 0 else 1
        ov["recent_files"] = _discover_pptx_files()[:15]
        return {
            "content": [{"type": "text", "text": f"Opened `{ov['file_name']}` ({ov['slide_count']} slides) in SlideCraft at slide {ov['active_slide']}."}],
            "structuredContent": ov,
        }

    if name == "app_load_deck":
        ov = engine.get_deck_overview(args["file_path"])
        _record_recent(ov["file_path"], ov["slide_count"])
        act = int(args.get("active_slide") or 1)
        ov["active_slide"] = max(1, min(ov["slide_count"], act)) if ov["slide_count"] > 0 else 1
        ov["recent_files"] = _discover_pptx_files()[:15]
        return {"content": [{"type": "text", "text": f"Loaded {ov['file_name']}"}], "structuredContent": ov}

    if name == "app_get_slide_scene":
        sc = engine.get_slide_scene(args["file_path"], int(args["slide_index"]))
        return {"content": [{"type": "text", "text": f"Slide {sc['index']}"}], "structuredContent": sc}

    if name == "app_get_thumbnails":
        res = engine.get_thumbnails_batch(args["file_path"], int(args.get("start") or 1), int(args.get("count") or 4))
        return {"content": [{"type": "text", "text": "Thumbnails loaded"}], "structuredContent": res}

    if name == "app_save_slide":
        res = engine.save_slide_edits(
            args["file_path"],
            int(args["slide_index"]),
            args.get("ops") or [],
            notes=args.get("notes"),
        )
        res["overview"]["recent_files"] = _discover_pptx_files()[:15]
        return {
            "content": [{"type": "text", "text": f"Saved {res['applied']} edit(s) to {os.path.basename(args['file_path'])}"}],
            "structuredContent": res,
        }

    if name == "app_slide_op":
        ov = engine.apply_deck_slide_op(
            args["file_path"],
            args["action"],
            int(args["slide_index"]),
            int(args["target_index"]) if args.get("target_index") is not None else None,
        )
        ov["recent_files"] = _discover_pptx_files()[:15]
        return {"content": [{"type": "text", "text": f"Applied slide op: {args['action']}"}], "structuredContent": ov}

    if name == "app_list_files":
        files = _discover_pptx_files(args.get("directory"))
        return {"content": [{"type": "text", "text": f"Found {len(files)} files"}], "structuredContent": {"files": files}}

    if name == "app_open_external":
        _open_or_reveal_file(args.get("file_path", ""), args.get("action", "open"))
        return {"content": [{"type": "text", "text": "Opened externally"}], "structuredContent": {"ok": True}}

    raise KeyError(f"Unknown tool: {name}")


def _send(msg: Dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(msg, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def handle_message(msg: Dict[str, Any]) -> None:
    method = msg.get("method")
    req_id = msg.get("id")
    params = msg.get("params") or {}
    if req_id is None:
        return
    try:
        if method == "initialize":
            _send({
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "protocolVersion": params.get("protocolVersion") or "2025-06-18",
                    "capabilities": {"tools": {"listChanged": False}, "resources": {"subscribe": False, "listChanged": False}},
                    "serverInfo": SERVER_INFO,
                },
            })
        elif method == "ping":
            _send({"jsonrpc": "2.0", "id": req_id, "result": {}})
        elif method == "tools/list":
            _send({"jsonrpc": "2.0", "id": req_id, "result": {"tools": TOOLS}})
        elif method == "tools/call":
            res = _handle_tool_call(params.get("name", ""), params.get("arguments") or {})
            _send({"jsonrpc": "2.0", "id": req_id, "result": res})
        elif method == "resources/list":
            _send({
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "resources": [{
                        "uri": APP_URI,
                        "name": "SlideCraft",
                        "description": "Lightweight lossless in-place .pptx slide editor",
                        "mimeType": "text/html;profile=mcp-app",
                        "_meta": {"ui": {"prefersBorder": False}},
                    }]
                },
            })
        elif method == "resources/read":
            if params.get("uri") != APP_URI:
                _send({"jsonrpc": "2.0", "id": req_id, "error": {"code": -32002, "message": "Resource not found"}})
                return
            _send({
                "jsonrpc": "2.0",
                "id": req_id,
                "result": {
                    "contents": [{
                        "uri": APP_URI,
                        "mimeType": "text/html;profile=mcp-app",
                        "text": _build_editor_html(),
                        "_meta": {"ui": {"prefersBorder": False}},
                    }]
                },
            })
        else:
            _send({"jsonrpc": "2.0", "id": req_id, "error": {"code": -32601, "message": f"Method not found: {method}"}})
    except Exception as exc:
        if method == "tools/call":
            _send({"jsonrpc": "2.0", "id": req_id, "result": {"isError": True, "content": [{"type": "text", "text": f"Error: {exc}"}]}})
        else:
            _send({"jsonrpc": "2.0", "id": req_id, "error": {"code": -32603, "message": str(exc)}})


def main() -> None:
    global _DATA_DIR
    ap = argparse.ArgumentParser()
    ap.add_argument("--data_dir", default=_DATA_DIR)
    args, _ = ap.parse_known_args()
    _DATA_DIR = os.path.expanduser(args.data_dir)
    os.makedirs(_DATA_DIR, exist_ok=True)
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(msg, dict):
            handle_message(msg)


if __name__ == "__main__":
    main()

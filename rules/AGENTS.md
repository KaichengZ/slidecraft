# SlideCraft Auto-Open Rule

Whenever you create, edit, convert, or are asked to preview or open a `.pptx` presentation file:

1. Immediately call the `open_slidecraft` tool (from `slidecraft_slidecraft`) with the absolute `file_path` so the deck opens or refreshes automatically in Antigravity's right-hand **SlideCraft** pane for instant visual preview and lossless in-place editing.
2. When modifying a specific slide (for example, in response to a user revision request for slide `N`), pass `slide_index: N` to `open_slidecraft` so the editor jumps straight to the updated slide.
3. Prefer generating editable `.pptx` files using SlideCraft's managed Python environment (`~/.gemini/antigravity/plugin_data/slidecraft/.venv/bin/python` on macOS/Linux or `.venv/Scripts/python.exe` on Windows, with `python-pptx` and `Pillow` pre-installed).

# SlideCraft — In-Place `.pptx` Slide Editor Plugin for Antigravity

**SlideCraft** is an open-source Antigravity plugin (MCP App) that embeds a fast, lightweight, **lossless in-place `.pptx` editor** directly into Antigravity's right-hand auxiliary pane.

Unlike read-only previewers or lossy format converters, SlideCraft maps every element on the slide to its real OOXML `shape_id` and surgically patches only modified shapes, table cells, and speaker notes back into the original `.pptx` file via atomic write-back — preserving untouched slides, masters, charts, and high-resolution media.

## Features

- **60fps Direct Canvas Editing**: Select, drag, and resize text boxes, shapes, images, tables, and charts smoothly with 8-point resize handles and arrow-key nudging (`Shift + Arrow` for 10px steps).
- **Rich Typography & Styling**: Adjust font size (`A-` / numeric input / `A+`), bold, italic, underline, text color, fill color (or transparent fill), and alignment across text boxes, shapes, **and individual table cells**.
- **In-Place Text & Table Cell Editing**: Double-click any text box or table cell to select words and edit text directly on the slide without dragging the container.
- **Lossless Surgical Save (`⌘S` / `Ctrl+S`)**: Writes only changed attributes back to the source `.pptx` atomically via `python-pptx`, with full local undo/redo (`⌘Z` / `⌘⇧Z`).
- **Space-Efficient Side-Pane Layout**: Bottom horizontal slide filmstrip (with wheel scrolling, drag-free slide reordering, and one-click collapse) maximizes slide width in narrow half-screen panes.
- **Resizable Speaker Notes & Presenter Mode**: Drag the top handle of the speaker notes panel to resize it freely in both editor mode and fullscreen presenter mode (`F`), with adjustable notes font size (`A-` / `A+`).
- **Cross-Platform & Bilingual**: Works on macOS, Windows, and Linux with automatic English / Chinese UI detection (`EN / 中文` switcher in toolbar).
- **Built-in Slide Creation Skill**: Includes the `slide-creation` skill for structured outline planning, style exploration, 16:9 `.pptx` generation, and automatic SlideCraft launch.

## Quick Install

### Option 1: One-Command Installer (macOS / Linux / Windows)

```bash
git clone https://github.com/KaichengZ/slidecraft.git
cd slidecraft
python3 install.py
```

Or download `slidecraft-v1.0.0.zip` from [GitHub Releases](https://github.com/KaichengZ/slidecraft/releases/latest), unzip it, and run `python3 install.py`.

This copies SlideCraft to `~/.gemini/config/plugins/slidecraft/` and initializes its isolated Python virtual environment (`python-pptx` + `Pillow`) inside `~/.gemini/antigravity/plugin_data/slidecraft/.venv/`.

### Option 2: Manual Install

1. Copy the `slidecraft/` folder into `~/.gemini/config/plugins/slidecraft/`.
2. Antigravity discovers the plugin automatically (no manual registration needed).
3. On first launch, `server/launcher.py` automatically creates its virtual environment and installs `requirements.txt` (`python-pptx`, `Pillow`).

## Usage

- **Automatic**: Ask Antigravity to create or edit any `.pptx` presentation — SlideCraft opens automatically in the right-hand pane.
- **Manual**: Click the **`+`** menu in Antigravity's right-hand auxiliary pane (or the left sidebar) and choose **SlideCraft** to open any local `.pptx` file.

## Keyboard Shortcuts

| Shortcut | Action |
| :--- | :--- |
| `⌘S` / `Ctrl+S` | Save edits losslessly to source `.pptx` |
| `⌘Z` / `Ctrl+Z` | Undo |
| `⌘⇧Z` / `Ctrl+Shift+Z` | Redo |
| `←` / `→` / `Space` | Previous / Next slide (or nudge selected element) |
| `Delete` / `Backspace` | Delete selected element |
| `F` | Toggle fullscreen presenter view |
| `N` | Toggle speaker notes panel |
| `Esc` | Exit text editing / deselect / exit fullscreen |

## License

MIT

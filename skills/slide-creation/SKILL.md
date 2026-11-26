---
name: slide-creation
description: >-
  End-to-end presentation creation and iterative editing workflow: plan a
  slide-by-slide narrative outline, propose 2-3 cohesive visual style
  directions, build clean 16:9 editable .pptx decks with python-pptx (charts,
  tables, KPI cards, speaker notes), automatically open in the right-hand
  SlideCraft editor pane, and support fast per-slide revisions. Use whenever
  the user asks to create, design, redesign, or polish a presentation (.pptx).
---

# SlideCraft Presentation Workflow

This skill guides the end-to-end creation of executive-ready, fully editable `.pptx` slide decks paired with **SlideCraft's** right-hand interactive editor.

## Python Environment

SlideCraft manages a dedicated virtual environment with `python-pptx` and `Pillow` installed:
- **macOS / Linux**: `~/.gemini/antigravity/plugin_data/slidecraft/.venv/bin/python`
- **Windows**: `%USERPROFILE%\.gemini\antigravity\plugin_data\slidecraft\.venv\Scripts\python.exe`

Fall back to `python3` if `python-pptx` is already installed globally.

## Five-Step Workflow

### 1. Narrative & Slide-by-Slide Outline
Clarify the audience, goal, and slide count if ambiguous (skip unnecessary questions when the user already provided rich source material):
- Use **action titles** (each slide title states one clear takeaway rather than a generic topic label).
- Typical executive/review flow: Cover -> Executive Summary / KPIs -> Deep-Dive Sections with Visuals -> Action Plan / Roadmap -> Closing / Q&A.

### 2. Visual Direction (2-3 Style Options)
Unless the user asks to generate immediately or provides brand guidelines, propose 2-3 concise visual directions before building the full deck:
1. **Executive Clean**: Deep Navy `#0F172A` + Cobalt `#2563EB` on crisp `#FFFFFF` / `#F8FAFC` surfaces.
2. **Modern Dark / Keynote**: Midnight `#0B1120` + Slate `#1E293B` cards + Cyan `#06B6D4` / Indigo `#6366F1` accents.
3. **Editorial Warm**: Warm Paper `#FAF7F2` + Terracotta `#C2410C` + Charcoal `#1C1917`.

### 3. Build Clean 16:9 Editable `.pptx`
Write a clean, reusable Python builder script using `python-pptx`:
- **Canvas**: 16:9 widescreen (`13.333" × 7.5"`), blank layout (`slide_layouts[6]`), margins `>= 0.6"`.
- **Typography**: Consistent cross-platform sans-serif stack; title `22-26 pt` bold, card headings `15-17 pt` bold, body `13-15 pt` (never `< 11 pt`), KPI callouts `30-42 pt` bold.
- **Visual Variety**: Never stack consecutive walls of bullets — alternate between KPI cards, two-column text + native chart, structured comparison tables, process timelines, and section dividers.
- **Speaker Notes**: Populate `slide.notes_slide.notes_text_frame.text` with concise talking points on every content slide.

### 4. Auto-Open in SlideCraft
Immediately call `open_slidecraft(file_path=...)` so the deck renders in the right-hand **SlideCraft** pane. Summarize the saved path, slide count, and remind the user they can tweak typography, colors, layout, table cells, or speaker notes directly in the pane and save with `⌘S` / `Ctrl+S`.

### 5. Iterative Per-Slide Refinement
When the user requests changes to specific slides, update the builder script (or patch the deck cleanly), regenerate, and call `open_slidecraft(file_path=..., slide_index=N)` to jump straight to the updated slide.

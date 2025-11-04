# QuickMacro
Quick Macro to record and replay mouse and keyboard actions (Tkinter + pynput).

## Setup
- Python 3.8+ is recommended.
- Install dependencies: `pip install -r requirements.txt`

## Run
- Start the GUI: `python QuickMacro.py`
- Record: click `Start recording` or press `F10` to start immediately (ESC/F10 to stop)
- Replay: select an `.action` from dropdown, set repeat times, click `Start replaying` or press `F11` to start immediately (ESC/F11 to stop)
- Hotkeys: `F10` start/stop recording, `F11` start/stop replaying

## UI Theme
- Switched to a clean, professional (business) ttk theme using native look where possible.
- Fonts prefer `Segoe UI`/`Microsoft YaHei UI`/`微软雅黑`/`Arial`.
- Neutral colors, clear spacing, and consistent sizes for a focused workflow.

## Assets
- Put optional images in `assets/` to customize:
  - `assets/icon.png` for the window icon
  - `assets/bg.png` decorative background (business theme does not auto-apply a background image)

## Action Files (.action)
- When recording, the app now creates a single `.action` file in the `actions/` folder using a timestamp name like `YYYYMMDD-HHMMSS.action`.
- The format is a simple, line-based pseudo language with timing:
  - `# QuickMacro action v1` header
  - `META SCREEN <w> <h>` record-time screen
  - Keyboard: `K DOWN <vk> <ms>`, `K UP <vk> <ms>`
  - Mouse move: `M MOVE <x> <y> [<nx> <ny>] <ms>`
  - Mouse click: `M CLICK <left|right> <DOWN|UP> <x> <y> [<nx> <ny>] <ms>`
  - Mouse scroll: `M SCROLL <dx> <dy> <ms>`
- In the GUI, select which `.action` to replay from the dropdown (it lists files from `actions/`). The latest recording is auto-selected when you start recording.

## Fixes in this update
- Execute phase UI is now coordinated centrally to avoid early reset when mouse and keyboard speeds differ.
- When only replaying mouse, the replay button text updates correctly after countdown.
- Mouse recording now stores normalized coordinates and a meta line with the record-time screen size; playback scales to the current DPI/screen size and moves to the exact click point before pressing.
 - Record/Replay now always include both mouse and keyboard; options removed.
 - To maximize precision, playback prefers recorded raw pixel coordinates when resolution is unchanged; it falls back to normalized coordinates only when resolution/DPI changed.
 - On stopping/exit, any keys or mouse buttons still held down by the macro are safely released to prevent auto-repeat.
 - Window size stabilized by setting DPI awareness at startup; background threads are daemonized and window close exits cleanly.
 - Removed all countdown flows; actions start immediately via button or hotkey.

## Known / Future Work
- Refactoring still needed to decouple UI, controller and workers more cleanly.
- Consider timing-aware replay (timestamps) to better preserve human rhythm.
- Multi-monitor: positions are recorded relative to the virtual desktop. If monitor arrangement changes between record and replay, positions can shift; recording per-monitor id is a potential improvement.

# QuickMacro
Quick Macro to record and replay mouse and keyboard actions (Tkinter + pynput).

## Setup
- Python 3.8+ is recommended.
- Install dependencies: `pip install -r requirements.txt`

## Run
- Start the GUI: `python QuickMacro.py`
- Record: set the countdown and click `Start recording` (ESC to stop)
- Replay: choose what to replay (mouse/keyboard), set countdown and repeat times, click `Start replaying` (ESC to stop)

## Fixes in this update
- Execute phase UI is now coordinated centrally to avoid early reset when mouse and keyboard speeds differ.
- When only replaying mouse, the replay button text updates correctly after countdown.
- Mouse recording now stores normalized coordinates and a meta line with the record-time screen size; playback scales to the current DPI/screen size and moves to the exact click point before pressing.

## Known / Future Work
- Refactoring still needed to decouple UI, controller and workers more cleanly.
- Consider timing-aware replay (timestamps) to better preserve human rhythm.
- Multi-monitor: positions are recorded relative to the virtual desktop. If monitor arrangement changes between record and replay, positions can shift; recording per-monitor id is a potential improvement.

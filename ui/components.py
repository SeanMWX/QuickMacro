import os
import time
from datetime import datetime
import tkinter
from tkinter import ttk
from tkinter import scrolledtext


class LogPanel:
    """Encapsulate log text area and helpers."""
    def __init__(self, parent, font_family):
        self.frame = parent
        self.text = scrolledtext.ScrolledText(self.frame, wrap='word', height=8, state='disabled', font=(font_family, 10))
        self.text.place(x=12, y=36, width=636, height=130)

    def log_event(self, msg: str):
        try:
            ts = datetime.now().strftime('%H:%M:%S')
            self.text.configure(state='normal')
            self.text.insert('end', f"[{ts}] {msg}\n")
            self.text.see('end')
            self.text.configure(state='disabled')
        except Exception:
            pass


class MonitorPanel:
    """Encapsulate monitor labels and state update callbacks."""
    def __init__(self, parent, font_family, state):
        self.state = state
        self.frame = parent
        self.loop_label = ttk.Label(parent, text='Loops: 0/0', style='CardLabel.TLabel')
        self.loop_label.place(x=12, y=40, width=200, height=24)
        self.time_label = ttk.Label(parent, text='Current loop time: 0.0s', style='CardLabel.TLabel')
        self.time_label.place(x=220, y=40, width=220, height=24)
        self.total_label = ttk.Label(parent, text='Total loop time: 0.0s', style='CardLabel.TLabel')
        self.total_label.place(x=460, y=40, width=180, height=24)
        self.dungeon_label = ttk.Label(parent, text='Dungeon time: 0.0s', style='CardLabel.TLabel')
        self.dungeon_label.place(x=12, y=70, width=250, height=24)
        self.restart_label = ttk.Label(parent, text='Restart time: 3600.0s', style='CardLabel.TLabel')
        self.restart_label.place(x=280, y=70, width=260, height=24)

    def update_labels(self):
        try:
            loops_text = f"Loops: {self.state.monitor_completed_loops}/{self.state.monitor_total_loops}" if self.state.monitor_total_loops else "Loops: 0/0"
            self.loop_label['text'] = loops_text
            elapsed = 0.0
            if self.state.monitor_loop_start_ts is not None:
                elapsed = max(0.0, time.monotonic() - self.state.monitor_loop_start_ts)
            self.time_label['text'] = f"Current loop time: {elapsed:.1f}s"
            self.total_label['text'] = f"Total loop time: {self.state.monitor_total_time_s:.1f}s"
            dungeon_elapsed = 0.0
            if self.state.dungeon_start_ts is not None:
                dungeon_elapsed = max(0.0, time.monotonic() - self.state.dungeon_start_ts)
            self.dungeon_label['text'] = f"Dungeon time: {dungeon_elapsed:.1f}s"
            restart_ms = getattr(self.state, 'restart_timeout_ms', 0) or 0
            self.restart_label['text'] = f"Restart time: {restart_ms/1000.0:.1f}s"
        except Exception:
            pass

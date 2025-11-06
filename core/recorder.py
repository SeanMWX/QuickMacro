import threading
import time
from typing import Optional
from pynput import keyboard, mouse
from pynput.keyboard import Key
from core.utils import get_screen_size


class _Writer:
    def __init__(self, path: str):
        self.path = path
        self._lock = threading.Lock()

    def write_line(self, line: str):
        with self._lock:
            with open(self.path, 'a', encoding='utf-8') as f:
                f.write(line + "\n")


class KeyboardRecorder(threading.Thread):
    def __init__(self, writer: _Writer, stop_event: threading.Event, start_monotonic: float):
        super().__init__()
        self.daemon = True
        self.writer = writer
        self.stop_event = stop_event
        self.start_mono = start_monotonic

    def run(self):
        def on_press(key):
            if key in (keyboard.Key.esc, keyboard.Key.f10, keyboard.Key.f11):
                return
            t = int((time.monotonic() - self.start_mono) * 1000)
            try:
                vk = key.vk
            except AttributeError:
                vk = key.value.vk
            self.writer.write_line(f"K DOWN {vk} {t}")

        def on_release(key):
            if self.stop_event.is_set():
                kb.stop(); return False
            if key in (keyboard.Key.f10, keyboard.Key.f11):
                return
            t = int((time.monotonic() - self.start_mono) * 1000)
            try:
                vk = key.vk
            except AttributeError:
                vk = key.value.vk
            self.writer.write_line(f"K UP {vk} {t}")

        with keyboard.Listener(on_press=on_press, on_release=on_release) as kb:
            kb.join()


class MouseRecorder(threading.Thread):
    def __init__(self, writer: _Writer, stop_event: threading.Event, start_monotonic: float):
        super().__init__()
        self.daemon = True
        self.writer = writer
        self.stop_event = stop_event
        self.start_mono = start_monotonic
        self.sw, self.sh = get_screen_size()

    def run(self):
        def on_move(x, y):
            if self.stop_event.is_set():
                ms.stop(); return
            t = int((time.monotonic() - self.start_mono) * 1000)
            try:
                nx = float(x) / float(self.sw)
                ny = float(y) / float(self.sh)
                self.writer.write_line(f"M MOVE {int(x)} {int(y)} {nx:.6f} {ny:.6f} {t}")
            except Exception:
                self.writer.write_line(f"M MOVE {int(x)} {int(y)} {t}")

        def on_click(x, y, button, pressed):
            if self.stop_event.is_set():
                ms.stop(); return
            t = int((time.monotonic() - self.start_mono) * 1000)
            btn = 'left' if button == mouse.Button.left else 'right'
            act = 'DOWN' if pressed else 'UP'
            try:
                nx = float(x) / float(self.sw)
                ny = float(y) / float(self.sh)
                self.writer.write_line(f"M CLICK {btn} {act} {int(x)} {int(y)} {nx:.6f} {ny:.6f} {t}")
            except Exception:
                self.writer.write_line(f"M CLICK {btn} {act} {int(x)} {int(y)} {t}")

        def on_scroll(x, y, dx, dy):
            if self.stop_event.is_set():
                ms.stop(); return
            t = int((time.monotonic() - self.start_mono) * 1000)
            self.writer.write_line(f"M SCROLL {int(dx)} {int(dy)} {t}")

        with mouse.Listener(on_move=on_move, on_click=on_click, on_scroll=on_scroll) as ms:
            ms.join()


class Recorder:
    def __init__(self, path: str, stop_event: threading.Event):
        self.path = path
        self.stop_event = stop_event
        self._writer = _Writer(path)
        self._start_mono = time.monotonic()
        self.kb_thread: Optional[KeyboardRecorder] = None
        self.ms_thread: Optional[MouseRecorder] = None

    def start(self):
        self.kb_thread = KeyboardRecorder(self._writer, self.stop_event, self._start_mono)
        self.ms_thread = MouseRecorder(self._writer, self.stop_event, self._start_mono)
        self.kb_thread.start()
        self.ms_thread.start()

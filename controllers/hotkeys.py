import time
import threading
from pynput import keyboard

class HotkeyController(threading.Thread):
    def __init__(self, state, root, on_toggle_record, on_toggle_replay):
        super().__init__()
        self.daemon = True
        self.state = state
        self.root = root
        self.on_toggle_record = on_toggle_record
        self.on_toggle_replay = on_toggle_replay

    def run(self):
        last_f10 = 0.0
        last_f11 = 0.0

        def on_press(key):
            nonlocal last_f10, last_f11
            now = time.time()
            try:
                if key == keyboard.Key.f10:
                    if now - last_f10 > 0.3:
                        last_f10 = now
                        self.on_toggle_record()
                elif key == keyboard.Key.f11:
                    if now - last_f11 > 0.3:
                        last_f11 = now
                        self.on_toggle_replay()
            except Exception:
                pass

        with keyboard.Listener(on_press=on_press) as hk:
            hk.join()

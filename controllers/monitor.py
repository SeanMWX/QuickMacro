import threading
import time
import cv2
import numpy as np
import pyautogui

class MonitorThread(threading.Thread):
    def __init__(self, target_path: str, timeout_s: float, interval_s: float, stop_callbacks=None, restart_callback=None, hit_callback=None):
        super().__init__()
        self.daemon = True
        self.target_path = target_path
        self.timeout_s = timeout_s
        self.interval_s = interval_s
        self.stop_callbacks = stop_callbacks or []
        self.restart_callback = restart_callback
        self._stop_ev = threading.Event()
        self._tmpl = self._load_template(target_path)
        self.hit_callback = hit_callback

    def _load_template(self, path):
        try:
            img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
            return img
        except Exception:
            return None

    def stop(self):
        try:
            self._stop_ev.set()
        except Exception:
            pass

    def run(self):
        if self._tmpl is None:
            return
        last_hit = time.monotonic()
        while not self._stop_ev.is_set():
            if (time.monotonic() - last_hit) >= self.timeout_s:
                for cb in self.stop_callbacks:
                    try:
                        cb()
                    except Exception:
                        pass
                try:
                    if callable(self.restart_callback):
                        self.restart_callback()
                except Exception:
                    pass
                last_hit = time.monotonic()
            try:
                shot = pyautogui.screenshot()
                shot = cv2.cvtColor(np.array(shot), cv2.COLOR_RGB2GRAY)
                res = cv2.matchTemplate(shot, self._tmpl, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, _ = cv2.minMaxLoc(res)
                if max_val >= 0.8:
                    last_hit = time.monotonic()
                    if callable(self.hit_callback):
                        try:
                            self.hit_callback()
                        except Exception:
                            pass
            except Exception:
                pass
            try:
                self._stop_ev.wait(self.interval_s)
            except Exception:
                pass

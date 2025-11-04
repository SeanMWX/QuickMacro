import time
import ctypes
import tkinter


def wait_until_or_stop(until_ts: float, stop_event, quantum: float = 0.01) -> bool:
    """Cooperative wait using monotonic clock. Returns True if time reached, False if stopped."""
    try:
        while True:
            if stop_event.is_set():
                return False
            now = time.monotonic()
            if now >= until_ts:
                return True
            time.sleep(min(quantum, max(0.0, until_ts - now)))
    except Exception:
        return False


def get_screen_size():
    try:
        user32 = ctypes.windll.user32
        return int(user32.GetSystemMetrics(0)), int(user32.GetSystemMetrics(1))
    except Exception:
        try:
            t = tkinter.Tk()
            w = int(t.winfo_screenwidth())
            h = int(t.winfo_screenheight())
            t.destroy()
            return w, h
        except Exception:
            return 1920, 1080


import threading
import time
from pynput import keyboard

def _safe_set(ev):
    try:
        ev.set()
    except Exception:
        pass

class ListenController(threading.Thread):
    def __init__(self, state, ui_refs=None, on_escape=None):
        super().__init__()
        self.daemon = True
        self.state = state
        self.ui_refs = ui_refs
        self.on_escape = on_escape

    def run(self):
        try:
            self.state.ev_stop_listen.clear()
        except Exception:
            pass

        def on_release(key):
            if key == keyboard.Key.esc:
                _safe_set(self.state.ev_stop_listen)
                self.state.can_start_listening = True
                self.state.can_start_executing = True
                try:
                    if self.ui_refs:
                        self.ui_refs.update_ui_for_state('idle')
                except Exception:
                    pass
                if callable(self.on_escape):
                    try:
                        self.on_escape()
                    except Exception:
                        pass
                listener.stop()

        with keyboard.Listener(on_release=on_release) as listener:
            listener.join()

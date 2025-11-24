import threading
import time
import json
from pynput.keyboard import Controller as KeyBoardController, KeyCode
from pynput.mouse import Button, Controller as MouseController
from core.actions import parse_action_line
from core.utils import wait_until_or_stop, get_screen_size

# Optional: DirectInput-style relative mouse (works better in many games)
try:
    import pydirectinput as _pdi
    _PDI_AVAILABLE = True
    try:
        _pdi.PAUSE = 0
        _pdi.FAILSAFE = False
    except Exception:
        pass
except Exception:
    _PDI_AVAILABLE = False


class Replayer:
    def __init__(self, path, stop_event_kb, stop_event_ms, infinite_event, repeat_count: int, use_relative_mouse: bool = False, relative_gain: float = 1.0, auto_detect: bool = False, progress_cb=None, loop_start_cb=None):
        self.path = path
        self.stop_event_kb = stop_event_kb
        self.stop_event_ms = stop_event_ms
        self.infinite_event = infinite_event
        self.repeat_count = repeat_count
        self.total_loops = repeat_count if repeat_count else 1
        self.use_relative_mouse = use_relative_mouse
        self.relative_gain = float(relative_gain or 1.0)
        self.auto_detect = bool(auto_detect)
        self.progress_cb = progress_cb
        self.loop_start_cb = loop_start_cb
        self._runner = None
        self._events = self._load_events(path)

    def _load_events(self, path):
        events = []
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith('#'):
                        continue
                    d = parse_action_line(s)
                    t_ms = int(d.get('ms') or 0) if d.get('ms') not in (None, '') else 0
                    events.append((t_ms, d))
        except Exception:
            pass
        events.sort(key=lambda x: x[0])
        return events

    def _dispatch_vk(self, kb, d):
        try:
            vk = int(d.get('vk') or 0)
            if d.get('op') == 'DOWN':
                kb.press(KeyCode.from_vk(vk))
            elif d.get('op') == 'UP':
                kb.release(KeyCode.from_vk(vk))
        except Exception:
            pass

    def _dispatch_ms(self, ms_ctrl, d, screen_wh, rel_state):
        cw, ch = screen_wh
        rw, rh = rel_state.get('rw_ch', (cw, ch))
        op = (d.get('op') or '').upper()
        if op == 'MOVE':
            try:
                x = int(d.get('x') or 0); y = int(d.get('y') or 0)
            except Exception:
                x = y = 0
            nx = d.get('nx'); ny = d.get('ny')
            use_norm = False
            try:
                if rw and rh and (abs(cw - rw)/float(rw) > 0.02 or abs(ch - rh)/float(rh) > 0.02):
                    use_norm = (nx not in (None,'') and ny not in (None,''))
            except Exception:
                use_norm = False
            tx = int(round(float(nx) * float(cw))) if use_norm else x
            ty = int(round(float(ny) * float(ch))) if use_norm else y
            if self.use_relative_mouse:
                # relative: use recorded deltas with gain
                prev = rel_state.get('prev')
                if prev is None:
                    rel_state['prev'] = (x, y)
                    return
                px, py = prev
                fdx = (x - px) * self.relative_gain + rel_state.get('resx', 0.0)
                fdy = (y - py) * self.relative_gain + rel_state.get('resy', 0.0)
                dx = int(round(fdx)); dy = int(round(fdy))
                rel_state['resx'] = fdx - dx
                rel_state['resy'] = fdy - dy
                if dx != 0 or dy != 0:
                    if _PDI_AVAILABLE:
                        try:
                            _pdi.moveRel(dx, dy, duration=0, relative=True)
                        except Exception:
                            try:
                                ms_ctrl.move(dx, dy)
                            except Exception:
                                pass
                    else:
                        try:
                            ms_ctrl.move(dx, dy)
                        except Exception:
                            pass
                rel_state['prev'] = (x, y)
            else:
                try:
                    ms_ctrl.position = (tx, ty)
                except Exception:
                    pass
        elif op == 'CLICK':
            btn = d.get('btn') or 'left'; act = d.get('act') or 'DOWN'
            try:
                x = int(d.get('x') or 0); y = int(d.get('y') or 0)
            except Exception:
                x = y = 0
            nx = d.get('nx'); ny = d.get('ny')
            use_norm = False
            try:
                if rw and rh and (abs(cw - rw)/float(rw) > 0.02 or abs(ch - rh)/float(rh) > 0.02):
                    use_norm = (nx not in (None,'') and ny not in (None,''))
            except Exception:
                use_norm = False
            tx = int(round(float(nx) * float(cw))) if use_norm else x
            ty = int(round(float(ny) * float(ch))) if use_norm else y
            try:
                ms_ctrl.position = (tx, ty)
            except Exception:
                pass
            if act == 'DOWN':
                if btn == 'left': ms_ctrl.press(Button.left)
                else: ms_ctrl.press(Button.right)
            else:
                if btn == 'left': ms_ctrl.release(Button.left)
                else: ms_ctrl.release(Button.right)
        elif op == 'SCROLL':
            try:
                dx = int(d.get('dx') or 0); dy = int(d.get('dy') or 0)
            except Exception:
                dx = dy = 0
            try:
                ms_ctrl.scroll(dx, dy)
            except Exception:
                pass

    def _run(self):
        kb = KeyBoardController()
        ms_ctrl = MouseController()
        loop_idx = 0
        while True:
            if self.stop_event_kb.is_set() or self.stop_event_ms.is_set():
                break
            # loop bookkeeping
            loop_idx += 1
            try:
                if callable(self.loop_start_cb):
                    self.loop_start_cb(loop_idx, self.total_loops)
            except Exception:
                pass
            cw, ch = get_screen_size()
            rel_state = {'prev': None, 'resx': 0.0, 'resy': 0.0, 'rw_ch': (cw, ch)}
            start_ts = time.monotonic()
            for t_ms, d in self._events:
                if self.stop_event_kb.is_set() or self.stop_event_ms.is_set():
                    break
                target = start_ts + (t_ms/1000.0)
                if not wait_until_or_stop(target, self.stop_event_kb):
                    break
                if d.get('type') == 'VK':
                    self._dispatch_vk(kb, d)
                elif d.get('type') == 'MS':
                    self._dispatch_ms(ms_ctrl, d, (cw, ch), rel_state)
            try:
                if callable(self.progress_cb):
                    self.progress_cb(loop_idx, self.total_loops)
            except Exception:
                pass
            if self.stop_event_kb.is_set() or self.stop_event_ms.is_set():
                break
            if self.infinite_event.is_set():
                continue
            self.repeat_count -= 1
            if self.repeat_count <= 0:
                break
        try:
            self.stop_event_kb.set()
            self.stop_event_ms.set()
        except Exception:
            pass

    def start(self):
        if self._runner and self._runner.is_alive():
            return
        self._runner = threading.Thread(target=self._run, daemon=True)
        self._runner.start()

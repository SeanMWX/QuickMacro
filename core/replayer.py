import threading
import time
import json
from pynput.keyboard import Controller as KeyBoardController, KeyCode
from pynput.mouse import Button, Controller as MouseController
from core.actions import parse_action_line
from core.utils import wait_until_or_stop, get_screen_size


class KeyboardReplayer(threading.Thread):
    def __init__(self, path, stop_event, infinite_event, repeat_count: int):
        super().__init__()
        self.daemon = True
        self.path = path
        self.stop_event = stop_event
        self.infinite_event = infinite_event
        self.repeat_count = repeat_count
        self._pressed = set()

    def run(self):
        while True:
            if self.stop_event.is_set():
                return
            try:
                with open(self.path, 'r', encoding='utf-8') as f:
                    kb = KeyBoardController()
                    start_ts = time.monotonic()
                    for line in f:
                        if self.stop_event.is_set():
                            break
                        s = line.strip()
                        if not s:
                            continue
                        d = parse_action_line(s)
                        if d.get('type') == 'VK':
                            try:
                                t_ms = int(d.get('ms') or 0)
                            except Exception:
                                t_ms = 0
                            target = start_ts + (t_ms/1000.0)
                            if not wait_until_or_stop(target, self.stop_event):
                                break
                            try:
                                vk = int(d.get('vk') or 0)
                                if d.get('op') == 'DOWN':
                                    kb.press(KeyCode.from_vk(vk)); self._pressed.add(vk)
                                elif d.get('op') == 'UP':
                                    kb.release(KeyCode.from_vk(vk)); self._pressed.discard(vk)
                            except Exception:
                                pass
                        else:
                            # legacy JSON support
                            try:
                                obj = json.loads(line)
                                if obj.get('name') == 'keyboard':
                                    if obj['event'] == 'press':
                                        vk = obj['vk']; kb.press(KeyCode.from_vk(vk)); self._pressed.add(vk)
                                    elif obj['event'] == 'release':
                                        vk = obj['vk']; kb.release(KeyCode.from_vk(vk)); self._pressed.discard(vk)
                            except Exception:
                                pass
            finally:
                try:
                    kb = KeyBoardController()
                    for vk in list(self._pressed):
                        try:
                            kb.release(KeyCode.from_vk(vk))
                        except Exception:
                            pass
                        self._pressed.discard(vk)
                except Exception:
                    pass
            if self.infinite_event.is_set():
                continue
            self.repeat_count -= 1
            if self.repeat_count <= 0:
                try:
                    self.stop_event.set()
                except Exception:
                    pass
                return


class MouseReplayer(threading.Thread):
    def __init__(self, path, stop_event, infinite_event, repeat_count: int):
        super().__init__()
        self.daemon = True
        self.path = path
        self.stop_event = stop_event
        self.infinite_event = infinite_event
        self.repeat_count = repeat_count
        self._pressed = set()

    def run(self):
        while True:
            if self.stop_event.is_set():
                return
            try:
                with open(self.path, 'r', encoding='utf-8') as f:
                    ms = MouseController()
                    cw, ch = get_screen_size()
                    rw, rh = cw, ch
                    start_ts = time.monotonic()
                    for line in f:
                        if self.stop_event.is_set():
                            break
                        s = line.strip()
                        if not s:
                            continue
                        d = parse_action_line(s)
                        if d.get('type') == 'META' and d.get('op') == 'SCREEN':
                            try:
                                rw = int(d.get('x') or cw); rh = int(d.get('y') or ch)
                            except Exception:
                                rw, rh = cw, ch
                            continue
                        if d.get('type') != 'MS':
                            # legacy JSON support
                            try:
                                obj = json.loads(line)
                                if obj.get('name') == 'meta':
                                    rw = int(obj['screen']['w']); rh = int(obj['screen']['h'])
                                elif obj.get('name') == 'mouse':
                                    # no timing for legacy
                                    if obj['event'] == 'move':
                                        ms.position = (int(obj['location']['x']), int(obj['location']['y']))
                                    elif obj['event'] == 'click':
                                        if obj['action']:
                                            (ms.press(Button.left) if obj['target']=='left' else ms.press(Button.right))
                                        else:
                                            (ms.release(Button.left) if obj['target']=='left' else ms.release(Button.right))
                                    elif obj['event'] == 'scroll':
                                        ms.scroll(obj['location']['x'], obj['location']['y'])
                            except Exception:
                                pass
                            continue

                        # timing
                        try:
                            t_ms = int(d.get('ms') or 0)
                        except Exception:
                            t_ms = 0
                        target = start_ts + (t_ms/1000.0)
                        if not wait_until_or_stop(target, self.stop_event):
                            break

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
                            ms.position = (tx, ty)
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
                                ms.position = (tx, ty)
                            except Exception:
                                pass
                            if act == 'DOWN':
                                if btn == 'left': ms.press(Button.left); self._pressed.add('left')
                                else: ms.press(Button.right); self._pressed.add('right')
                            else:
                                if btn == 'left': ms.release(Button.left); self._pressed.discard('left')
                                else: ms.release(Button.right); self._pressed.discard('right')
                        elif op == 'SCROLL':
                            try:
                                dx = int(d.get('dx') or 0); dy = int(d.get('dy') or 0)
                            except Exception:
                                dx = dy = 0
                            ms.scroll(dx, dy)
            finally:
                try:
                    ms = MouseController()
                    if 'left' in self._pressed:
                        ms.release(Button.left); self._pressed.discard('left')
                    if 'right' in self._pressed:
                        ms.release(Button.right); self._pressed.discard('right')
                except Exception:
                    pass
            if self.infinite_event.is_set():
                continue
            self.repeat_count -= 1
            if self.repeat_count <= 0:
                try:
                    self.stop_event.set()
                except Exception:
                    pass
                return


class Replayer:
    def __init__(self, path, stop_event_kb, stop_event_ms, infinite_event, repeat_count: int):
        self.path = path
        self.stop_event_kb = stop_event_kb
        self.stop_event_ms = stop_event_ms
        self.infinite_event = infinite_event
        self.repeat_count = repeat_count
        self.kb_thread = None
        self.ms_thread = None

    def start(self):
        self.kb_thread = KeyboardReplayer(self.path, self.stop_event_kb, self.infinite_event, self.repeat_count)
        self.ms_thread = MouseReplayer(self.path, self.stop_event_ms, self.infinite_event, self.repeat_count)
        self.kb_thread.start()
        self.ms_thread.start()


import json
import logging
import sys
import logging
import ctypes
import os
import glob
from datetime import datetime
from pathlib import Path
import threading
import time
import tkinter
from tkinter import ttk
import tkinter.font as tkfont
import tkinter.messagebox as messagebox
import cv2
import numpy as np
import pyautogui

from pynput import keyboard, mouse
from pynput.keyboard import Controller as KeyBoardController, KeyCode, Key
from pynput.mouse import Button, Controller as MouseController
from core.actions import parse_action_line as action_parse_line, compose_action_line as action_compose_line, extract_meta as action_extract_meta
from core.replayer import Replayer
from core.recorder import Recorder
from core.utils import get_screen_size
from core import settings as settings_mod

######################################################################
# Helpers
######################################################################
def wait_until_or_stop(until_ts, stop_event, quantum=0.01):
    """使用单调时钟的可中断等待：到达时间返回 True，途中停止返回 False"""
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
def set_process_dpi_aware():
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

def _win_message_box(title, text, flags=0x40):
    try:
        ctypes.windll.user32.MessageBoxW(None, str(text), str(title), flags)
    except Exception:
        pass

def _is_running_as_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False

def relaunch_as_admin_if_needed():
    """Windows: 若非管理员尝试以管理员重启自身；失败时提示用户手动以管理员运行。"""
    try:
        if os.name != 'nt':
            return
        if _is_running_as_admin():
            logging.info('Admin check: running as administrator.')
            return
        script = os.path.abspath(__file__)
        params = ' '.join([f'"{script}"'] + [f'"{a}"' for a in sys.argv[1:]])
        logging.info('Admin check: attempting UAC elevation...')
        rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        if rc > 32:
            os._exit(0)
        else:
            logging.warning(f'UAC elevation failed, rc={rc}.')
            _win_message_box('需要管理员权限', '未能自动获取管理员权限，请右键以管理员方式运行 QuickMacro。\n否则游戏内全局热键与鼠标可能失效。', 0x30)
    except Exception as e:
        try:
            logging.exception('UAC elevation exception: %s', e)
        except Exception:
            pass

def release_all_inputs():
    # Release any keys/buttons that might have been left pressed
    global pressed_vks, pressed_mouse_buttons
    try:
        kb = KeyBoardController()
        for vk in list(pressed_vks):
            try:
                kb.release(KeyCode.from_vk(vk))
            except Exception:
                pass
            pressed_vks.discard(vk)
    except Exception:
        pass

def ensure_actions_dir():
    p = Path('actions')
    p.mkdir(parents=True, exist_ok=True)
    return p

def list_action_files():
    # Prefer files under actions/; if none, fall back to root
    actions_dir = ensure_actions_dir()
    files = sorted(str(f.name) for f in actions_dir.glob('*.action'))
    if files:
        return files
    return sorted(glob.glob('*.action'))

def init_new_action_file():
    # Create a new action file with timestamp-based name and header/meta
    global action_file_name, record_start_time
    ts = datetime.now().strftime('%Y%m%d-%H%M%S')
    actions_dir = ensure_actions_dir()
    action_file_name = str(actions_dir / f"{ts}.action")
    sw, sh = get_screen_size()
    record_start_time = time.time()
    try:
        with open(action_file_name, 'w', encoding='utf-8') as f:
            f.write('# QuickMacro action v1\n')
            f.write(f"META SCREEN {sw} {sh}\n")
            f.write(f"META START {ts}\n")
    except Exception:
        pass

# Legacy recording writer removed; core.recorder handles writing

def ensure_assets_dir():
    p = Path('assets')
    try:
        if not p.exists():
            p.mkdir(parents=True, exist_ok=True)
        readme = p / 'README.txt'
        if not readme.exists():
            readme.write_text(
                'Place optional UI images here:\n'
                '- bg.png   : window background (PNG)\n'
                '- icon.png : window icon (PNG)\n'
                '- monitor_target.png : image to detect (template matching)\n'
                'Replace these with your own cute/moe assets.\n',
                encoding='utf-8'
            )
    except Exception:
        pass

######################################################################
# Settings persistence (delegate to core.settings)
######################################################################
SETTINGS_PATH = 'settings.json'

def load_settings():
    return settings_mod.load_settings(SETTINGS_PATH)

def compute_action_total_ms(path: str) -> int:
    try:
        max_ms = 0
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                s = line.strip()
                if not s or s.startswith('#'):
                    continue
                d = action_parse_line(s)
                try:
                    t_ms = int(d.get('ms') or 0)
                except Exception:
                    t_ms = 0
                if t_ms > max_ms:
                    max_ms = t_ms
        return max_ms
    except Exception:
        return 0

# Simple monitor: template matching using OpenCV
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
                # timeout: stop current playback and request restart
                for cb in self.stop_callbacks:
                    try:
                        cb()
                    except Exception:
                        pass
                if callable(self.restart_callback):
                    try:
                        self.restart_callback()
                    except Exception:
                        pass
                return
            # take screenshot and match
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

def save_settings():
    data = {}
    try:
        data['play_count'] = int(playCount.get())
    except Exception:
        data['play_count'] = 1
    try:
        data['infinite'] = bool(infiniteRepeatVar.get())
    except Exception:
        data['infinite'] = False
    try:
        val = actionFileVar.get().strip()
        if val:
            data['last_action'] = val
    except Exception:
        pass
    # persist game mode selection if available
    try:
        data['game_mode_relative'] = bool(gameModeVar.get())
    except Exception:
        pass
    # persist game mode gain/auto if available
    try:
        data['game_mode_gain'] = float(gameModeGainVar.get())
    except Exception:
        pass
    try:
        data['game_mode_auto'] = bool(gameModeAutoVar.get())
    except Exception:
        pass
    settings_mod.save_settings(data, SETTINGS_PATH)

def apply_settings_to_ui(settings: dict):
    try:
        settings_mod.apply_settings_to_ui(settings, playCount, infiniteRepeatVar, actionFileVar, list_action_files)
    except Exception:
        pass
    # apply game mode selection from settings
    try:
        if 'game_mode_relative' in settings and 'gameModeVar' in globals():
            gameModeVar.set(bool(settings.get('game_mode_relative', False)))
    except Exception:
        pass
    try:
        if 'gameModeGainVar' in globals() and 'game_mode_gain' in settings:
            gameModeGainVar.set(float(settings.get('game_mode_gain', 1.0) or 1.0))
    except Exception:
        pass
    try:
        if 'gameModeAutoVar' in globals() and 'game_mode_auto' in settings:
            gameModeAutoVar.set(bool(settings.get('game_mode_auto', True)))
    except Exception:
        pass

# Sync the Combobox to point at the current recording file
def select_current_action_in_dropdown():
    try:
        if 'actionFileVar' in globals():
            name = os.path.basename(action_file_name) if 'action_file_name' in globals() else ''
            files2 = list_action_files()
            # ensure list contains the current file name
            if name and name not in files2:
                files2.append(name)
            if 'actionFileSelect' in globals():
                actionFileSelect['values'] = files2
            if name:
                actionFileVar.set(name)
            elif files2:
                actionFileVar.set(files2[-1])
        # persist settings when selection changes
        try:
            save_settings()
        except Exception:
            pass
    except Exception:
        pass

######################################################################
# json template
######################################################################
def keyboard_action_template():
    return {
        "name": "keyboard",
        "event": "default",
        "vk": "default"
    }


def mouse_action_template():
    return {
        "name": "mouse",
        "event": "default",
        "target": "default",
        "action": "default",
        "location": {
            "x": "0",
            "y": "0"
        }
    }


######################################################################
# Receive Command
######################################################################
def command_adapter(action):
    # global variables
    global can_start_listening 
    global can_start_executing
    global execute_time_keyboard
    global execute_time_mouse
    global action_file_name
    
    # command list
    custom_thread_list = []
    print(can_start_listening)
    
    if can_start_listening and can_start_executing:
        if action == 'listen':
            # setup shared action file and start time
            init_new_action_file()
            # update UI selection to the new file
            select_current_action_in_dropdown()
            # reset listen stop event
            try:
                ev_stop_listen.clear()
            except Exception:
                pass
            # UI updates
            update_ui_for_state(AppState.RECORDING)
            # start recorder (keyboard + mouse)
            try:
                global current_recorder
                current_recorder = Recorder(action_file_name, ev_stop_listen)
                current_recorder.start()
            except Exception:
                pass
            can_start_listening = False
            can_start_executing = False

        elif action == 'execute':
            # set the selected action file for replay
            selected = actionFileVar.get().strip() if 'actionFileVar' in globals() else ''
            # resolve to actions/ if present
            sel_path = os.path.join('actions', selected) if selected else ''
            if not selected:
                startExecuteBtn['text'] = 'Select a .action file'
                return
            else:
                # use selected file (prefer actions/, fallback root)
                action_file_name = sel_path if os.path.exists(sel_path) else selected
            # init counters and flags
            try:
                if 'infiniteRepeatVar' in globals() and bool(infiniteRepeatVar.get()):
                    ev_infinite_replay.set()
                else:
                    ev_infinite_replay.clear()
            except Exception:
                pass
            execute_time_keyboard = playCount.get()
            execute_time_mouse = playCount.get()
            try:
                ev_stop_execute_keyboard.clear()
                ev_stop_execute_mouse.clear()
            except Exception:
                pass
            total_ms = 0
            try:
                total_ms = compute_action_total_ms(action_file_name)
            except Exception:
                total_ms = 0
            try:
                start_monitor(execute_time_keyboard, total_ms/1000.0)
            except Exception:
                pass
            # start monitor thread (template detection) if image exists
            try:
                global monitor_thread
                monitor_img = os.path.join('assets', 'monitor_target.png')
                if os.path.exists(monitor_img):
                    # stop previous monitor if any
                    try:
                        if monitor_thread and monitor_thread.is_alive():
                            monitor_thread.stop()
                    except Exception:
                        pass
                    def _stop_current():
                        try:
                            ev_stop_execute_keyboard.set()
                            ev_stop_execute_mouse.set()
                        except Exception:
                            pass
                    def _restart_main():
                        # restart primary action after stop
                        try:
                            root.after(0, lambda: command_adapter('execute'))
                        except Exception:
                            pass
                    monitor_thread = MonitorThread(
                        target_path=monitor_img,
                        timeout_s=300,      # 5 minutes
                        interval_s=3,       # check every 3s
                        stop_callbacks=[_stop_current],
                        restart_callback=_restart_main,
                        hit_callback=on_monitor_hit
                    )
                    monitor_thread.start()
            except Exception:
                pass
            # UI updates
            update_ui_for_state(AppState.REPLAYING)
            # start replayer (keyboard + mouse) and controller (ESC monitor)
            try:
                global current_replayer
                use_rel = False
                try:
                    use_rel = bool(gameModeVar.get())
                except Exception:
                    pass
                # pass gain/auto settings (editable via settings.json)
                try:
                    rel_gain = float(gameModeGainVar.get())
                except Exception:
                    try:
                        _s = load_settings(); rel_gain = float(_s.get('game_mode_gain', 1.0) or 1.0)
                    except Exception:
                        rel_gain = 1.0
                try:
                    rel_auto = bool(gameModeAutoVar.get())
                except Exception:
                    try:
                        _s = load_settings(); rel_auto = bool(_s.get('game_mode_auto', True))
                    except Exception:
                        rel_auto = True
                # progress callback updates monitor panel
                def _on_progress(done, total):
                    try:
                        update_replay_progress(done, total)
                    except Exception:
                        pass
                # loop_start_cb resets the loop timer to 0 for each new loop
                def _on_loop_start(idx, total):
                    try:
                        update_replay_loop_start(idx, total)
                    except Exception:
                        pass
                current_replayer = Replayer(action_file_name, ev_stop_execute_keyboard, ev_stop_execute_mouse, ev_infinite_replay, execute_time_keyboard, use_rel, rel_gain, rel_auto, _on_progress, _on_loop_start)
                current_replayer.start()
            except Exception:
                pass
            ExecuteController().start()
            can_start_listening = False
            can_start_executing = False


######################################################################
# Update UI
######################################################################        
## Countdown removed — direct start/stop via F10/F11


######################################################################
# Listen
###################################################################### 
class KeyboardActionListener(threading.Thread):
    
    def __init__(self, file_name='keyboard.action'):
        super().__init__()
        self.daemon = True
        self.file_name = file_name

    def run(self):
        # Deprecated: use core.recorder.Recorder
        return
                

class MouseActionListener(threading.Thread):

    def __init__(self, file_name='mouse.action'):
        super().__init__()
        self.daemon = True
        self.file_name = file_name

    def run(self):
        # Deprecated: use core.recorder.Recorder
        return


######################################################################
# Executing
######################################################################
class KeyboardActionExecute(threading.Thread):

    def __init__(self, file_name='keyboard.action'):
        super().__init__()
        self.daemon = True
        self.file_name = file_name

    def run(self):
        global execute_time_keyboard
        global ev_stop_execute_keyboard
        global pressed_vks
        while True:
            if ev_stop_execute_keyboard.is_set():
                return
            try:
                path = action_file_name if os.path.exists(action_file_name) else self.file_name
                with open(path, 'r', encoding='utf-8') as file:
                    keyboard_exec = KeyBoardController()
                    start_ts = time.monotonic()
                    line = file.readline()
                    while line:
                        s = line.strip()
                        if not s:
                            line = file.readline(); continue
                        d = action_parse_line(s)
                        if d.get('type') == 'VK':
                            try:
                                t_ms = int(d.get('ms') or 0)
                            except Exception:
                                t_ms = 0
                            target = start_ts + (t_ms/1000.0)
                            if not wait_until_or_stop(target, ev_stop_execute_keyboard):
                                break
                            try:
                                vk = int(d.get('vk') or 0)
                                if d.get('op') == 'DOWN':
                                    keyboard_exec.press(KeyCode.from_vk(vk))
                                    pressed_vks.add(vk)
                                elif d.get('op') == 'UP':
                                    keyboard_exec.release(KeyCode.from_vk(vk))
                                    pressed_vks.discard(vk)
                            except Exception:
                                pass
                        else:
                            # legacy json support
                            try:
                                obj = json.loads(line)
                                if obj.get('name') == 'keyboard':
                                    if obj['event'] == 'press':
                                        vk = obj['vk']
                                        keyboard_exec.press(KeyCode.from_vk(vk))
                                        pressed_vks.add(vk)
                                    elif obj['event'] == 'release':
                                        vk = obj['vk']
                                        keyboard_exec.release(KeyCode.from_vk(vk))
                                        pressed_vks.discard(vk)
                                    time.sleep(0.005)
                            except Exception:
                                pass
                        line = file.readline()
            finally:
                # ensure all pressed keys are released
                try:
                    keyboard_exec = KeyBoardController()
                    for vk in list(pressed_vks):
                        try:
                            keyboard_exec.release(KeyCode.from_vk(vk))
                        except Exception:
                            pass
                        pressed_vks.discard(vk)
                except Exception:
                    pass
            if 'ev_infinite_replay' in globals() and ev_infinite_replay.is_set():
                continue
            execute_time_keyboard = execute_time_keyboard - 1
            if execute_time_keyboard <= 0:
                try:
                    ev_stop_execute_keyboard.set()
                except Exception:
                    pass
                return

class MouseActionExecute(threading.Thread):

    def __init__(self, file_name='mouse.action'):
        super().__init__()
        self.daemon = True
        self.file_name = file_name

    def run(self):
        global execute_time_mouse
        global ev_stop_execute_mouse
        while True:
            if ev_stop_execute_mouse.is_set():
                return
            try:
                path = action_file_name if os.path.exists(action_file_name) else self.file_name
                with open(path, 'r', encoding='utf-8') as file:
                    mouse_exec = MouseController()
                    # playback-time screen size
                    cw, ch = get_screen_size()
                    rw, rh = cw, ch
                    start_ts = time.monotonic()
                    # pressed buttons tracking for cleanup
                    global pressed_mouse_buttons
                    line = file.readline()
                    while line:
                        s = line.strip()
                        if not s:
                            line = file.readline(); continue
                        d = action_parse_line(s)
                        if d.get('type') == 'META' and d.get('op') == 'SCREEN':
                            try:
                                rw = int(d.get('x') or cw); rh = int(d.get('y') or ch)
                            except Exception:
                                rw, rh = cw, ch
                        elif d.get('type') == 'MS':
                            # compute target time
                            try:
                                t_ms = int(d.get('ms') or 0)
                            except Exception:
                                t_ms = 0
                            target = start_ts + (t_ms/1000.0)
                            if not wait_until_or_stop(target, ev_stop_execute_mouse):
                                break
                            if d.get('op') == 'MOVE':
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
                                mouse_exec.position = (tx, ty)
                            elif d.get('op') == 'CLICK':
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
                                    mouse_exec.position = (tx, ty)
                                except Exception:
                                    pass
                                if act == 'DOWN':
                                    if btn == 'left':
                                        mouse_exec.press(Button.left); pressed_mouse_buttons.add('left')
                                    else:
                                        mouse_exec.press(Button.right); pressed_mouse_buttons.add('right')
                                else:
                                    if btn == 'left':
                                        mouse_exec.release(Button.left); pressed_mouse_buttons.discard('left')
                                    else:
                                        mouse_exec.release(Button.right); pressed_mouse_buttons.discard('right')
                            elif d.get('op') == 'SCROLL':
                                try:
                                    dx = int(d.get('dx') or 0); dy = int(d.get('dy') or 0)
                                except Exception:
                                    dx = dy = 0
                                mouse_exec.scroll(dx, dy)
                        else:
                            # legacy json support
                            try:
                                obj = json.loads(line)
                                if obj.get('name') == 'meta':
                                    rw = int(obj['screen']['w']); rh = int(obj['screen']['h'])
                                elif obj.get('name') == 'mouse':
                                    if obj['event'] == 'move':
                                        mouse_exec.position = (int(obj['location']['x']), int(obj['location']['y']))
                                        time.sleep(0.005)
                                    elif obj['event'] == 'click':
                                        if obj['action']:
                                            (mouse_exec.press(Button.left) if obj['target']=='left' else mouse_exec.press(Button.right))
                                        else:
                                            (mouse_exec.release(Button.left) if obj['target']=='left' else mouse_exec.release(Button.right))
                                        time.sleep(0.005)
                                    elif obj['event'] == 'scroll':
                                        mouse_exec.scroll(obj['location']['x'], obj['location']['y'])
                                        time.sleep(0.005)
                            except Exception:
                                pass
                        line = file.readline()
            finally:
                # ensure buttons are released
                try:
                    mouse_exec = MouseController()
                    if 'left' in pressed_mouse_buttons:
                        mouse_exec.release(Button.left)
                        pressed_mouse_buttons.discard('left')
                    if 'right' in pressed_mouse_buttons:
                        mouse_exec.release(Button.right)
                        pressed_mouse_buttons.discard('right')
                except Exception:
                    pass
            if 'ev_infinite_replay' in globals() and ev_infinite_replay.is_set():
                continue
            execute_time_mouse = execute_time_mouse - 1
            if execute_time_mouse <= 0:
                try:
                    ev_stop_execute_mouse.set()
                except Exception:
                    pass
                return
                
                
######################################################################
# Controller
######################################################################
class ListenController(threading.Thread):
    
    def __init__(self):
        super().__init__()
        self.daemon = True

    def run(self):
        global ev_stop_listen
        try:
            ev_stop_listen.clear()
        except Exception:
            pass
        
        def on_release(key):
            global can_start_listening 
            global can_start_executing
            global ev_stop_listen
            
            if key == keyboard.Key.esc:
                try:
                    ev_stop_listen.set()
                except Exception:
                    pass
                can_start_listening = True
                can_start_executing = True
                update_ui_for_state(AppState.IDLE)
                keyboardListener.stop()

        with keyboard.Listener(on_release=on_release) as keyboardListener:
            keyboardListener.join()

class ExecuteController(threading.Thread):
    
    def __init__(self):
        super().__init__()
        self.daemon = True

    def run(self):
        global ev_stop_execute_keyboard
        global ev_stop_execute_mouse
        global can_start_listening 
        global can_start_executing

        # Listener to allow ESC to stop replaying
        def on_release(key):
            global ev_stop_execute_keyboard
            global ev_stop_execute_mouse
            if key == keyboard.Key.esc:
                try:
                    ev_stop_execute_keyboard.set()
                    ev_stop_execute_mouse.set()
                except Exception:
                    pass

        keyboardListener = keyboard.Listener(on_release=on_release)
        keyboardListener.start()

        # Wait until all active workers have finished (or ESC pressed)
        while not (ev_stop_execute_keyboard.is_set() and ev_stop_execute_mouse.is_set()):
            time.sleep(0.05)

        # Safety: release any stuck inputs
        release_all_inputs()

        # Reset UI and states once everything is done
        can_start_listening = True
        can_start_executing = True
        update_ui_for_state(AppState.IDLE)
        keyboardListener.stop()


######################################################################
# Global Hotkeys (F10/F11)
######################################################################
class HotkeyController(threading.Thread):
    def __init__(self):
        super().__init__()
        self.daemon = True

    def run(self):
        def toggle_record():
            global can_start_listening, can_start_executing
            global root, ev_stop_listen
            # if idle, start recording; else stop via ESC
            if can_start_listening and can_start_executing:
                try:
                    # schedule on Tk main thread to avoid cross-thread UI ops
                    root.after(0, lambda: command_adapter('listen'))
                except Exception:
                    command_adapter('listen')
            else:
                # do NOT inject ESC; directly signal recorder to stop
                try:
                    ev_stop_listen.set()
                except Exception:
                    pass
                # reset UI state immediately
                try:
                    can_start_listening = True
                    can_start_executing = True
                    root.after(0, lambda: update_ui_for_state(AppState.IDLE))
                except Exception:
                    pass

        def toggle_replay():
            global can_start_listening, can_start_executing
            global ev_stop_execute_keyboard, ev_stop_execute_mouse
            global root
            # if idle, start replay; else request stop
            if can_start_listening and can_start_executing:
                try:
                    # schedule on Tk main thread to avoid cross-thread UI ops
                    root.after(0, lambda: command_adapter('execute'))
                except Exception:
                    command_adapter('execute')
            else:
                try:
                    ev_stop_execute_keyboard.set()
                    ev_stop_execute_mouse.set()
                    root.after(0, lambda: update_ui_for_state(AppState.IDLE))
                except Exception:
                    pass
        
        last_f10 = 0.0
        last_f11 = 0.0

        def on_press(key):
            nonlocal last_f10, last_f11
            now = time.time()
            try:
                if key == Key.f10:
                    if now - last_f10 > 0.3:
                        last_f10 = now
                        toggle_record()
                elif key == Key.f11:
                    if now - last_f11 > 0.3:
                        last_f11 = now
                        toggle_replay()
            except Exception:
                pass

        with keyboard.Listener(on_press=on_press) as hk:
            hk.join()

            
######################################################################
# GUI
######################################################################
if __name__ == '__main__':
    # Logging setup
    try:
        logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s %(message)s')
    except Exception:
        pass
    # UAC: attempt to elevate before creating any UI (Windows)
    try:
        relaunch_as_admin_if_needed()
    except Exception:
        pass
    # Ensure DPI awareness before creating Tk to avoid window size jumps
    set_process_dpi_aware()

    can_start_listening = True
    can_start_executing = True
    execute_time_keyboard = 0
    execute_time_mouse = 0
    # threading Events for coordination
    ev_stop_execute_keyboard = threading.Event(); ev_stop_execute_keyboard.set()
    ev_stop_execute_mouse = threading.Event(); ev_stop_execute_mouse.set()
    ev_stop_listen = threading.Event(); ev_stop_listen.set()
    ev_infinite_replay = threading.Event(); ev_infinite_replay.clear()
    pressed_vks = set()
    pressed_mouse_buttons = set()
    
    root = tkinter.Tk()

    # Business UI theme setup
    def setup_business_theme(win):
        style = ttk.Style()
        try:
            # Prefer native-looking theme on Windows
            style.theme_use('vista')
        except Exception:
            pass
        # pick a cute font if available
        preferred = [
            'Segoe UI', 'Microsoft YaHei UI', '微软雅黑', 'Arial'
        ]
        fams = set(tkfont.families())
        font_family = None
        for f in preferred:
            if f in fams:
                font_family = f
                break
        if not font_family:
            font_family = 'Segoe UI'

        bg = '#ffffff'        # white
        fg = '#1f2937'        # slate-800
        win.configure(bg=bg)

        # global default font tweaks
        try:
            default_font = tkfont.nametofont('TkDefaultFont')
            default_font.configure(family=font_family, size=12)
        except Exception:
            pass
        # Labels
        style.configure('Biz.TLabel', background=bg, foreground=fg, font=(font_family, 12))
        style.configure('BizTitle.TLabel', background=bg, foreground=fg, font=(font_family, 18, 'bold'))
        # Buttons/Entries/Combobox use native theme visuals; add padding only
        style.configure('Biz.TButton', font=(font_family, 13), padding=10)
        style.configure('Biz.TEntry', padding=6)
        style.configure('Biz.TCombobox', padding=6)

        return font_family, bg

    font_family, bg_color = setup_business_theme(root)
    ensure_assets_dir()
    ensure_actions_dir()

    # optional assets: icon and background
    try:
        icon_path = os.path.join('assets', 'icon.png')
        if os.path.exists(icon_path):
            root._icon_img = tkinter.PhotoImage(file=icon_path)
            root.iconphoto(True, root._icon_img)
    except Exception:
        pass
    # Business theme: avoid decorative background image for a clean look

    root.title('Quick Macro')
    root.geometry('720x470')
    root.resizable(0,0)

    # title
    titleLabel = ttk.Label(root, text='Quick Macro', style='BizTitle.TLabel')
    titleLabel.place(x=24, y=14, width=220, height=36)

    # Card style containers
    style = ttk.Style()
    style.configure('Card.TFrame', background='#f8fafc')
    style.configure('CardLabel.TLabel', background='#f8fafc', font=(font_family, 10), foreground='#374151')
    # Ensure checkbutton blends with card background (no visible patch)
    style.configure('Card.TCheckbutton', background='#f8fafc', font=(font_family, 10))
    style.map('Card.TCheckbutton', background=[('active', '#f8fafc'), ('!active', '#f8fafc')])
    style.configure('Biz.TButton', anchor='center', font=(font_family, 10), padding=(20, 0))
    # Explicit centered button style with symmetric padding for perfect centering
    style.configure('Center.TButton', anchor='center', font=(font_family, 10), padding=(20, 0))

    recordCard = ttk.Frame(root, style='Card.TFrame', borderwidth=1, relief='solid')
    recordCard.place(x=30, y=70, width=310, height=90)
    replayCard = ttk.Frame(root, style='Card.TFrame', borderwidth=1, relief='solid')
    replayCard.place(x=360, y=70, width=330, height=250)
    monitorCard = ttk.Frame(root, style='Card.TFrame', borderwidth=1, relief='solid')
    monitorCard.place(x=30, y=330, width=660, height=110)

    # start recording
    startListenerBtn = ttk.Button(recordCard, text="Start recording (F10)", command=lambda: command_adapter('listen'), style='Center.TButton')
    startListenerBtn.place(x=15, y=10, width=280, height=44)

    # times for replaying
    playCountLabel = ttk.Label(replayCard, text='Repeat Times', style='CardLabel.TLabel')
    playCountLabel.place(x=15, y=15, width=120, height=30)
    
    playCount = tkinter.IntVar()
    playCount.set(1)
    
    playCountEdit = ttk.Entry(replayCard, textvariable=playCount, style='Biz.TEntry')
    playCountEdit.place(x=140, y=15, width=80, height=30)

    playCountTipLabel = ttk.Label(replayCard, text='#', style='CardLabel.TLabel')
    playCountTipLabel.place(x=225, y=15, width=20, height=30)

    # infinite repeat checkbox
    global infiniteRepeatVar
    infiniteRepeatVar = tkinter.BooleanVar()
    infiniteRepeatVar.set(False)
    infiniteCheck = ttk.Checkbutton(replayCard, text='Inf.', variable=infiniteRepeatVar, style='Card.TCheckbutton')
    infiniteCheck.place(x=250, y=15, width=70, height=30)

    # start replaying button centered in card
    startExecuteBtn = ttk.Button(replayCard, text="Start replaying (F11)", command=lambda: command_adapter('execute'), style='Center.TButton')
    startExecuteBtn.place(x=15, y=60, width=280, height=40)

    # Game mode (relative mouse) toggle
    global gameModeVar
    gameModeVar = tkinter.BooleanVar()
    gameModeVar.set(False)
    gameModeCheck = ttk.Checkbutton(replayCard, text='Game mode', variable=gameModeVar, style='Card.TCheckbutton')
    gameModeCheck.place(x=15, y=105, width=130, height=26)
    # Game mode options: Gain and Auto detect
    global gameModeGainVar, gameModeAutoVar
    gameModeGainVar = tkinter.DoubleVar(); gameModeGainVar.set(1.0)
    gameModeAutoVar = tkinter.BooleanVar(); gameModeAutoVar.set(True)
    gainLabel = ttk.Label(replayCard, text='Gain', style='CardLabel.TLabel')
    gainLabel.place(x=155, y=105, width=35, height=26)
    gainEntry = ttk.Entry(replayCard, textvariable=gameModeGainVar, style='Biz.TEntry')
    gainEntry.place(x=195, y=105, width=60, height=26)
    autoCheck = ttk.Checkbutton(replayCard, text='Auto detect', variable=gameModeAutoVar, style='Card.TCheckbutton')
    autoCheck.place(x=15, y=135, width=140, height=24)

    # Monitor area (bottom)
    monitorTitle = ttk.Label(monitorCard, text='Monitor', style='CardLabel.TLabel')
    monitorTitle.place(x=12, y=8, width=80, height=24)
    monitorLoopLabel = ttk.Label(monitorCard, text='Loops: 0/0', style='CardLabel.TLabel')
    monitorLoopLabel.place(x=12, y=40, width=200, height=24)
    monitorTimeLabel = ttk.Label(monitorCard, text='Current loop time: 0.0s', style='CardLabel.TLabel')
    monitorTimeLabel.place(x=220, y=40, width=220, height=24)
    monitorTotalLabel = ttk.Label(monitorCard, text='Total loop time: 0.0s', style='CardLabel.TLabel')
    monitorTotalLabel.place(x=460, y=40, width=180, height=24)
    dungeonTimeLabel = ttk.Label(monitorCard, text='Dungeon time: 0.0s', style='CardLabel.TLabel')
    dungeonTimeLabel.place(x=12, y=70, width=250, height=24)
    # Monitor state & helpers
    monitor_total_loops = 0
    monitor_completed_loops = 0
    monitor_loop_start_ts = None
    monitor_timer_job = None
    monitor_total_time_s = 0.0
    dungeon_start_ts = None

    def update_monitor_labels():
        try:
            loops_text = f"Loops: {monitor_completed_loops}/{monitor_total_loops}" if monitor_total_loops else "Loops: 0/0"
            monitorLoopLabel['text'] = loops_text
            elapsed = 0.0
            if monitor_loop_start_ts is not None:
                elapsed = max(0.0, time.monotonic() - monitor_loop_start_ts)
            monitorTimeLabel['text'] = f"Current loop time: {elapsed:.1f}s"
            monitorTotalLabel['text'] = f"Total loop time: {monitor_total_time_s:.1f}s"
            dungeon_elapsed = 0.0
            if dungeon_start_ts is not None:
                dungeon_elapsed = max(0.0, time.monotonic() - dungeon_start_ts)
            dungeonTimeLabel['text'] = f"Dungeon time: {dungeon_elapsed:.1f}s"
        except Exception:
            pass

    def _tick_monitor():
        global monitor_timer_job
        update_monitor_labels()
        try:
            monitor_timer_job = root.after(200, _tick_monitor)
        except Exception:
            monitor_timer_job = None

    def start_monitor(total_loops: int, total_time_s: float = 0.0):
        global monitor_total_loops, monitor_completed_loops, monitor_loop_start_ts, monitor_timer_job, monitor_total_time_s, dungeon_start_ts
        monitor_total_loops = max(1, int(total_loops or 1))
        monitor_completed_loops = 0
        monitor_loop_start_ts = time.monotonic()
        dungeon_start_ts = None
        try:
            monitor_total_time_s = max(0.0, float(total_time_s or 0.0))
        except Exception:
            monitor_total_time_s = 0.0
        if monitor_timer_job:
            try:
                root.after_cancel(monitor_timer_job)
            except Exception:
                pass
            monitor_timer_job = None
        update_monitor_labels()
        _tick_monitor()

    def update_replay_progress(done: int, total: int):
        global monitor_completed_loops, monitor_total_loops, monitor_loop_start_ts
        try:
            monitor_completed_loops = int(done or 0)
        except Exception:
            monitor_completed_loops = done
        try:
            if total:
                monitor_total_loops = int(total)
        except Exception:
            pass
        update_monitor_labels()

    def update_replay_loop_start(loop_idx: int, total: int):
        global monitor_loop_start_ts, monitor_total_loops
        try:
            monitor_total_loops = int(total or monitor_total_loops or 0)
        except Exception:
            pass
        monitor_loop_start_ts = time.monotonic()
        update_monitor_labels()

    # Monitor thread handle
    monitor_thread = None

    def on_monitor_hit():
        # mark dungeon start at first detection of target image
        global dungeon_start_ts
        dungeon_start_ts = time.monotonic()
        update_monitor_labels()

    def reset_monitor():
        global monitor_total_loops, monitor_completed_loops, monitor_loop_start_ts, monitor_timer_job, dungeon_start_ts
        monitor_total_loops = 0
        monitor_completed_loops = 0
        monitor_loop_start_ts = None
        dungeon_start_ts = None
        if monitor_timer_job:
            try:
                root.after_cancel(monitor_timer_job)
            except Exception:
                pass
            monitor_timer_job = None
        update_monitor_labels()

    # UI state helper
    class AppState:
        IDLE = 'idle'
        RECORDING = 'recording'
        REPLAYING = 'replaying'

    def update_ui_for_state(state: str):
        if state == AppState.IDLE:
            startListenerBtn.state(['!disabled'])
            startExecuteBtn.state(['!disabled'])
            startListenerBtn['text'] = 'Start recording (F10)'
            startExecuteBtn['text'] = 'Start replaying (F11)'
            reset_monitor()
        elif state == AppState.RECORDING:
            startListenerBtn.state(['disabled'])
            startExecuteBtn.state(['disabled'])
            startListenerBtn['text'] = 'Recording, "F10" to stop.'
        elif state == AppState.REPLAYING:
            startListenerBtn.state(['disabled'])
            startExecuteBtn.state(['disabled'])
            startExecuteBtn['text'] = 'Replaying, "ESC/F11" to stop.'
        # ensure monitor thread is stopped when exiting replay
        if state == AppState.IDLE:
            try:
                if 'monitor_thread' in globals() and monitor_thread:
                    monitor_thread.stop()
            except Exception:
                pass
        try:
            root.update_idletasks()
        except Exception:
            pass

    actionFileVar = tkinter.StringVar()
    files = list_action_files()
    actionFileVar.set(files[-1] if files else '')

    # Action file controls inside the Replay card
    actionFileLabel = ttk.Label(replayCard, text='Action file', style='CardLabel.TLabel')
    actionFileLabel.place(x=15, y=170, width=100, height=26)
    actionFileSelect = ttk.Combobox(replayCard, textvariable=actionFileVar, values=files if files else [], state='readonly', style='Biz.TCombobox')
    actionFileSelect.place(x=120, y=170, width=190, height=28)

    # Refresh button removed; list auto-updates after recording
    
    # 加载设置并应用
    try:
        _settings = load_settings()
        apply_settings_to_ui(_settings)
    except Exception:
        pass
    # 变更即保存
    try:
        playCount.trace_add('write', lambda *_: save_settings())
    except Exception:
        try:
            playCount.trace('w', lambda *_: save_settings())
        except Exception:
            pass
    try:
        infiniteRepeatVar.trace_add('write', lambda *_: save_settings())
    except Exception:
        try:
            infiniteRepeatVar.trace('w', lambda *_: save_settings())
        except Exception:
            pass
    try:
        actionFileSelect.bind('<<ComboboxSelected>>', lambda *_: save_settings())
    except Exception:
        pass
    # Persist game mode changes
    try:
        gameModeVar.trace_add('write', lambda *_: save_settings())
    except Exception:
        try:
            gameModeVar.trace('w', lambda *_: save_settings())
        except Exception:
            pass
    # Persist Auto toggle changes
    try:
        gameModeAutoVar.trace_add('write', lambda *_: save_settings())
    except Exception:
        try:
            gameModeAutoVar.trace('w', lambda *_: save_settings())
        except Exception:
            pass
    # Gain: validate and persist on change
    def _commit_gain(*_):
        try:
            val = float(gameModeGainVar.get())
            if not (0.01 <= val <= 10.0):
                gameModeGainVar.set(1.0)
        except Exception:
            gameModeGainVar.set(1.0)
        try:
            save_settings()
        except Exception:
            pass
    try:
        gameModeGainVar.trace_add('write', _commit_gain)
    except Exception:
        try:
            gameModeGainVar.trace('w', _commit_gain)
        except Exception:
            pass
    
    # Editor for .action files (Excel-like simple table: line number + text)
    def resolve_selected_action_path():
        try:
            selected = actionFileVar.get().strip()
            if not selected:
                return None
            p_actions = os.path.join('actions', selected)
            return p_actions if os.path.exists(p_actions) else (selected if os.path.exists(selected) else None)
        except Exception:
            return None

    def open_actions_folder():
        try:
            actions_dir = os.path.abspath('actions')
            if os.path.isdir(actions_dir):
                if os.name == 'nt':
                    os.startfile(actions_dir)
                else:
                    import subprocess
                    subprocess.Popen(['xdg-open', actions_dir])
        except Exception:
            pass

    def open_action_editor():
        # Prevent editing while recording/replaying
        if not (can_start_listening and can_start_executing):
            messagebox.showwarning('Busy', 'Please stop recording/replaying before editing.')
            return
        path = resolve_selected_action_path()
        if not path:
            messagebox.showinfo('No file', 'Please select an .action file first.')
            return

        editor = tkinter.Toplevel(root)
        editor.title(f'Edit Action - {os.path.basename(path)}')
        editor.geometry('800x500')
        editor.grab_set()

        editor.grid_rowconfigure(1, weight=1)
        editor.grid_columnconfigure(0, weight=1)

        frame_top = ttk.Frame(editor)
        frame_top.grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 0))

        frame = ttk.Frame(editor)
        frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)

        # Styled, taller rows and centered content
        estyle = ttk.Style(editor)
        estyle.configure('Action.Treeview', rowheight=40, padding=0)
        estyle.configure('Action.Heading', padding=0)

        columns = ('line','type','op','vk','btn','act','x','y','nx','ny','dx','dy','ms','raw')
        headings = {
            'line':'#','type':'Type','op':'Op','vk':'VK','btn':'Btn','act':'Act',
            'x':'X','y':'Y','nx':'NX','ny':'NY','dx':'DX','dy':'DY','ms':'MS','raw':'Raw'
        }
        tree = ttk.Treeview(frame, columns=columns, show='headings', style='Action.Treeview', selectmode='extended')
        for c in columns:
            tree.heading(c, text=headings[c], anchor='center')
        center_cols = ['line','type','op','vk','btn','act','x','y','nx','ny','dx','dy','ms']
        for c in center_cols:
            tree.column(c, width=60, anchor='center', stretch=True)
        tree.column('raw', width=200, anchor='w', stretch=True)
        tree.column('line', width=40)

        vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        # use shared parser/serializer
        def parse_action_line(s):
            return action_parse_line(s)
        def compose_action_line(d):
            return action_compose_line(d)

        def allowed_columns_for(row_type, row_op):
            if row_type == 'VK':
                return {'op','vk','ms'}
            if row_type == 'MS':
                if (row_op or '').upper() == 'MOVE':
                    return {'op','x','y','nx','ny','ms'}
                if (row_op or '').upper() == 'CLICK':
                    return {'op','btn','act','x','y','nx','ny','ms'}
                if (row_op or '').upper() == 'SCROLL':
                    return {'op','dx','dy','ms'}
                return {'op','ms'}
            # META or comments
            return set()

        # Load lines
        try:
            with open(path, 'r', encoding='utf-8') as f:
                lines = [ln.rstrip('\n') for ln in f.readlines()]
        except Exception as e:
            messagebox.showerror('Error', f'Failed to open file:\n{e}')
            editor.destroy(); return

        # Extract META lines for header and exclude them from the grid
        meta_lines = []
        meta_screen = ''
        meta_start = ''
        data_lines = []
        for text in lines:
            if text.startswith('META '):
                meta_lines.append(text)
                parts = text.split()
                if len(parts)>=3 and parts[1] == 'SCREEN' and len(parts)>=4:
                    meta_screen = f"{parts[2]}x{parts[3]}"
                if len(parts)>=3 and parts[1] == 'START':
                    meta_start = parts[2]
            elif text.startswith('#') or text.strip() == '':
                # skip comments and blank lines from grid view
                continue
            else:
                data_lines.append(text)

        meta_disp = f"Screen: {meta_screen or 'N/A'}    Start: {meta_start or 'N/A'}"
        meta_label = ttk.Label(frame_top, text=meta_disp, style='Biz.TLabel')
        meta_label.pack(anchor='w')

        for idx, text in enumerate(data_lines, start=1):
            d = parse_action_line(text)
            vals = [idx, d['type'], d['op'], d['vk'], d['btn'], d['act'], d['x'], d['y'], d['nx'], d['ny'], d['dx'], d['dy'], d['ms'], d['raw']]
            tree.insert('', 'end', values=tuple(vals))

        edit_entry = None
        edit_item = None
        edit_col = None
        edit_commit = None
        last_spawn_ts = 0.0
        selecting_in_editor = False
        last_spawn_ts = 0.0

        def cancel_inline_editor(*_):
            nonlocal edit_entry, edit_item, edit_col
            try:
                if edit_entry:
                    edit_entry.destroy()
            except Exception:
                pass
            edit_entry = None
            edit_item = None
            edit_col = None

        def begin_edit(event):
            nonlocal edit_entry, edit_item, edit_col
            region = tree.identify('region', event.x, event.y)
            if region != 'cell':
                return
            colid = tree.identify_column(event.x)
            # disallow editing of line number
            if colid == '#1':
                return
            row = tree.identify_row(event.y)
            if not row:
                return
            bbox = tree.bbox(row, colid)
            if not bbox:
                return
            x, y, w, h = bbox
            spawn_editor(row, columns[int(colid[1:]) - 1], x, y, w, h)

        def spawn_editor(row, colname, x, y, w, h):
            nonlocal edit_entry, edit_item, edit_col, edit_commit, last_spawn_ts, selecting_in_editor
            row_type = tree.set(row, 'type')
            row_op = tree.set(row, 'op')
            # prevent editing columns that are not applicable
            if colname not in allowed_columns_for(row_type, row_op) and colname not in ('type','raw'):
                return
            val = tree.set(row, colname)
            edit_item = row
            edit_col = colname

            # Choose editor widget: combobox for enumerations, entry for others
            def finish_edit_value(new_val):
                nonlocal edit_entry, edit_item, edit_col, edit_commit
                if edit_entry and edit_item and edit_col:
                    # normalize on type/op changes
                    if edit_col == 'type':
                        # only VK or MS
                        new_val = 'VK' if new_val == 'VK' else 'MS'
                        tree.set(edit_item, 'type', new_val)
                        # set default op and clear irrelevant fields
                        if new_val == 'VK':
                            tree.set(edit_item, 'op', 'DOWN' if tree.set(edit_item,'op') not in ('DOWN','UP') else tree.set(edit_item,'op'))
                            for c in ['btn','act','x','y','nx','ny','dx','dy']: tree.set(edit_item, c, '')
                        else:
                            tree.set(edit_item, 'op', 'MOVE' if tree.set(edit_item,'op') not in ('MOVE','CLICK','SCROLL') else tree.set(edit_item,'op'))
                            tree.set(edit_item, 'vk', '')
                    elif edit_col == 'op':
                        t = tree.set(edit_item,'type')
                        if t == 'VK':
                            new_val = 'DOWN' if new_val not in ('DOWN','UP') else new_val
                            tree.set(edit_item,'op', new_val)
                        else:
                            new_val = new_val if new_val in ('MOVE','CLICK','SCROLL') else 'MOVE'
                            tree.set(edit_item,'op', new_val)
                            # clear fields not applicable for chosen op
                            if new_val == 'MOVE':
                                for c in ['btn','act','dx','dy']: tree.set(edit_item,c,'')
                            elif new_val == 'CLICK':
                                for c in ['dx','dy']: tree.set(edit_item,c,'')
                                if tree.set(edit_item,'btn') not in ('left','right'): tree.set(edit_item,'btn','left')
                                if tree.set(edit_item,'act') not in ('DOWN','UP'): tree.set(edit_item,'act','DOWN')
                            elif new_val == 'SCROLL':
                                for c in ['btn','act','x','y','nx','ny']: tree.set(edit_item,c,'')
                    else:
                        # numeric validation
                        int_fields = {'vk','x','y','dx','dy','ms'}
                        float_fields = {'nx','ny'}
                        if edit_col in int_fields and new_val != '':
                            try:
                                int(new_val)
                            except Exception:
                                messagebox.showerror('Invalid', f'{edit_col} must be an integer')
                                cancel_inline_editor()
                                return
                        if edit_col in float_fields and new_val != '':
                            try:
                                float(new_val)
                            except Exception:
                                messagebox.showerror('Invalid', f'{edit_col} must be a float')
                                cancel_inline_editor()
                                return
                        tree.set(edit_item, edit_col, new_val)
                    edit_entry.destroy()
                edit_entry = None
                edit_item = None
                edit_col = None
                edit_commit = None

            # Editors
            # prepare a callable to commit from outside
            edit_commit = lambda: finish_edit_value(edit_entry.get() if hasattr(edit_entry, 'get') else '')
            last_spawn_ts = time.time()

            if colname == 'type':
                edit_entry = ttk.Combobox(tree, values=['VK','MS'], state='readonly')
                edit_entry.set(val if val in ('VK','MS') else 'VK')
            elif colname == 'op':
                if row_type == 'VK':
                    edit_entry = ttk.Combobox(tree, values=['DOWN','UP'], state='readonly')
                else:
                    edit_entry = ttk.Combobox(tree, values=['MOVE','CLICK','SCROLL'], state='readonly')
                edit_entry.set(val if val else ( 'DOWN' if row_type=='VK' else 'MOVE'))
            elif colname == 'btn':
                edit_entry = ttk.Combobox(tree, values=['left','right'], state='readonly')
                edit_entry.set(val if val in ('left','right') else 'left')
            elif colname == 'act':
                edit_entry = ttk.Combobox(tree, values=['DOWN','UP'], state='readonly')
                edit_entry.set(val if val in ('DOWN','UP') else 'DOWN')
            else:
                # plain entry for numeric/text
                edit_entry = ttk.Entry(tree)
                edit_entry.insert(0, val)

            edit_entry.place(x=x, y=y, width=w, height=h)
            edit_entry.focus()
            # Track mouse selection inside entry to avoid committing while drag-selecting
            def _entry_press(_e=None):
                nonlocal selecting_in_editor
                selecting_in_editor = True
            def _entry_release(_e=None):
                nonlocal selecting_in_editor
                selecting_in_editor = False
            try:
                edit_entry.bind('<Button-1>', _entry_press, add='+')
                edit_entry.bind('<B1-Motion>', _entry_press, add='+')
                edit_entry.bind('<ButtonRelease-1>', _entry_release, add='+')
            except Exception:
                pass

            def finish_edit(*_):
                new_val = edit_entry.get() if hasattr(edit_entry, 'get') else ''
                finish_edit_value(new_val)

            edit_entry.bind('<Return>', finish_edit)
            edit_entry.bind('<Escape>', lambda *_: (edit_entry.destroy(), None))
            edit_entry.bind('<FocusOut>', finish_edit)
            # When selecting from dropdown, commit on selection
            try:
                edit_entry.bind('<<ComboboxSelected>>', lambda *_: finish_edit())
            except Exception:
                pass

        tree.bind('<Double-1>', begin_edit)

        # Commit inline editor when clicking elsewhere in the grid
        def commit_inline_if_any(event=None):
            nonlocal edit_commit, last_spawn_ts, selecting_in_editor, edit_entry
            try:
                if not edit_commit:
                    return
                # Skip immediate commit right after spawn (for double-click) or when selecting inside entry
                if (time.time() - last_spawn_ts) <= 0.2 or selecting_in_editor:
                    return
                # If click is inside the entry bounds, ignore
                if event is not None and edit_entry is not None:
                    try:
                        ex = edit_entry.winfo_rootx(); ey = edit_entry.winfo_rooty()
                        ew = edit_entry.winfo_width(); eh = edit_entry.winfo_height()
                        if ex <= event.x_root <= ex + ew and ey <= event.y_root <= ey + eh:
                            return
                    except Exception:
                        pass
                edit_commit()
            except Exception:
                pass

        tree.bind('<Button-1>', commit_inline_if_any, add='+')
        editor.bind('<Button-1>', commit_inline_if_any, add='+')

        # Drag-to-reorder support (multi-row with visual insert indicator)
        dragging_selection = []
        drag_insert_index = None
        insert_line = tkinter.Frame(frame, height=2, background='#ff4d4f')
        # auto-scroll state while dragging near edges
        last_drag_y = None
        drag_scroll_job = None
        # highlight tag for dragging block
        try:
            tree.tag_configure('drag_sel', background='#e6f7ff')
        except Exception:
            pass

        def clear_drag_highlight():
            try:
                for iid in tree.get_children(''):
                    tags = list(tree.item(iid, 'tags') or [])
                    if 'drag_sel' in tags:
                        tags.remove('drag_sel')
                        tree.item(iid, tags=tuple(tags))
            except Exception:
                pass

        def apply_drag_highlight(items):
            try:
                clear_drag_highlight()
                for iid in items:
                    tags = list(tree.item(iid, 'tags') or [])
                    if 'drag_sel' not in tags:
                        tags.append('drag_sel')
                        tree.item(iid, tags=tuple(tags))
            except Exception:
                pass

        pressed_row = None

        def on_tree_press(event):
            nonlocal dragging_selection, drag_insert_index, pressed_row
            commit_inline_if_any(event)
            row = tree.identify_row(event.y)
            pressed_row = row
            if not row:
                dragging_selection = []
                clear_drag_highlight()
                return
            # honor Ctrl/Shift multi-select; only force single select when no modifiers
            ctrl = (event.state & 0x0004) != 0
            shift = (event.state & 0x0001) != 0
            current_sel = list(tree.selection())
            if (row not in current_sel) and not (ctrl or shift):
                tree.selection_set((row,))
            else:
                # clicked inside current multiselection without modifiers: keep selection intact
                # and prevent default behavior from collapsing to single selection
                if (row in current_sel) and not (ctrl or shift):
                    editor.after(0, lambda: tree.selection_set(tuple(current_sel)))
                    apply_drag_highlight(current_sel)
                    drag_insert_index = None
                    return "break"
            drag_insert_index = None
            # do not compute dragging_selection yet; wait until motion so default selection can settle
            clear_drag_highlight()

        def compute_insert_at(event):
            children = tree.get_children('')
            if not children:
                return 0, 0
            target = tree.identify_row(event.y)
            if target:
                idx = tree.index(target)
                bbox = tree.bbox(target)
                if bbox and len(bbox) == 4:
                    x, y, w, h = bbox
                    above = event.y < (y + h/2)
                    insert_at = idx if above else idx + 1
                    line_y = y if above else y + h
                else:
                    # target may be scrolled out or bbox unavailable; fall back to cursor y
                    h_widget = max(0, tree.winfo_height())
                    insert_at = idx
                    line_y = max(0, min(event.y, h_widget))
                return insert_at, line_y
            # outside rows: use first/last bbox if available, else fall back to widget bounds
            first_bbox = tree.bbox(children[0])
            last_bbox = tree.bbox(children[-1])
            fy = first_bbox[1] if first_bbox and len(first_bbox) == 4 else 0
            ly = last_bbox[1] if last_bbox and len(last_bbox) == 4 else max(0, tree.winfo_height() - 1)
            lh = last_bbox[3] if last_bbox and len(last_bbox) == 4 else 0
            if event.y < fy:
                return 0, fy
            else:
                return len(children), ly + lh

        def on_tree_motion(event):
            nonlocal drag_insert_index, dragging_selection, last_drag_y, drag_scroll_job
            if not dragging_selection:
                # initialize dragging block from current selection (or pressed row)
                sel_now = list(tree.selection())
                if not sel_now and pressed_row:
                    sel_now = [pressed_row]
                children = tree.get_children('')
                idx_map = {iid: i for i, iid in enumerate(children)}
                dragging_selection = sorted(sel_now, key=lambda i: idx_map.get(i, 0))
                apply_drag_highlight(dragging_selection)
                # reinforce multiselect throughout drag
                try:
                    tree.selection_set(tuple(dragging_selection))
                except Exception:
                    pass
            # record last cursor y and (re)schedule autoscroll
            last_drag_y = event.y
            def autoscroll_tick():
                nonlocal drag_scroll_job, last_drag_y, drag_insert_index
                drag_scroll_job = None
                if not dragging_selection:
                    return
                try:
                    zone = 24
                    h = tree.winfo_height()
                    if last_drag_y is None:
                        return
                    if last_drag_y < zone:
                        tree.yview_scroll(-1, 'units')
                    elif last_drag_y > (h - zone):
                        tree.yview_scroll(1, 'units')
                    # after scrolling, update insert indicator at current cursor y
                    fake_event = type('E', (), {'y': last_drag_y})()
                    insert_at, line_y = compute_insert_at(fake_event)
                    drag_insert_index = insert_at
                    try:
                        insert_line.place(in_=tree, x=0, y=line_y, width=tree.winfo_width(), height=2)
                    except Exception:
                        pass
                finally:
                    # keep ticking while dragging
                    if dragging_selection:
                        drag_scroll_job = editor.after(50, autoscroll_tick)
            # schedule ticker if not already scheduled
            if drag_scroll_job is None:
                drag_scroll_job = editor.after(50, autoscroll_tick)
            insert_at, line_y = compute_insert_at(event)
            drag_insert_index = insert_at
            try:
                insert_line.place(in_=tree, x=0, y=line_y, width=tree.winfo_width(), height=2)
            except Exception:
                pass

        def on_tree_release(event):
            nonlocal dragging_selection, drag_insert_index, drag_scroll_job, last_drag_y
            try:
                insert_line.place_forget()
            except Exception:
                pass
            # cancel autoscroll ticker
            try:
                if drag_scroll_job is not None:
                    editor.after_cancel(drag_scroll_job)
            except Exception:
                pass
            drag_scroll_job = None
            last_drag_y = None
            if not dragging_selection or drag_insert_index is None:
                dragging_selection = []
                drag_insert_index = None
                clear_drag_highlight()
                return
            # Build new order by removing selection and inserting it at target index
            children = list(tree.get_children(''))
            idx_map = {iid: i for i, iid in enumerate(children)}
            sel_sorted = sorted(dragging_selection, key=lambda i: idx_map.get(i, 0))
            base = [iid for iid in children if iid not in sel_sorted]
            # insertion index in base list
            insertion = drag_insert_index
            # subtract how many selected were before the drop index
            before_count = sum(1 for iid in sel_sorted if idx_map[iid] < drag_insert_index)
            insertion -= before_count
            insertion = max(0, min(insertion, len(base)))
            new_order = base[:insertion] + sel_sorted + base[insertion:]
            # Apply moves according to new order
            for pos, iid in enumerate(new_order):
                try:
                    tree.move(iid, '', pos)
                except Exception:
                    pass
            renumber_lines()
            # keep the block selected and focused after move
            try:
                tree.selection_set(tuple(sel_sorted))
                if sel_sorted:
                    tree.focus(sel_sorted[0])
            except Exception:
                pass
            dragging_selection = []
            drag_insert_index = None
            clear_drag_highlight()

        tree.bind('<ButtonPress-1>', on_tree_press, add='+')
        tree.bind('<B1-Motion>', on_tree_motion, add='+')
        tree.bind('<ButtonRelease-1>', on_tree_release, add='+')

        def start_cell_edit(iid, colname):
            # ensure row visible, then compute bbox for column and spawn editor
            tree.see(iid)
            editor.update_idletasks()
            col_index = columns.index(colname) + 1
            bbox = tree.bbox(iid, f"#{col_index}")
            if not bbox:
                return
            x, y, w, h = bbox
            spawn_editor(iid, colname, x, y, w, h)
        # Remove cancel-on-click binding; commits are handled above

        btn_frame = ttk.Frame(editor)
        btn_frame.grid(row=2, column=0, sticky='ew', padx=10, pady=(0,10))

        def renumber_lines():
            for idx, iid in enumerate(tree.get_children(''), start=1):
                tree.set(iid, 'line', str(idx))

        def save_and_close():
            # collect rows in current order
            try:
                items = tree.get_children('')
                new_lines = []
                for i in items:
                    vals = {c: tree.set(i, c) for c in columns}
                    # compose from structured fields if possible
                    new_line = compose_action_line(vals)
                    new_lines.append(new_line)
                with open(path, 'w', encoding='utf-8') as f:
                    # write original meta lines first
                    if meta_lines:
                        f.write('\n'.join(meta_lines) + '\n')
                    f.write('\n'.join(new_lines) + '\n')
            except Exception as e:
                messagebox.showerror('Error', f'Failed to save file:\n{e}')
                return
            editor.destroy()

        def add_row():
            commit_inline_if_any()
            # default to VK DOWN 0; insert after current selection if any
            vals = {'type':'VK','op':'DOWN','vk':'0','btn':'','act':'','x':'','y':'','nx':'','ny':'','dx':'','dy':'','ms':'0','raw':''}
            sel = tree.selection()
            if sel:
                insert_index = tree.index(sel[-1]) + 1
            else:
                insert_index = 'end'
            iid = tree.insert('', insert_index, values=( '', vals['type'], vals['op'], vals['vk'], vals['btn'], vals['act'], vals['x'], vals['y'], vals['nx'], vals['ny'], vals['dx'], vals['dy'], vals['ms'], vals['raw']))
            renumber_lines()
            # start edit Type cell for convenience using explicit spawner
            start_cell_edit(iid, 'type')

        def delete_rows():
            commit_inline_if_any()
            sel = tree.selection()
            if not sel:
                return
            for iid in sel:
                tree.delete(iid)
            renumber_lines()

        def move_up():
            commit_inline_if_any()
            sel = list(tree.selection())
            if not sel:
                return
            for iid in sel:
                idx = tree.index(iid)
                if idx > 0:
                    tree.move(iid, '', idx-1)
            renumber_lines()

        def move_down():
            commit_inline_if_any()
            sel = list(tree.selection())
            if not sel:
                return
            for iid in reversed(sel):
                idx = tree.index(iid)
                tree.move(iid, '', idx+1)
            renumber_lines()

        save_btn = ttk.Button(btn_frame, text='Save', command=save_and_close)
        save_btn.pack(side='right', padx=6)
        close_btn = ttk.Button(btn_frame, text='Close', command=editor.destroy)
        close_btn.pack(side='right')
        add_btn = ttk.Button(btn_frame, text='Add', command=add_row)
        add_btn.pack(side='left')
        del_btn = ttk.Button(btn_frame, text='Delete', command=delete_rows)
        del_btn.pack(side='left', padx=6)
        # Removed Up/Down buttons in favor of drag-to-reorder

    editBtn = ttk.Button(replayCard, text='Edit', command=open_action_editor, style='Biz.TButton')
    editBtn.place(x=120, y=205, width=80, height=28)
    openBtn = ttk.Button(replayCard, text='Folder', command=open_actions_folder, style='Biz.TButton')
    openBtn.place(x=15, y=205, width=100, height=28)
    
    # Start hotkeys listener (F10/F11)
    HotkeyController().start()

    # Removed Tk window key binds to avoid double-trigger; global hotkeys handle F10/F11

    # Ensure closing window terminates app
    def on_close():
        global ev_stop_execute_keyboard, ev_stop_execute_mouse, ev_stop_listen
        try:
            ev_stop_execute_keyboard.set()
            ev_stop_execute_mouse.set()
            ev_stop_listen.set()
        except Exception:
            pass
        try:
            save_settings()
        except Exception:
            pass
        release_all_inputs()
        try:
            root.destroy()
        except Exception:
            pass

    root.protocol('WM_DELETE_WINDOW', on_close)

    # run
    root.mainloop()
    

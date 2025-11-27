import json
import sys
import logging
import ctypes
import os
import glob
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import threading
import time
import tkinter
from tkinter import ttk
import tkinter.font as tkfont
import tkinter.messagebox as messagebox
from tkinter import scrolledtext
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
from controllers.monitor import MonitorThread
from controllers.recording import RecordingController
from controllers.playback import PlaybackController
from controllers.hotkeys import HotkeyController
from controllers.listen import ListenController
from controllers.execute import ExecuteController

@dataclass
class AppState:
    can_start_listening: bool = True
    can_start_executing: bool = True
    action_file_name: str = ''
    record_start_time: float = 0.0
    execute_time_keyboard: int = 0
    execute_time_mouse: int = 0
    ev_stop_execute_keyboard: threading.Event = field(default_factory=lambda: threading.Event())
    ev_stop_execute_mouse: threading.Event = field(default_factory=lambda: threading.Event())
    ev_stop_listen: threading.Event = field(default_factory=lambda: threading.Event())
    ev_infinite_replay: threading.Event = field(default_factory=lambda: threading.Event())
    pressed_vks: set = field(default_factory=set)
    pressed_mouse_buttons: set = field(default_factory=set)
    monitor_thread: object = None
    monitor_timer_job: object = None
    monitor_total_loops: int = 0
    monitor_completed_loops: int = 0
    monitor_loop_start_ts: float = None
    monitor_total_time_s: float = 0.0
    dungeon_start_ts: float = None
    pending_main_action: str = None
    pending_main_playcount: int = None
    restarting_flag: bool = False
    restart_back_job: object = None
    restart_running: bool = False
    current_recorder: Recorder = None
    current_replayer: Replayer = None
    current_run_idx: int = 0
    current_run_action: str = ''
    current_run_interrupted: bool = False
    skip_run_increment: bool = False

state = AppState()
ui_refs = None

# UI references container
@dataclass
class UIRefs:
    root: object
    actionFileVar: object
    actionFileSelect: object
    startExecuteBtn: object
    startListenerBtn: object
    playCount: object
    infiniteRepeatVar: object
    gameModeVar: object
    gameModeGainVar: object
    gameModeAutoVar: object
    monitorTimeoutMs: object
    log_event: object
    update_ui_for_state: object
    begin_run: object
    mark_interrupted: object
    mark_finished: object
    recording_controller: object
    playback_controller: object
    listen_controller: object
    execute_controller: object
    listen_controller: object
    execute_controller: object

# Simple UI state enum for readability
class UiState:
    IDLE = 'idle'
    RECORDING = 'recording'
    REPLAYING = 'replaying'

######################################################################
# Helpers
######################################################################
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
    """Windows: try to elevate; on success exit current instance so only elevated UI remains."""
    try:
        if os.name != 'nt':
            return
        if _is_running_as_admin() or '--elevated' in sys.argv:
            logging.info('Admin check: running as administrator or elevated flag present.')
            return
        script = os.path.abspath(__file__)
        params = ' '.join([f'"{script}"', '--elevated'] + [f'"{a}"' for a in sys.argv[1:]])
        logging.info('Admin check: attempting UAC elevation...')
        rc = ctypes.windll.shell32.ShellExecuteW(None, 'runas', sys.executable, params, None, 1)
        if rc > 32:
            os._exit(0)
        else:
            logging.warning(f'UAC elevation failed, rc={rc}.')
            _win_message_box('Need admin', 'Could not elevate automatically. Please run as administrator for better compatibility.', 0x30)
    except Exception as e:
        try:
            logging.exception('UAC elevation exception: %s', e)
        except Exception:
            pass

def release_all_inputs():
    # Release any keys/buttons that might have been left pressed
    try:
        kb = KeyBoardController()
        for vk in list(state.pressed_vks):
            try:
                kb.release(KeyCode.from_vk(vk))
            except Exception:
                pass
            state.pressed_vks.discard(vk)
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
    ts = datetime.now().strftime('%Y%m%d-%H%M%S')
    actions_dir = ensure_actions_dir()
    state.action_file_name = str(actions_dir / f"{ts}.action")
    sw, sh = get_screen_size()
    state.record_start_time = time.time()
    try:
        with open(state.action_file_name, 'w', encoding='utf-8') as f:
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
######################################################################
# Controllers
######################################################################
def save_settings():
    if ui_refs is None:
        return
    data = {}
    try:
        data['play_count'] = int(ui_refs.playCount.get())
    except Exception:
        data['play_count'] = 1
    try:
        data['infinite'] = bool(ui_refs.infiniteRepeatVar.get())
    except Exception:
        data['infinite'] = False
    try:
        val = ui_refs.actionFileVar.get().strip()
        if val:
            data['last_action'] = val
    except Exception:
        pass
    # persist game mode selection if available
    try:
        data['game_mode_relative'] = bool(ui_refs.gameModeVar.get())
    except Exception:
        pass
    # persist game mode gain/auto if available
    try:
        data['game_mode_gain'] = float(ui_refs.gameModeGainVar.get())
    except Exception:
        pass
    try:
        data['game_mode_auto'] = bool(ui_refs.gameModeAutoVar.get())
    except Exception:
        pass
    try:
        data['monitor_timeout_ms'] = int(ui_refs.monitorTimeoutMs.get())
    except Exception:
        pass
    settings_mod.save_settings(data, SETTINGS_PATH)

def apply_settings_to_ui(settings: dict):
    if ui_refs is None:
        return
    try:
        settings_mod.apply_settings_to_ui(settings, ui_refs.playCount, ui_refs.infiniteRepeatVar, ui_refs.actionFileVar, list_action_files)
    except Exception:
        pass
    # apply game mode selection from settings
    try:
        if 'game_mode_relative' in settings:
            ui_refs.gameModeVar.set(bool(settings.get('game_mode_relative', False)))
    except Exception:
        pass
    try:
        if 'game_mode_gain' in settings:
            ui_refs.gameModeGainVar.set(float(settings.get('game_mode_gain', 1.0) or 1.0))
    except Exception:
        pass
    try:
        if 'ui_refs.gameModeAutoVar' in globals() and 'game_mode_auto' in settings:
            ui_refs.gameModeAutoVar.set(bool(settings.get('game_mode_auto', True)))
    except Exception:
        pass
    try:
        if 'monitor_timeout_ms' in settings and 'ui_refs.monitorTimeoutMs' in globals():
            ui_refs.monitorTimeoutMs.set(int(settings.get('monitor_timeout_ms', 240000)))
    except Exception:
        pass

# Sync the Combobox to point at the current recording file
def select_current_action_in_dropdown():
    if ui_refs is None:
        return
    try:
        name = os.path.basename(state.action_file_name) if state.action_file_name else ''
        files2 = list_action_files()
        # ensure list contains the current file name
        if name and name not in files2:
            files2.append(name)
        if ui_refs.actionFileSelect:
            ui_refs.actionFileSelect['values'] = files2
        if name:
            ui_refs.actionFileVar.set(name)
        elif files2:
            ui_refs.actionFileVar.set(files2[-1])
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
    # command list
    custom_thread_list = []
    print(state.can_start_listening)
    
    if ui_refs is None:
        return
    
    if state.can_start_listening and state.can_start_executing:
        if action == 'listen':
            if ui_refs.recording_controller:
                ui_refs.recording_controller.start()
            # start listen controller to capture ESC during recording
            try:
                if ui_refs.listen_controller and ui_refs.listen_controller.is_alive():
                    pass
                else:
                    ui_refs.listen_controller = ListenController(state, ui_refs)
                    ui_refs.listen_controller.start()
            except Exception:
                pass

        elif action == 'execute':
            # set the selected action file for replay
            selected = ui_refs.actionFileVar.get().strip()
            # resolve to actions/ if present
            sel_path = os.path.join('actions', selected) if selected else ''
            if not selected:
                ui_refs.startExecuteBtn['text'] = 'Select a .action file'
                return
            else:
                # use selected file (prefer actions/, fallback root)
                state.action_file_name = sel_path if os.path.exists(sel_path) else selected
            # bookkeeping for run index
            if not state.skip_run_increment:
                ui_refs.begin_run(state.action_file_name, resume=False)
            else:
                ui_refs.begin_run(state.action_file_name, resume=True)
                state.skip_run_increment = False
            # 每次启动常规脚本时重置重启标记，避免上次流程遗留导致后续超时不触发
            try:
                if os.path.basename(state.action_file_name).lower() != 'restart.action':
                    state.restarting_flag = False
            except Exception:
                pass
            # init counters and flags
            try:
                if ui_refs and bool(ui_refs.infiniteRepeatVar.get()):
                    state.ev_infinite_replay.set()
                else:
                    state.ev_infinite_replay.clear()
            except Exception:
                pass
            state.execute_time_keyboard = ui_refs.playCount.get()
            state.execute_time_mouse = ui_refs.playCount.get()
            try:
                state.ev_stop_execute_keyboard.clear()
                state.ev_stop_execute_mouse.clear()
            except Exception:
                pass
            # 确保上一轮残留按键/鼠标释放干净，避免后续按键被吞掉
            try:
                release_all_inputs()
            except Exception:
                pass
            total_ms = 0
            try:
                total_ms = compute_action_total_ms(state.action_file_name)
            except Exception:
                total_ms = 0
            try:
                start_monitor(state.execute_time_keyboard, total_ms/1000.0)
            except Exception:
                pass
            # start monitor thread (template detection) if image exists
            try:
                monitor_img = os.path.join('assets', 'monitor_target.png')
                if os.path.exists(monitor_img):
                    # stop previous monitor if any
                    try:
                        if state.monitor_thread and state.monitor_thread.is_alive():
                            state.monitor_thread.stop()
                    except Exception:
                        pass
                    def _stop_current():
                        try:
                            state.ev_stop_execute_keyboard.set()
                            state.ev_stop_execute_mouse.set()
                        except Exception:
                            pass
                    def _restart_main():
                        # run restart.action once, then resume original main action
                        try:
                            if state.restarting_flag or state.restart_running:
                                return
                            state.restarting_flag = True
                            state.pending_main_action = state.action_file_name
                            try:
                                state.pending_main_playcount = ui_refs.playCount.get()
                            except Exception:
                                state.pending_main_playcount = None
                            # stop current run; ExecuteController will start restart.action next
                            state.ev_stop_execute_keyboard.set()
                            state.ev_stop_execute_mouse.set()
                        except Exception:
                            pass
                    try:
                        timeout_s = max(1.0, float(ui_refs.monitorTimeoutMs.get())/1000.0)
                    except Exception:
                        timeout_s = 240.0  # default 4 minutes
                    state.monitor_thread = MonitorThread(
                        target_path=monitor_img,
                        timeout_s=timeout_s,
                        interval_s=3,       # check every 3s
                        stop_callbacks=[_stop_current],
                        restart_callback=_restart_main,
                        hit_callback=on_monitor_hit
                    )
                    state.monitor_thread.start()

                    # restart fallback handled inside monitor timeout flow
            except Exception:
                pass
            # UI updates
            ui_refs.update_ui_for_state(UiState.REPLAYING)
            # start replayer (keyboard + mouse) and controller (ESC monitor)
            try:
                use_rel = False
                try:
                    use_rel = bool(ui_refs.gameModeVar.get())
                except Exception:
                    pass
                try:
                    rel_gain = float(ui_refs.gameModeGainVar.get())
                except Exception:
                    try:
                        _s = load_settings(); rel_gain = float(_s.get('game_mode_gain', 1.0) or 1.0)
                    except Exception:
                        rel_gain = 1.0
                try:
                    rel_auto = bool(ui_refs.gameModeAutoVar.get())
                except Exception:
                    try:
                        _s = load_settings(); rel_auto = bool(_s.get('game_mode_auto', True))
                    except Exception:
                        rel_auto = True
                if ui_refs.playback_controller:
                    ui_refs.playback_controller.start(
                        state.action_file_name,
                        state.execute_time_keyboard,
                        infinite=bool(ui_refs.infiniteRepeatVar.get()) if ui_refs else False,
                        use_rel=use_rel,
                        rel_gain=rel_gain,
                        rel_auto=rel_auto,
                        total_ms=total_ms
                    )
            except Exception:
                pass
            try:
                if ui_refs.execute_controller and ui_refs.execute_controller.is_alive():
                    pass
                else:
                    ui_refs.execute_controller = ExecuteController(state, ui_refs, command_adapter, release_all_inputs)
                    ui_refs.execute_controller.start()
            except Exception:
                pass
            state.can_start_listening = False
            state.can_start_executing = False
    else:
        # ???F11 ???????????????
        if action == 'execute':
            try:
                state.ev_stop_execute_keyboard.set()
                state.ev_stop_execute_mouse.set()
            except Exception:
                pass
            state.can_start_listening = True
            state.can_start_executing = True
            ui_refs.update_ui_for_state(UiState.IDLE)
            ui_refs.log_event("Replay stop requested")
        else:
            try:
                ui_refs.log_event(f"Skip action '{action}': busy (state.can_start_listening={state.can_start_listening}, state.can_start_executing={state.can_start_executing})")
            except Exception:
                pass
            pass

######################################################################
# Update UI
######################################################################        
## Countdown removed — direct start/stop via F10/F11

######################################################################
######################################################################
######################################################################
# GUI
######################################################################
if __name__ == '__main__':
    import sys
    from ui.app import run_app
    run_app(sys.modules[__name__])

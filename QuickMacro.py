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

RULES_PATH = 'rules.json'

######################################################################
# Event hub and simple registries/rules
######################################################################
class EventHub:
    def __init__(self):
        self._handlers = {}

    def on(self, event: str, handler):
        if not event or handler is None:
            return
        self._handlers.setdefault(event, []).append(handler)

    def emit(self, event: str, **kwargs):
        for h in list(self._handlers.get(event, [])):
            try:
                h(**kwargs)
            except Exception:
                pass

class ActionRegistry:
    def __init__(self, base_dir='actions'):
        self.base_dir = base_dir

    def resolve(self, name: str) -> str:
        if not name:
            return ''
        cand = os.path.join(self.base_dir, name)
        if os.path.exists(cand):
            return cand
        return name if os.path.exists(name) else ''

class RuleEngine:
    """Minimal rule engine to react to events (e.g., monitor_hit) and trigger actions."""
    def __init__(self, event_hub: EventHub, runner=None, registry: ActionRegistry = None, rules: dict = None):
        self.event_hub = event_hub or EventHub()
        self.runner = runner
        self.registry = registry or ActionRegistry()
        self.rules = rules or {}
        self.last_params = None
        self.logger = None
        # Subscribe to known events
        self.event_hub.on('monitor_hit', lambda **kw: self.dispatch('monitor_hit', kw))
        self.event_hub.on('monitor_timeout', lambda **kw: self.dispatch('monitor_timeout', kw))

    def update_context(self, params: dict):
        try:
            self.last_params = dict(params or {})
        except Exception:
            self.last_params = None

    def set_rules(self, rules: dict):
        try:
            self.rules = dict(rules or {})
        except Exception:
            self.rules = {}
        # ignore invalid values silently to keep app running

    def dispatch(self, event_name: str, payload: dict = None):
        if not event_name or not self.runner:
            return
        rule = self.rules.get(event_name)
        if not rule:
            return
        if isinstance(rule, str):
            action_name = rule
            overrides = {}
        elif isinstance(rule, dict):
            action_name = rule.get('action') or ''
            overrides = {k:v for k,v in rule.items() if k != 'action'}
        else:
            return
        path = self.registry.resolve(action_name)
        if not path:
            return
        params = dict(self.last_params or {})
        params.update(overrides)
        params['action'] = action_name
        if not params.get('repeat'):
            params['repeat'] = 1
        try:
            if callable(getattr(self.runner, '_log', None)):
                self.runner._log(f"Rule hit: {event_name} -> {action_name}")
            self.runner.start(path, params, resume=False)
        except Exception:
            pass

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
    restart_timeout_ms: int = 3600000
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
app_service = None

# UI references container
@dataclass
class UIRefs:
    root: object
    actionFileVar: object
    actionFileSelect: object
    playCount: object
    infiniteRepeatVar: object
    gameModeVar: object
    gameModeGainVar: object
    gameModeAutoVar: object
    log_event: object
    update_ui_for_state: object
    begin_run: object
    mark_interrupted: object
    mark_finished: object
    recording_controller: object
    playback_controller: object
    listen_controller: object
    execute_controller: object

# Simple UI state enum for readability
class UiState:
    IDLE = 'idle'
    RECORDING = 'recording'
    REPLAYING = 'replaying'

######################################################################
# Service layer
######################################################################
class AppService:
    def __init__(self, state: AppState, recording_controller, playback_controller, listen_controller, execute_controller, *, start_monitor, compute_action_total_ms, get_restart_timeout_ms, release_all_inputs, hooks: dict, replay_params_provider, execute_controller_factory=None, event_hub=None):
        self.state = state
        self.recording_controller = recording_controller
        self.playback_controller = playback_controller
        self.listen_controller = listen_controller
        self.execute_controller = execute_controller
        self.start_monitor_fn = start_monitor
        self.compute_action_total_ms = compute_action_total_ms
        self.get_restart_timeout_ms = get_restart_timeout_ms
        self.release_all_inputs = release_all_inputs
        self.hooks = hooks or {}
        self.replay_params_provider = replay_params_provider
        self.last_replay_params = None
        self._on_monitor_hit = self.hooks.get('on_monitor_hit')
        self.event_hub = event_hub or EventHub()
        self._execute_controller_factory = execute_controller_factory
        self.action_registry = ActionRegistry()
        self.rule_engine = RuleEngine(self.event_hub, runner=None, registry=self.action_registry, rules={})
        # runner init deferred until class is defined below

    def _log(self, msg: str):
        try:
            cb = self.hooks.get('log_event')
            if cb:
                cb(msg)
        except Exception:
            pass

    def _update_state(self, ui_state: str):
        try:
            cb = self.hooks.get('update_ui_for_state')
            if cb:
                cb(ui_state)
        except Exception:
            pass

    def _begin_run(self, path: str, resume: bool):
        try:
            cb = self.hooks.get('begin_run')
            if cb:
                cb(path, resume=resume)
        except Exception:
            pass

    def _resolve_action_path(self, name: str) -> str:
        if not name:
            return ''
        sel_path = os.path.join('actions', name)
        return sel_path if os.path.exists(sel_path) else name

    def _ensure_execute_controller(self):
        if self.execute_controller and self.execute_controller.is_alive():
            return self.execute_controller
        if callable(self._execute_controller_factory):
            self.execute_controller = self._execute_controller_factory()
            return self.execute_controller
        return None

    def start_replay(self, params: dict = None, *, resume: bool = False):
        if not (self.state.can_start_listening and self.state.can_start_executing):
            return
        if params is None and self.replay_params_provider:
            params = self.replay_params_provider()
        if not params:
            return
        name = (params.get('action') or '').strip()
        if not name:
            self._log('Please select an action file.')
            return
        action_path = self._resolve_action_path(name)
        if not action_path:
            self._log('Action file not found.')
            return
        self.last_replay_params = params
        if not hasattr(self, 'runner') or self.runner is None:
            self.runner = ActionRunner(
                state=self.state,
                playback_controller=self.playback_controller,
                start_monitor=self.start_monitor_fn,
                compute_action_total_ms=self.compute_action_total_ms,
                get_restart_timeout_ms=self.get_restart_timeout_ms,
                release_all_inputs=self.release_all_inputs,
                execute_controller_supplier=self._ensure_execute_controller,
                hooks=self.hooks,
                event_hub=self.event_hub,
            )
            self.rule_engine.runner = self.runner
        self.rule_engine.update_context(params)
        self.runner.start(action_path, params, resume=resume)

    def stop_replay(self):
        if hasattr(self, 'runner') and self.runner:
            self.runner.stop()

    def toggle_record(self):
        if self.state.can_start_listening and self.state.can_start_executing:
            self.start_record()
        else:
            self.stop_record()

    def toggle_replay(self):
        if self.state.can_start_listening and self.state.can_start_executing:
            params = self.replay_params_provider() if self.replay_params_provider else None
            self.start_replay(params)
        else:
            self.stop_replay()


class ActionRunner:
    def __init__(self, state: AppState, playback_controller, start_monitor, compute_action_total_ms, get_restart_timeout_ms, release_all_inputs, execute_controller_supplier, hooks: dict, event_hub: EventHub):
        self.state = state
        self.playback_controller = playback_controller
        self.start_monitor = start_monitor
        self.compute_action_total_ms = compute_action_total_ms
        self.get_restart_timeout_ms = get_restart_timeout_ms
        self.release_all_inputs = release_all_inputs
        self.execute_controller_supplier = execute_controller_supplier
        self.hooks = hooks or {}
        self.event_hub = event_hub or EventHub()
        self._log = self.hooks.get('log_event')
        self._update_ui = self.hooks.get('update_ui_for_state')
        self._begin_run = self.hooks.get('begin_run')
        self._mark_interrupted = self.hooks.get('mark_interrupted')
        self._mark_finished = self.hooks.get('mark_finished')
        self._on_monitor_hit = self.hooks.get('on_monitor_hit')

    def start(self, action_path: str, params: dict, resume: bool = False):
        state = self.state
        state.action_file_name = action_path
        try:
            state.restart_timeout_ms = self.get_restart_timeout_ms(action_path)
        except Exception:
            state.restart_timeout_ms = 3600000
        if not state.skip_run_increment:
            if callable(self._begin_run):
                self._begin_run(state.action_file_name, resume=False)
        else:
            if callable(self._begin_run):
                self._begin_run(state.action_file_name, resume=True)
            state.skip_run_increment = False
        try:
            if os.path.basename(state.action_file_name).lower() != 'restart.action':
                state.restarting_flag = False
        except Exception:
            pass
        try:
            if params.get('infinite'):
                state.ev_infinite_replay.set()
            else:
                state.ev_infinite_replay.clear()
        except Exception:
            pass
        try:
            repeat = int(params.get('repeat') or 1)
        except Exception:
            repeat = 1
        state.execute_time_keyboard = repeat
        state.execute_time_mouse = repeat
        try:
            state.ev_stop_execute_keyboard.clear()
            state.ev_stop_execute_mouse.clear()
        except Exception:
            pass
        try:
            self.release_all_inputs()
        except Exception:
            pass
        total_ms = 0
        try:
            total_ms = self.compute_action_total_ms(state.action_file_name)
        except Exception:
            total_ms = 0
        try:
            self.start_monitor(repeat, total_ms/1000.0)
        except Exception:
            pass
        try:
            monitor_img = os.path.join('assets', 'monitor_target.png')
            if os.path.exists(monitor_img):
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
                    try:
                        self.event_hub.emit('monitor_timeout', action=state.action_file_name)
                    except Exception:
                        pass
                    try:
                        if state.restarting_flag or state.restart_running:
                            return
                        state.restarting_flag = True
                        state.pending_main_action = state.action_file_name
                        state.pending_main_playcount = repeat
                        state.ev_stop_execute_keyboard.set()
                        state.ev_stop_execute_mouse.set()
                    except Exception:
                        pass
                try:
                    timeout_s = max(1.0, float(state.restart_timeout_ms)/1000.0)
                except Exception:
                    timeout_s = 3600.0
                state.monitor_thread = MonitorThread(
                    target_path=monitor_img,
                    timeout_s=timeout_s,
                    interval_s=3,
                    stop_callbacks=[_stop_current],
                    restart_callback=_restart_main,
                    hit_callback=self._handle_monitor_hit
                )
                state.monitor_thread.start()
        except Exception:
            pass
        if callable(self._update_ui):
            self._update_ui(UiState.REPLAYING)
        try:
            use_rel = bool(params.get('use_rel'))
        except Exception:
            use_rel = False
        try:
            rel_gain = float(params.get('rel_gain', 1.0) or 1.0)
        except Exception:
            rel_gain = 1.0
        try:
            rel_auto = bool(params.get('rel_auto', True))
        except Exception:
            rel_auto = True
        if self.playback_controller:
            try:
                self.playback_controller.start(
                    state.action_file_name,
                    repeat,
                    infinite=bool(params.get('infinite')),
                    use_rel=use_rel,
                    rel_gain=rel_gain,
                    rel_auto=rel_auto,
                    total_ms=total_ms
                )
            except Exception:
                pass
        try:
            ctrl = self.execute_controller_supplier() if callable(self.execute_controller_supplier) else None
            if ctrl and not ctrl.is_alive():
                ctrl.start()
        except Exception:
            pass
        state.can_start_listening = False
        state.can_start_executing = False
        try:
            self.event_hub.emit('replay_started', action=state.action_file_name, repeat=repeat)
        except Exception:
            pass

    def stop(self):
        try:
            self.state.ev_stop_execute_keyboard.set()
            self.state.ev_stop_execute_mouse.set()
        except Exception:
            pass
        self.state.can_start_listening = True
        self.state.can_start_executing = True
        try:
            if callable(self._update_ui):
                self._update_ui(UiState.IDLE)
        except Exception:
            pass
        if callable(self._log):
            self._log("Replay stop requested")
        try:
            self.event_hub.emit('replay_stopped', action=self.state.action_file_name)
        except Exception:
            pass

    def _handle_monitor_hit(self, **kwargs):
        try:
            if callable(self._on_monitor_hit):
                self._on_monitor_hit()
        except Exception:
            pass
        try:
            self.event_hub.emit('monitor_hit', action=self.state.action_file_name, **kwargs)
        except Exception:
            pass

    def start_record(self):
        if not (self.state.can_start_listening and self.state.can_start_executing):
            return
        if self.recording_controller:
            self.recording_controller.start()
        try:
            if self.listen_controller and self.listen_controller.is_alive():
                pass
            else:
                self.listen_controller = ListenController(self.state, ui_refs)
                self.listen_controller.start()
        except Exception:
            pass

    def stop_record(self):
        if self.recording_controller:
            try:
                self.recording_controller.stop()
            except Exception:
                pass

    def start_replay(self, params: dict = None, *, resume: bool = False):
        if not (self.state.can_start_listening and self.state.can_start_executing):
            return
        if params is None and self.replay_params_provider:
            params = self.replay_params_provider()
        if not params:
            return
        self.last_replay_params = params
        selected = (params.get('action') or '').strip()
        sel_path = self._resolve_action_path(selected)
        if not selected:
            self._log('Please select an action file.')
            return
        self.state.action_file_name = sel_path
        try:
            self.state.restart_timeout_ms = self.get_restart_timeout_ms(sel_path)
        except Exception:
            self.state.restart_timeout_ms = 3600000
        if not self.state.skip_run_increment:
            self._begin_run(self.state.action_file_name, resume=False)
        else:
            self._begin_run(self.state.action_file_name, resume=True)
            self.state.skip_run_increment = False
        try:
            if os.path.basename(self.state.action_file_name).lower() != 'restart.action':
                self.state.restarting_flag = False
        except Exception:
            pass
        try:
            if params.get('infinite'):
                self.state.ev_infinite_replay.set()
            else:
                self.state.ev_infinite_replay.clear()
        except Exception:
            pass
        try:
            repeat = int(params.get('repeat') or 1)
        except Exception:
            repeat = 1
        self.state.execute_time_keyboard = repeat
        self.state.execute_time_mouse = repeat
        try:
            self.state.ev_stop_execute_keyboard.clear()
            self.state.ev_stop_execute_mouse.clear()
        except Exception:
            pass
        try:
            self.release_all_inputs()
        except Exception:
            pass
        total_ms = 0
        try:
            total_ms = self.compute_action_total_ms(self.state.action_file_name)
        except Exception:
            total_ms = 0
        try:
            self.start_monitor_fn(self.state.execute_time_keyboard, total_ms/1000.0)
        except Exception:
            pass
        try:
            monitor_img = os.path.join('assets', 'monitor_target.png')
            if os.path.exists(monitor_img):
                try:
                    if self.state.monitor_thread and self.state.monitor_thread.is_alive():
                        self.state.monitor_thread.stop()
                except Exception:
                    pass
                def _stop_current():
                    try:
                        self.state.ev_stop_execute_keyboard.set()
                        self.state.ev_stop_execute_mouse.set()
                    except Exception:
                        pass
                def _restart_main():
                    try:
                        if self.state.restarting_flag or self.state.restart_running:
                            return
                        self.state.restarting_flag = True
                        self.state.pending_main_action = self.state.action_file_name
                        self.state.pending_main_playcount = repeat
                        self.state.ev_stop_execute_keyboard.set()
                        self.state.ev_stop_execute_mouse.set()
                    except Exception:
                        pass
                try:
                    timeout_s = max(1.0, float(self.state.restart_timeout_ms)/1000.0)
                except Exception:
                    timeout_s = 3600.0
                self.state.monitor_thread = MonitorThread(
                    target_path=monitor_img,
                    timeout_s=timeout_s,
                    interval_s=3,
                    stop_callbacks=[_stop_current],
                    restart_callback=_restart_main,
                    hit_callback=self._on_monitor_hit
                )
                self.state.monitor_thread.start()
        except Exception:
            pass
        self._update_state(UiState.REPLAYING)
        try:
            use_rel = bool(params.get('use_rel'))
        except Exception:
            use_rel = False
        try:
            rel_gain = float(params.get('rel_gain', 1.0) or 1.0)
        except Exception:
            rel_gain = 1.0
        try:
            rel_auto = bool(params.get('rel_auto', True))
        except Exception:
            rel_auto = True
        if self.playback_controller:
            try:
                self.playback_controller.start(
                    self.state.action_file_name,
                    repeat,
                    infinite=bool(params.get('infinite')),
                    use_rel=use_rel,
                    rel_gain=rel_gain,
                    rel_auto=rel_auto,
                    total_ms=total_ms
                )
            except Exception:
                pass
        try:
            if self.execute_controller and self.execute_controller.is_alive():
                pass
            else:
                self.execute_controller = ExecuteController(self.state, ui_refs, command_adapter, self.release_all_inputs)
                self.execute_controller.start()
        except Exception:
            pass
        self.state.can_start_listening = False
        self.state.can_start_executing = False

    def stop_replay(self):
        try:
            self.state.ev_stop_execute_keyboard.set()
            self.state.ev_stop_execute_mouse.set()
        except Exception:
            pass
        self.state.can_start_listening = True
        self.state.can_start_executing = True
        self._update_state(UiState.IDLE)
        self._log("Replay stop requested")

    def toggle_record(self):
        if self.state.can_start_listening and self.state.can_start_executing:
            self.start_record()
        else:
            self.stop_record()

    def toggle_replay(self):
        if self.state.can_start_listening and self.state.can_start_executing:
            params = self.replay_params_provider() if self.replay_params_provider else None
            self.start_replay(params)
        else:
            self.stop_replay()
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
    """Windows: try to elevate; if spawn succeeds, exit current process so只保留提权后的 UI."""
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
        if rc <= 32:
            logging.warning(f'UAC elevation failed, rc={rc}.')
            _win_message_box('Need admin', 'Could not elevate automatically. Please run as administrator for better compatibility.', 0x30)
        else:
            logging.info('UAC elevation triggered; exiting current instance to keep only elevated UI.')
            os._exit(0)
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

def load_rules(path: str = RULES_PATH) -> dict:
    try:
        with open(path, 'r', encoding='utf-8') as f:
            import json
            data = json.load(f) or {}
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {}

def compute_action_total_ms(path: str) -> int:
    try:
        total_ms = 0
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                s = line.strip()
                if not s or s.startswith('#'):
                    continue
                d = action_parse_line(s)
                try:
                    t_ms = int(float(d.get('ms') or 0))
                except Exception:
                    t_ms = 0
                if t_ms > 0:
                    total_ms = max(total_ms, t_ms)
        return total_ms
    except Exception:
        return 0

def _get_default_restart_ms():
    try:
        cfg = load_settings() or {}
        val = int(float(cfg.get('default_restart_ms', 3600000)))
        return max(1000, val)
    except Exception:
        return 3600000

def get_restart_timeout_ms(path: str) -> int:
    default_ms = _get_default_restart_ms()
    if not path:
        return default_ms
    try:
        with open(path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except Exception:
        return default_ms
    restart_ms = None
    for line in lines:
        s = line.strip()
        if not s.startswith('META RESTART'):
            continue
        parts = s.split()
        if len(parts) >= 3:
            try:
                restart_ms = int(float(parts[2]))
            except Exception:
                restart_ms = None
        break
    if restart_ms is None:
        restart_ms = default_ms
    try:
        restart_ms = int(restart_ms)
    except Exception:
        restart_ms = default_ms
    return max(1000, restart_ms)

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
    global app_service
    if app_service is None:
        return
    if action == 'listen':
        app_service.toggle_record()
    elif action == 'execute':
        app_service.toggle_replay()

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

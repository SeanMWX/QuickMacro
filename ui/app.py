import os
import logging
import threading
import time
from datetime import datetime
import tkinter
from tkinter import ttk
import tkinter.font as tkfont
import tkinter.messagebox as messagebox
from tkinter import scrolledtext
from pathlib import Path

def run_app(qm):
    # bind symbols from QuickMacro module
    state = qm.state
    command_adapter = qm.command_adapter
    relaunch_as_admin_if_needed = qm.relaunch_as_admin_if_needed
    set_process_dpi_aware = qm.set_process_dpi_aware
    ensure_assets_dir = qm.ensure_assets_dir
    ensure_actions_dir = qm.ensure_actions_dir
    init_new_action_file = qm.init_new_action_file
    list_action_files = qm.list_action_files
    release_all_inputs = qm.release_all_inputs
    apply_settings_to_ui = qm.apply_settings_to_ui
    save_settings = qm.save_settings
    load_settings = qm.load_settings
    select_current_action_in_dropdown = qm.select_current_action_in_dropdown
    compute_action_total_ms = qm.compute_action_total_ms
    get_restart_timeout_ms = qm.get_restart_timeout_ms if 'get_restart_timeout_ms' in vars(qm) else None
    action_parse_line = qm.action_parse_line
    action_compose_line = qm.action_compose_line
    UiState = qm.UiState if 'UiState' in vars(qm) else qm.__dict__.get('UiState', None)
    RecordingController = qm.RecordingController
    PlaybackController = qm.PlaybackController
    MonitorThread = qm.MonitorThread
    HotkeyController = qm.HotkeyController
    ListenController = qm.ListenController
    ExecuteController = qm.ExecuteController
    UIRefs = qm.UIRefs
    # helper assignments for globals compatibility
    globals_ref = qm.__dict__

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

    state.can_start_listening = True
    state.can_start_executing = True
    state.execute_time_keyboard = 0
    state.execute_time_mouse = 0
    # threading Events for coordination
    state.ev_stop_execute_keyboard = threading.Event(); state.ev_stop_execute_keyboard.set()
    state.ev_stop_execute_mouse = threading.Event(); state.ev_stop_execute_mouse.set()
    state.ev_stop_listen = threading.Event(); state.ev_stop_listen.set()
    state.ev_infinite_replay = threading.Event(); state.ev_infinite_replay.clear()
    state.pressed_vks = set()
    state.pressed_mouse_buttons = set()
    
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
    root.geometry('720x680')
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
    replayCard.place(x=360, y=70, width=330, height=280)
    monitorCard = ttk.Frame(root, style='Card.TFrame', borderwidth=1, relief='solid')
    monitorCard.place(x=30, y=360, width=660, height=110)
    logCard = ttk.Frame(root, style='Card.TFrame', borderwidth=1, relief='solid')
    logCard.place(x=30, y=480, width=660, height=180)
    
    # Log area
    logLabel = ttk.Label(logCard, text='Logs', style='CardLabel.TLabel')
    logLabel.place(x=12, y=8, width=80, height=24)
    logText = scrolledtext.ScrolledText(logCard, wrap='word', height=8, state='disabled', font=(font_family, 10))
    logText.place(x=12, y=36, width=636, height=130)
    
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
    restartTimeLabel = ttk.Label(monitorCard, text='Restart time: 3600.0s', style='CardLabel.TLabel')
    restartTimeLabel.place(x=220, y=70, width=260, height=24)
    # Monitor state & helpers
    state.monitor_total_loops = 0
    state.monitor_completed_loops = 0
    state.monitor_loop_start_ts = None
    state.monitor_timer_job = None
    state.monitor_total_time_s = 0.0
    state.dungeon_start_ts = None
    state.restart_timeout_ms = getattr(state, 'restart_timeout_ms', 3600000)
    state.current_run_idx = 0
    state.current_run_action = ''
    state.current_run_interrupted = False
    state.skip_run_increment = False
    recording_controller = None
    playback_controller = None
    state.current_run_idx = 0
    state.current_run_action = ''
    state.current_run_interrupted = False
    state.skip_run_increment = False
    def _get_replay_params():
        try:
            gain_val = float(gameModeGainVar.get())
        except Exception:
            gain_val = 1.0
        return {
            'action': actionFileVar.get().strip(),
            'repeat': playCount.get(),
            'infinite': bool(infiniteRepeatVar.get()),
            'use_rel': bool(gameModeVar.get()),
            'rel_gain': gain_val,
            'rel_auto': bool(gameModeAutoVar.get())
        }
    def _resolve_action_path(name: str) -> str:
        if not name:
            return ''
        candidate = os.path.join('actions', name)
        return candidate if os.path.exists(candidate) else name
    def log_event(msg: str):
        try:
            ts = datetime.now().strftime('%H:%M:%S')
            logText.configure(state='normal')
            logText.insert('end', f"[{ts}] {msg}\n")
            logText.see('end')
            logText.configure(state='disabled')
        except Exception:
            pass
    
    # Run bookkeeping helpers
    def begin_run(action_path: str, resume: bool = False):
        """记录一次新的播放或恢复"""
        name = os.path.basename(action_path) if action_path else ''
        if resume:
            state.skip_run_increment = False
            state.current_run_action = state.current_run_action or name
            state.current_run_interrupted = False
            log_event(f"Run #{state.current_run_idx} resume: {state.current_run_action}")
            return
        state.current_run_idx += 1
        state.current_run_action = name
        state.current_run_interrupted = False
        log_event(f"Run #{state.current_run_idx} start: {state.current_run_action}")
    
    def mark_interrupted(reason: str = ''):
        """标记当前 run 被中断"""
        state.current_run_interrupted = True
        log_event(f"Run #{state.current_run_idx} interrupted{': ' + reason if reason else ''}")
    
    def mark_finished():
        """标记正常结束"""
        if not state.current_run_interrupted:
            log_event(f"Run #{state.current_run_idx} finished: {state.current_run_action}")
    
    def update_monitor_labels():
        try:
            loops_text = f"Loops: {state.monitor_completed_loops}/{state.monitor_total_loops}" if state.monitor_total_loops else "Loops: 0/0"
            monitorLoopLabel['text'] = loops_text
            elapsed = 0.0
            if state.monitor_loop_start_ts is not None:
                elapsed = max(0.0, time.monotonic() - state.monitor_loop_start_ts)
            monitorTimeLabel['text'] = f"Current loop time: {elapsed:.1f}s"
            monitorTotalLabel['text'] = f"Total loop time: {state.monitor_total_time_s:.1f}s"
            dungeon_elapsed = 0.0
            if state.dungeon_start_ts is not None:
                dungeon_elapsed = max(0.0, time.monotonic() - state.dungeon_start_ts)
            dungeonTimeLabel['text'] = f"Dungeon time: {dungeon_elapsed:.1f}s"
            restart_ms = getattr(state, 'restart_timeout_ms', 0) or 0
            restartTimeLabel['text'] = f"Restart time: {restart_ms/1000.0:.1f}s"
        except Exception:
            pass

    def refresh_restart_timeout_from_selection(*_):
        try:
            path = _resolve_action_path(actionFileVar.get().strip())
            if get_restart_timeout_ms:
                state.restart_timeout_ms = get_restart_timeout_ms(path)
            else:
                state.restart_timeout_ms = getattr(state, 'restart_timeout_ms', 3600000) or 3600000
        except Exception:
            state.restart_timeout_ms = 3600000
        update_monitor_labels()

    def _tick_monitor():
        update_monitor_labels()
        # auto restart if dungeon time exceeds configured timeout
        try:
            timeout_ms = getattr(state, 'restart_timeout_ms', 0)
            if timeout_ms and timeout_ms > 0 and state.dungeon_start_ts is not None:
                dungeon_elapsed = max(0.0, time.monotonic() - state.dungeon_start_ts)
                if dungeon_elapsed >= (timeout_ms / 1000.0):
                    # trigger restart flow once
                    if not state.restarting_flag and not state.restart_running:
                        state.restarting_flag = True
                        state.pending_main_action = state.action_file_name
                        try:
                            state.pending_main_playcount = playCount.get()
                        except Exception:
                            state.pending_main_playcount = None
                        actionFileVar.set('restart.action')
                        state.ev_stop_execute_keyboard.set()
                        state.ev_stop_execute_mouse.set()
        except Exception:
            pass
        try:
            state.monitor_timer_job = root.after(200, _tick_monitor)
        except Exception:
            state.monitor_timer_job = None
    
    def start_monitor(total_loops: int, total_time_s: float = 0.0):
        state.monitor_total_loops = max(1, int(total_loops or 1))
        state.monitor_completed_loops = 0
        state.monitor_loop_start_ts = time.monotonic()
        # start dungeon timer on first monitor start; refreshed again when target hits
        if state.dungeon_start_ts is None:
            state.dungeon_start_ts = time.monotonic()
        try:
            state.monitor_total_time_s = max(0.0, float(total_time_s or 0.0))
        except Exception:
            state.monitor_total_time_s = 0.0
        if state.monitor_timer_job:
            try:
                root.after_cancel(state.monitor_timer_job)
            except Exception:
                pass
            state.monitor_timer_job = None
        update_monitor_labels()
        _tick_monitor()
    
    def update_replay_progress(done: int, total: int):
        try:
            state.monitor_completed_loops = int(done or 0)
        except Exception:
            state.monitor_completed_loops = done
        try:
            if total:
                state.monitor_total_loops = int(total)
        except Exception:
            pass
        update_monitor_labels()
    
    def update_replay_loop_start(loop_idx: int, total: int):
        try:
            state.monitor_total_loops = int(total or state.monitor_total_loops or 0)
        except Exception:
            pass
        state.monitor_loop_start_ts = time.monotonic()
        update_monitor_labels()
    
    # Monitor thread handle
    state.monitor_thread = None
    state.pending_main_action = None
    state.pending_main_playcount = None
    state.restarting_flag = False
    state.restart_back_job = None
    
    def on_monitor_hit():
        # mark dungeon start at first detection of target image
        state.dungeon_start_ts = time.monotonic()
        update_monitor_labels()
    
    def reset_monitor():
        state.monitor_total_loops = 0
        state.monitor_completed_loops = 0
        state.monitor_loop_start_ts = None
        state.dungeon_start_ts = None
        if state.monitor_timer_job:
            try:
                root.after_cancel(state.monitor_timer_job)
            except Exception:
                pass
            state.monitor_timer_job = None
        update_monitor_labels()
    
    # UI state helper
    class UiState:
        IDLE = 'idle'
        RECORDING = 'recording'
        REPLAYING = 'replaying'
    
    def update_ui_for_state(ui_state: str):
        if ui_state in (UiState.IDLE, 'idle'):
            startListenerBtn.state(['!disabled'])
            startExecuteBtn.state(['!disabled'])
            startListenerBtn['text'] = 'Start recording (F10)'
            startExecuteBtn['text'] = 'Start replaying (F11)'
            reset_monitor()
        elif ui_state in (UiState.RECORDING, 'recording'):
            startListenerBtn.state(['disabled'])
            startExecuteBtn.state(['disabled'])
            startListenerBtn['text'] = 'Recording, "F10" to stop.'
        elif ui_state in (UiState.REPLAYING, 'replaying'):
            startListenerBtn.state(['disabled'])
            startExecuteBtn.state(['disabled'])
            startExecuteBtn['text'] = 'Replaying, "F11" to stop.'
        # ensure monitor thread is stopped when exiting replay
        if ui_state in (UiState.IDLE, 'idle'):
            try:
                if state.monitor_thread:
                    state.monitor_thread.stop()
            except Exception:
                pass
        try:
            root.update_idletasks()
        except Exception:
            pass
    
    # controllers
    recording_controller = RecordingController(
        state=state,
        init_new_action_file=init_new_action_file,
        select_current_action_in_dropdown=select_current_action_in_dropdown,
        update_ui_for_state=update_ui_for_state,
        logger=log_event
    )
    playback_controller = PlaybackController(
        state=state,
        update_ui_for_state=update_ui_for_state,
        release_all_inputs=release_all_inputs,
        start_monitor=start_monitor,
        on_progress=update_replay_progress,
        on_loop_start=update_replay_loop_start,
        logger=log_event
    )
    listen_controller = ListenController(state, None)  # ui_refs patched later
    execute_controller = ExecuteController(state, None, command_adapter, release_all_inputs)
    
    actionFileVar = tkinter.StringVar()
    files = list_action_files()
    actionFileVar.set(files[-1] if files else '')
    
    # Action file controls inside the Replay card
    actionFileLabel = ttk.Label(replayCard, text='Action file', style='CardLabel.TLabel')
    actionFileLabel.place(x=15, y=170, width=100, height=26)
    actionFileSelect = ttk.Combobox(replayCard, textvariable=actionFileVar, values=files if files else [], state='readonly', style='Biz.TCombobox')

    # expose key UI refs back to qm module via UIRefs container
    qm.ui_refs = UIRefs(
        root=root,
        actionFileVar=actionFileVar,
        actionFileSelect=actionFileSelect,
        startExecuteBtn=startExecuteBtn,
        startListenerBtn=startListenerBtn,
        playCount=playCount,
        infiniteRepeatVar=infiniteRepeatVar,
        gameModeVar=gameModeVar,
        gameModeGainVar=gameModeGainVar,
        gameModeAutoVar=gameModeAutoVar,
        log_event=log_event,
        update_ui_for_state=update_ui_for_state,
        begin_run=begin_run,
        mark_interrupted=mark_interrupted,
        mark_finished=mark_finished,
        recording_controller=recording_controller,
        playback_controller=playback_controller,
        listen_controller=listen_controller,
        execute_controller=execute_controller,
    )
    listen_controller.ui_refs = qm.ui_refs
    execute_controller.ui_refs = qm.ui_refs

    app_service = qm.AppService(
        state=state,
        recording_controller=recording_controller,
        playback_controller=playback_controller,
        listen_controller=listen_controller,
        execute_controller=execute_controller,
        start_monitor=start_monitor,
        compute_action_total_ms=compute_action_total_ms,
        get_restart_timeout_ms=get_restart_timeout_ms,
        release_all_inputs=release_all_inputs,
        hooks={
            'log_event': log_event,
            'update_ui_for_state': update_ui_for_state,
            'begin_run': begin_run,
            'mark_interrupted': mark_interrupted,
            'mark_finished': mark_finished,
        },
        replay_params_provider=_get_replay_params
    )
    qm.app_service = app_service


    actionFileSelect.place(x=120, y=170, width=190, height=28)
    
    # Refresh button removed; list auto-updates after recording
    
    # 加载设置并应用
    try:
        _settings = load_settings()
        apply_settings_to_ui(_settings)
    except Exception:
        pass
    refresh_restart_timeout_from_selection()
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
    def _on_action_selected(*_):
        try:
            save_settings()
        except Exception:
            pass
        refresh_restart_timeout_from_selection()
    try:
        actionFileSelect.bind('<<ComboboxSelected>>', _on_action_selected)
    except Exception:
        pass
    try:
        actionFileVar.trace_add('write', refresh_restart_timeout_from_selection)
    except Exception:
        try:
            actionFileVar.trace('w', refresh_restart_timeout_from_selection)
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
        if not (state.can_start_listening and state.can_start_executing):
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
    editBtn.place(x=120, y=235, width=80, height=28)
    openBtn = ttk.Button(replayCard, text='Folder', command=open_actions_folder, style='Biz.TButton')
    openBtn.place(x=15, y=235, width=100, height=28)

    # Start hotkeys listener (F10/F11)
    HotkeyController(state, root, lambda: command_adapter('listen'), lambda: command_adapter('execute')).start()
    
    # Removed Tk window key binds to avoid double-trigger; global hotkeys handle F10/F11
    
    # Ensure closing window terminates app
    def on_close():
        try:
            state.ev_stop_execute_keyboard.set()
            state.ev_stop_execute_mouse.set()
            state.ev_stop_listen.set()
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
    

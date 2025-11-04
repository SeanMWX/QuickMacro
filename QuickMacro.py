import json
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

from pynput import keyboard, mouse
from pynput.keyboard import Controller as KeyBoardController, KeyCode, Key
from pynput.mouse import Button, Controller as MouseController

######################################################################
# Helpers
######################################################################
def get_screen_size():
    try:
        user32 = ctypes.windll.user32
        return int(user32.GetSystemMetrics(0)), int(user32.GetSystemMetrics(1))
    except Exception:
        try:
            # Fallback via tkinter if available
            t = tkinter.Tk()
            w = int(t.winfo_screenwidth())
            h = int(t.winfo_screenheight())
            t.destroy()
            return w, h
        except Exception:
            return 1920, 1080

def set_process_dpi_aware():
    try:
        ctypes.windll.user32.SetProcessDPIAware()
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

record_write_lock = threading.Lock()

def write_action_line(line):
    global action_file_name
    try:
        with record_write_lock:
            with open(action_file_name, 'a', encoding='utf-8') as f:
                f.write(line + "\n")
    except Exception:
        pass

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
                'Replace these with your own cute/moe assets.\n',
                encoding='utf-8'
            )
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
    except Exception:
        pass
    try:
        ms = MouseController()
        if 'left' in pressed_mouse_buttons:
            try:
                ms.release(Button.left)
            except Exception:
                pass
            pressed_mouse_buttons.discard('left')
        if 'right' in pressed_mouse_buttons:
            try:
                ms.release(Button.right)
            except Exception:
                pass
            pressed_mouse_buttons.discard('right')
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
    global stop_execute_keyboard
    global stop_execute_mouse
    
    # command list
    custom_thread_list = []
    print(can_start_listening)
    
    if can_start_listening and can_start_executing:
        if action == 'listen':
            # setup shared action file and start time
            init_new_action_file()
            # update UI selection to the new file
            select_current_action_in_dropdown()
            # UI updates
            startListenerBtn['state'] = 'disabled'
            startExecuteBtn['state'] = 'disabled'
            startListenerBtn['text'] = 'Recording, "ESC/F10" to stop.'
            # threads
            custom_thread_list = [
                {'obj_thread': ListenController()},
                {'obj_thread': MouseActionListener()},
                {'obj_thread': KeyboardActionListener()}
            ]
            for t in custom_thread_list:
                t['obj_thread'].start()
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
                global action_file_name
                action_file_name = sel_path if os.path.exists(sel_path) else selected
            # init counters and flags
            global infinite_replay
            infinite_replay = bool(infiniteRepeatVar.get()) if 'infiniteRepeatVar' in globals() else False
            execute_time_keyboard = playCount.get()
            execute_time_mouse = playCount.get()
            stop_execute_keyboard = False
            stop_execute_mouse = False
            # UI updates
            startListenerBtn['state'] = 'disabled'
            startExecuteBtn['state'] = 'disabled'
            startExecuteBtn['text'] = 'Replaying, "ESC/F11" to stop.'
            # threads
            custom_thread_list = [
                {'obj_thread': ExecuteController()},
                {'obj_thread': MouseActionExecute()},
                {'obj_thread': KeyboardActionExecute()}
            ]
            for t in custom_thread_list:
                t['obj_thread'].start()
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
        # press keyboard
        def on_press(key):
            global record_start_time
            # ignore control hotkeys and ESC in recording
            if key in (keyboard.Key.esc, keyboard.Key.f10, keyboard.Key.f11):
                return
            t = int((time.time() - record_start_time) * 1000)
            try:
                vk = key.vk
            except AttributeError:
                vk = key.value.vk
            write_action_line(f"K DOWN {vk} {t}")

        # release keyboard
        def on_release(key):
            global can_start_listening
            global can_start_executing
            global stop_listen
            global record_start_time

            if key == keyboard.Key.esc:
                # Stop by pressing "ESC"
                stop_listen = True
                keyboardListener.stop()
                return False

            if not stop_listen:
                # ignore control hotkeys in recording
                if key in (keyboard.Key.f10, keyboard.Key.f11):
                    return
                t = int((time.time() - record_start_time) * 1000)
                try:
                    vk = key.vk
                except AttributeError:
                    vk = key.value.vk
                write_action_line(f"K UP {vk} {t}")
        
        with keyboard.Listener(on_press=on_press, on_release=on_release) as keyboardListener:
            keyboardListener.join()          
                

class MouseActionListener(threading.Thread):

    def __init__(self, file_name='mouse.action'):
        super().__init__()
        self.daemon = True
        self.file_name = file_name

    def run(self):
        # record-time screen size
        sw, sh = get_screen_size()
        # move mouse
        def on_move(x, y):
            global stop_listen, record_start_time
            if stop_listen:
                mouseListener.stop()
            t = int((time.time() - record_start_time) * 1000)
            try:
                nx = float(x) / float(sw)
                ny = float(y) / float(sh)
            except Exception:
                nx = None
                ny = None
            if nx is not None and ny is not None:
                write_action_line(f"M MOVE {int(x)} {int(y)} {nx:.6f} {ny:.6f} {t}")
            else:
                write_action_line(f"M MOVE {int(x)} {int(y)} {t}")

        # click mouse
        def on_click(x, y, button, pressed):
            global stop_listen, record_start_time
            if stop_listen:
                mouseListener.stop()
            t = int((time.time() - record_start_time) * 1000)
            btn = 'left' if button == Button.left else 'right'
            try:
                nx = float(x) / float(sw)
                ny = float(y) / float(sh)
            except Exception:
                nx = None
                ny = None
            act = 'DOWN' if pressed else 'UP'
            if nx is not None and ny is not None:
                write_action_line(f"M CLICK {btn} {act} {int(x)} {int(y)} {nx:.6f} {ny:.6f} {t}")
            else:
                write_action_line(f"M CLICK {btn} {act} {int(x)} {int(y)} {t}")

        # scroll mouse
        def on_scroll(x, y, x_axis, y_axis):
            global stop_listen, record_start_time
            if stop_listen:
                mouseListener.stop()
            t = int((time.time() - record_start_time) * 1000)
            write_action_line(f"M SCROLL {int(x_axis)} {int(y_axis)} {t}")

        with mouse.Listener(on_move=on_move, on_click=on_click, on_scroll=on_scroll) as mouseListener:
            mouseListener.join()                


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
        global stop_execute_keyboard
        global pressed_vks
        while True:
            if stop_execute_keyboard:
                return
            try:
                path = action_file_name if os.path.exists(action_file_name) else self.file_name
                with open(path, 'r', encoding='utf-8') as file:
                    keyboard_exec = KeyBoardController()
                    start_ts = time.time()
                    line = file.readline()
                    while line:
                        s = line.strip()
                        if not s or s.startswith('#'):
                            line = file.readline(); continue
                        parts = s.split()
                        if parts[0] == 'K' and len(parts) >= 4:
                            kind = parts[1]
                            try:
                                vk = int(parts[2])
                                t_ms = int(parts[-1])
                            except Exception:
                                line = file.readline(); continue
                            target = start_ts + (t_ms/1000.0)
                            delay = target - time.time()
                            if delay > 0:
                                time.sleep(delay)
                            try:
                                if kind == 'DOWN':
                                    keyboard_exec.press(KeyCode.from_vk(vk))
                                    pressed_vks.add(vk)
                                elif kind == 'UP':
                                    keyboard_exec.release(KeyCode.from_vk(vk))
                                    pressed_vks.discard(vk)
                            except Exception:
                                pass
                        # legacy json support
                        else:
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
            if 'infinite_replay' in globals() and infinite_replay:
                continue
            execute_time_keyboard = execute_time_keyboard - 1
            if execute_time_keyboard <= 0:
                stop_execute_keyboard = True
                return

class MouseActionExecute(threading.Thread):

    def __init__(self, file_name='mouse.action'):
        super().__init__()
        self.daemon = True
        self.file_name = file_name

    def run(self):
        global execute_time_mouse
        global stop_execute_mouse
        while True:
            if stop_execute_mouse:
                return
            try:
                path = action_file_name if os.path.exists(action_file_name) else self.file_name
                with open(path, 'r', encoding='utf-8') as file:
                    mouse_exec = MouseController()
                    # playback-time screen size
                    cw, ch = get_screen_size()
                    rw, rh = cw, ch
                    start_ts = time.time()
                    # pressed buttons tracking for cleanup
                    global pressed_mouse_buttons
                    line = file.readline()
                    while line:
                        s = line.strip()
                        if not s or s.startswith('#'):
                            line = file.readline(); continue
                        parts = s.split()
                        if parts[0] == 'META' and len(parts) >= 4 and parts[1] == 'SCREEN':
                            try:
                                rw = int(parts[2]); rh = int(parts[3])
                            except Exception:
                                rw, rh = cw, ch
                        elif parts[0] == 'M' and len(parts) >= 3:
                            if parts[1] == 'MOVE':
                                # formats: M MOVE x y t  OR  M MOVE x y nx ny t
                                try:
                                    if len(parts) >= 6:
                                        # with norm
                                        x = int(parts[2]); y = int(parts[3]); nx = float(parts[4]); ny = float(parts[5]); t_ms = int(parts[-1])
                                    else:
                                        x = int(parts[2]); y = int(parts[3]); nx = ny = None; t_ms = int(parts[-1])
                                except Exception:
                                    line = file.readline(); continue
                                use_norm = False
                                try:
                                    if rw and rh and (abs(cw - rw) / float(rw) > 0.02 or abs(ch - rh) / float(rh) > 0.02):
                                        use_norm = (nx is not None and ny is not None)
                                except Exception:
                                    use_norm = False
                                if use_norm:
                                    tx = int(round(nx * float(cw)))
                                    ty = int(round(ny * float(ch)))
                                else:
                                    tx = x; ty = y
                                target = start_ts + (t_ms/1000.0)
                                delay = target - time.time()
                                if delay > 0:
                                    time.sleep(delay)
                                mouse_exec.position = (tx, ty)
                            elif parts[1] == 'CLICK':
                                # formats: M CLICK btn DOWN/UP x y [nx ny] t
                                try:
                                    btn = parts[2]; act = parts[3]
                                    x = int(parts[4]); y = int(parts[5])
                                    if len(parts) >= 9:
                                        nx = float(parts[6]); ny = float(parts[7]); t_ms = int(parts[-1])
                                    else:
                                        nx = ny = None; t_ms = int(parts[-1])
                                except Exception:
                                    line = file.readline(); continue
                                use_norm = False
                                try:
                                    if rw and rh and (abs(cw - rw) / float(rw) > 0.02 or abs(ch - rh) / float(rh) > 0.02):
                                        use_norm = (nx is not None and ny is not None)
                                except Exception:
                                    use_norm = False
                                if use_norm:
                                    tx = int(round(nx * float(cw)))
                                    ty = int(round(ny * float(ch)))
                                else:
                                    tx = x; ty = y
                                target = start_ts + (t_ms/1000.0)
                                delay = target - time.time()
                                if delay > 0:
                                    time.sleep(delay)
                                try:
                                    mouse_exec.position = (tx, ty)
                                except Exception:
                                    pass
                                if act == 'DOWN':
                                    if btn == 'left':
                                        mouse_exec.press(Button.left)
                                        pressed_mouse_buttons.add('left')
                                    else:
                                        mouse_exec.press(Button.right)
                                        pressed_mouse_buttons.add('right')
                                else:
                                    if btn == 'left':
                                        mouse_exec.release(Button.left)
                                        pressed_mouse_buttons.discard('left')
                                    else:
                                        mouse_exec.release(Button.right)
                                        pressed_mouse_buttons.discard('right')
                            elif parts[1] == 'SCROLL':
                                try:
                                    dx = int(parts[2]); dy = int(parts[3]); t_ms = int(parts[-1])
                                except Exception:
                                    line = file.readline(); continue
                                target = start_ts + (t_ms/1000.0)
                                delay = target - time.time()
                                if delay > 0:
                                    time.sleep(delay)
                                mouse_exec.scroll(dx, dy)
                        else:
                            # legacy json support
                            try:
                                obj = json.loads(line)
                                if obj.get('name') == 'meta':
                                    rw = int(obj['screen']['w']); rh = int(obj['screen']['h'])
                                elif obj.get('name') == 'mouse':
                                    # legacy path retained (no timing)
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
            if 'infinite_replay' in globals() and infinite_replay:
                continue
            execute_time_mouse = execute_time_mouse - 1
            if execute_time_mouse <= 0:
                stop_execute_mouse = True
                return
                
                
######################################################################
# Controller
######################################################################
class ListenController(threading.Thread):
    
    def __init__(self):
        super().__init__()
        self.daemon = True

    def run(self):
        global stop_listen
        stop_listen = False
        
        def on_release(key):
            global can_start_listening 
            global can_start_executing
            global stop_listen
            
            if key == keyboard.Key.esc:
                stop_listen = True
                can_start_listening = True
                can_start_executing = True
                startListenerBtn['text'] = 'Start recording (F10)'
                startListenerBtn['state'] = 'normal'
                startExecuteBtn['state'] = 'normal'
                keyboardListener.stop()

        with keyboard.Listener(on_release=on_release) as keyboardListener:
            keyboardListener.join()

class ExecuteController(threading.Thread):
    
    def __init__(self):
        super().__init__()
        self.daemon = True

    def run(self):
        global stop_execute_keyboard
        global stop_execute_mouse
        global can_start_listening 
        global can_start_executing

        # Listener to allow ESC to stop replaying
        def on_release(key):
            global stop_execute_keyboard
            global stop_execute_mouse
            if key == keyboard.Key.esc:
                stop_execute_keyboard = True
                stop_execute_mouse = True

        keyboardListener = keyboard.Listener(on_release=on_release)
        keyboardListener.start()

        # Wait until all active workers have finished (or ESC pressed)
        while not (stop_execute_keyboard and stop_execute_mouse):
            time.sleep(0.05)

        # Safety: release any stuck inputs
        release_all_inputs()

        # Reset UI and states once everything is done
        can_start_listening = True
        can_start_executing = True
        startExecuteBtn['text'] = 'Start replaying (F11)'
        startListenerBtn['state'] = 'normal'
        startExecuteBtn['state'] = 'normal'
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
            global root
            # if idle, start recording; else stop via ESC
            if can_start_listening and can_start_executing:
                try:
                    # schedule on Tk main thread to avoid cross-thread UI ops
                    root.after(0, lambda: command_adapter('listen'))
                except Exception:
                    command_adapter('listen')
            else:
                try:
                    kb = KeyBoardController()
                    kb.press(Key.esc)
                    kb.release(Key.esc)
                except Exception:
                    pass

        def toggle_replay():
            global can_start_listening, can_start_executing
            global stop_execute_keyboard, stop_execute_mouse
            global root
            # if idle, start replay; else request stop
            if can_start_listening and can_start_executing:
                try:
                    # schedule on Tk main thread to avoid cross-thread UI ops
                    root.after(0, lambda: command_adapter('execute'))
                except Exception:
                    command_adapter('execute')
            else:
                stop_execute_keyboard = True
                stop_execute_mouse = True
        
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
    # Ensure DPI awareness before creating Tk to avoid window size jumps
    set_process_dpi_aware()

    can_start_listening = True
    can_start_executing = True
    execute_time_keyboard = 0
    execute_time_mouse = 0
    stop_execute_keyboard = True
    stop_execute_mouse = True
    stop_listen = False
    pressed_vks = set()
    pressed_mouse_buttons = set()
    infinite_replay = False
    
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
    root.geometry('720x360')
    root.resizable(0,0)

    # title
    titleLabel = ttk.Label(root, text='Quick Macro', style='BizTitle.TLabel')
    titleLabel.place(x=24, y=14, width=220, height=36)

    # Card style containers
    style = ttk.Style()
    style.configure('Card.TFrame', background='#f8fafc')
    style.configure('CardLabel.TLabel', background='#f8fafc', font=(font_family, 10), foreground='#374151')
    # Ensure checkbutton blends with card background (no visible patch)
    style.configure('Card.TCheckbutton', background='#f8fafc')
    style.map('Card.TCheckbutton', background=[('active', '#f8fafc'), ('!active', '#f8fafc')])
    style.configure('Biz.TButton', anchor='center', font=(font_family, 10), padding=(20, 0))
    # Explicit centered button style with symmetric padding for perfect centering
    style.configure('Center.TButton', anchor='center', font=(font_family, 10), padding=(20, 0))

    recordCard = ttk.Frame(root, style='Card.TFrame', borderwidth=1, relief='solid')
    recordCard.place(x=30, y=70, width=310, height=90)
    replayCard = ttk.Frame(root, style='Card.TFrame', borderwidth=1, relief='solid')
    replayCard.place(x=360, y=70, width=330, height=220)

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

    actionFileVar = tkinter.StringVar()
    files = list_action_files()
    actionFileVar.set(files[-1] if files else '')

    # Action file controls inside the Replay card
    actionFileLabel = ttk.Label(replayCard, text='Action file', style='CardLabel.TLabel')
    actionFileLabel.place(x=15, y=120, width=100, height=26)
    actionFileSelect = ttk.Combobox(replayCard, textvariable=actionFileVar, values=files if files else [], state='readonly', style='Biz.TCombobox')
    actionFileSelect.place(x=120, y=120, width=190, height=28)

    # Refresh button removed; list auto-updates after recording
    
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

        def parse_action_line(s):
            s = s.strip('\n')
            if not s:
                return {'type':'','op':'','vk':'','btn':'','act':'','x':'','y':'','nx':'','ny':'','dx':'','dy':'','ms':'','raw':s}
            if s.startswith('#'):
                return {'type':'#','op':'','vk':'','btn':'','act':'','x':'','y':'','nx':'','ny':'','dx':'','dy':'','ms':'','raw':s}
            parts = s.split()
            try:
                if parts[0] == 'META':
                    return {'type':'META','op':parts[1] if len(parts)>1 else '', 'vk':'','btn':'','act':'','x':(parts[2] if len(parts)>2 else ''), 'y':(parts[3] if len(parts)>3 else ''), 'nx':'','ny':'','dx':'','dy':'','ms':'','raw':s}
                if parts[0] == 'K':
                    return {'type':'VK','op':parts[1], 'vk':(parts[2] if len(parts)>2 else ''), 'btn':'','act':'','x':'','y':'','nx':'','ny':'','dx':'','dy':'','ms':(parts[-1] if len(parts)>3 else ''), 'raw':s}
                if parts[0] == 'M':
                    if parts[1] == 'MOVE':
                        if len(parts) >= 7:
                            # M MOVE x y nx ny ms
                            return {'type':'MS','op':'MOVE','vk':'','btn':'','act':'','x':parts[2],'y':parts[3],'nx':parts[4],'ny':parts[5],'dx':'','dy':'','ms':parts[6],'raw':s}
                        else:
                            # M MOVE x y ms
                            return {'type':'MS','op':'MOVE','vk':'','btn':'','act':'','x':parts[2],'y':parts[3],'nx':'','ny':'','dx':'','dy':'','ms':parts[4] if len(parts)>4 else '','raw':s}
                    if parts[1] == 'CLICK':
                        # M CLICK btn act x y [nx ny] ms
                        btn = parts[2] if len(parts)>2 else ''
                        act = parts[3] if len(parts)>3 else ''
                        if len(parts) >= 10:
                            return {'type':'MS','op':'CLICK','vk':'','btn':btn,'act':act,'x':parts[4],'y':parts[5],'nx':parts[6],'ny':parts[7],'dx':'','dy':'','ms':parts[8] if len(parts)>8 else '','raw':s}
                        else:
                            return {'type':'MS','op':'CLICK','vk':'','btn':btn,'act':act,'x':parts[4] if len(parts)>4 else '','y':parts[5] if len(parts)>5 else '','nx':'','ny':'','dx':'','dy':'','ms':parts[-1] if len(parts)>6 else '','raw':s}
                    if parts[1] == 'SCROLL':
                        return {'type':'MS','op':'SCROLL','vk':'','btn':'','act':'','x':'','y':'','nx':'','ny':'','dx':parts[2] if len(parts)>2 else '','dy':parts[3] if len(parts)>3 else '','ms':parts[4] if len(parts)>4 else '','raw':s}
            except Exception:
                pass
            # Fallback
            return {'type':'','op':'','vk':'','btn':'','act':'','x':'','y':'','nx':'','ny':'','dx':'','dy':'','ms':'','raw':s}

        def compose_action_line(d):
            t = (d.get('type') or '').strip()
            op = (d.get('op') or '').strip()
            if t == '#':
                return d.get('raw','')
            if t == 'META':
                if op.upper() == 'SCREEN':
                    return f"META SCREEN {d.get('x','')} {d.get('y','')}".strip()
                elif op.upper() == 'START':
                    return f"META START {d.get('x','')}".strip()
                return d.get('raw','')
            if t == 'VK':
                vk = d.get('vk','')
                ms = d.get('ms','')
                return f"K {op} {vk} {ms}".strip()
            if t == 'MS':
                if op.upper() == 'MOVE':
                    x,y,nx,ny,ms = d.get('x',''),d.get('y',''),d.get('nx',''),d.get('ny',''),d.get('ms','')
                    if nx and ny:
                        return f"M MOVE {x} {y} {nx} {ny} {ms}".strip()
                    return f"M MOVE {x} {y} {ms}".strip()
                if op.upper() == 'CLICK':
                    btn = d.get('btn',''); act = d.get('act',''); x,y,nx,ny,ms = d.get('x',''),d.get('y',''),d.get('nx',''),d.get('ny',''),d.get('ms','')
                    if nx and ny:
                        return f"M CLICK {btn} {act} {x} {y} {nx} {ny} {ms}".strip()
                    return f"M CLICK {btn} {act} {x} {y} {ms}".strip()
                if op.upper() == 'SCROLL':
                    dx,dy,ms = d.get('dx',''),d.get('dy',''),d.get('ms','')
                    return f"M SCROLL {dx} {dy} {ms}".strip()
            # Fallback
            raw = d.get('raw','')
            return raw if raw else ' '

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
                x, y, w, h = tree.bbox(target)
                above = event.y < (y + h/2)
                idx = tree.index(target)
                insert_at = idx if above else idx + 1
                line_y = y if above else y + h
                return insert_at, line_y
            # outside rows
            first_bbox = tree.bbox(children[0])
            last_bbox = tree.bbox(children[-1])
            if event.y < first_bbox[1]:
                return 0, first_bbox[1]
            else:
                return len(children), last_bbox[1] + last_bbox[3]

        def on_tree_motion(event):
            nonlocal drag_insert_index, dragging_selection
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
            insert_at, line_y = compute_insert_at(event)
            drag_insert_index = insert_at
            try:
                insert_line.place(in_=tree, x=0, y=line_y, width=tree.winfo_width(), height=2)
            except Exception:
                pass

        def on_tree_release(event):
            nonlocal dragging_selection, drag_insert_index
            try:
                insert_line.place_forget()
            except Exception:
                pass
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
    editBtn.place(x=120, y=155, width=80, height=28)
    openBtn = ttk.Button(replayCard, text='Folder', command=open_actions_folder, style='Biz.TButton')
    openBtn.place(x=15, y=155, width=100, height=28)
    
    # Start hotkeys listener (F10/F11)
    HotkeyController().start()

    # Removed Tk window key binds to avoid double-trigger; global hotkeys handle F10/F11

    # Ensure closing window terminates app
    def on_close():
        global stop_execute_keyboard, stop_execute_mouse, stop_listen
        stop_execute_keyboard = True
        stop_execute_mouse = True
        stop_listen = True
        release_all_inputs()
        try:
            root.destroy()
        except Exception:
            pass

    root.protocol('WM_DELETE_WINDOW', on_close)

    # run
    root.mainloop()
    

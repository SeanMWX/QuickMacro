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
    

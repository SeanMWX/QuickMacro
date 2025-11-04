import json
import ctypes
import threading
import time
import tkinter

from pynput import keyboard, mouse
from pynput.keyboard import Controller as KeyBoardController, KeyCode
from pynput.mouse import Button, Controller as MouseController

######################################################################
# Helpers
######################################################################
def get_screen_size():
    try:
        user32 = ctypes.windll.user32
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass
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
    
    if can_start_listening or can_start_executing:
        
        if action == 'listen':
            custom_thread_list.append(
                {
                    'obj_thread': ListenController(),
                    'obj_ui': startListenerBtn,
                    'action': 'listen'
                }
            )
            if isRecordMouse.get():
                custom_thread_list.append(
                    {
                        'obj_thread': MouseActionListener(),
                        'obj_ui': None
                    }
                )
            if isRecordKeyboard.get():
                custom_thread_list.append(
                    {
                        'obj_thread': KeyboardActionListener(),
                        'obj_ui': None
                    }
                )
            count_down = listenCountDown.get()

        elif action == 'execute':
            custom_thread_list.append(
                {
                    'obj_thread': ExecuteController(),
                    'obj_ui': startExecuteBtn,
                    'action': 'execute'
                }
            )
            if isReplayMouse.get():
                custom_thread_list.append(
                    {
                        'obj_thread': MouseActionExecute(),
                        'obj_ui': None
                    }
                )
            if isReplayKeyboard.get():
                custom_thread_list.append(
                    {
                        'obj_thread': KeyboardActionExecute(),
                        'obj_ui': None
                    }
                )
            count_down = executeCountDown.get()
            execute_time_keyboard = playCount.get()
            execute_time_mouse = playCount.get()
            # initialize stop flags based on which workers are active
            stop_execute_keyboard = not isReplayKeyboard.get()
            stop_execute_mouse = not isReplayMouse.get()

        can_start_listening = False
        can_start_executing = False
        UIUpdateCutDownExecute(count_down, custom_thread_list).start()


######################################################################
# Update UI
######################################################################        
class UIUpdateCutDownExecute(threading.Thread):
    def __init__(self, count_down, custom_thread_list):
        super().__init__()
        self.count_down = count_down
        self.custom_thread_list = custom_thread_list

    def run(self):
        while self.count_down > 0:
            startListenerBtn['state'] = 'disabled'
            startExecuteBtn['state'] = 'disabled'
            
            for custom_thread in self.custom_thread_list:
                if custom_thread['obj_ui'] is not None:
                    custom_thread['obj_ui']['text'] = str(self.count_down)
                    self.count_down = self.count_down - 1
            time.sleep(1)
        else:
            for custom_thread in self.custom_thread_list:
                if custom_thread['obj_ui'] is not None:
                    if custom_thread['action'] == 'listen':
                        custom_thread['obj_ui']['text'] = str('Recording, "ESC" to stop.')
                    elif custom_thread['action'] == 'execute':
                        custom_thread['obj_ui']['text'] = str('Replaying, "ESC" to stop.')
                if custom_thread['obj_thread'] is not None:
                    custom_thread['obj_thread'].start()


######################################################################
# Listen
###################################################################### 
class KeyboardActionListener(threading.Thread):
    
    def __init__(self, file_name='keyboard.action'):
        super().__init__()
        self.file_name = file_name

    def run(self):
        with open(self.file_name, 'w', encoding='utf-8') as file:
            # press keyboard
            def on_press(key):
                template = keyboard_action_template()
                template['event'] = 'press'
                try:
                    template['vk'] = key.vk
                except AttributeError:
                    template['vk'] = key.value.vk
                finally:
                    file.writelines(json.dumps(template) + "\n")
                    file.flush()

            # release keyboard
            def on_release(key):
                global can_start_listening
                global can_start_executing
                global stop_listen

                if key == keyboard.Key.esc:
                    # Stop by pressing "ESC"
                    stop_listen= True
                    keyboardListener.stop()
                    return False

                if not stop_listen:
                    template = keyboard_action_template()
                    template['event'] = 'release'
                    try:
                        template['vk'] = key.vk
                    except AttributeError:
                        template['vk'] = key.value.vk
                    finally:
                        file.writelines(json.dumps(template) + "\n")
                        file.flush()
            
            with keyboard.Listener(on_press=on_press, on_release=on_release) as keyboardListener:
                keyboardListener.join()          
                

class MouseActionListener(threading.Thread):

    def __init__(self, file_name='mouse.action'):
        super().__init__()
        self.file_name = file_name

    def run(self):
        with open(self.file_name, 'w', encoding='utf-8') as file:
            # record-time screen size (DPI aware)
            sw, sh = get_screen_size()
            file.writelines(json.dumps({"name": "meta", "screen": {"w": sw, "h": sh}}) + "\n")
            # move mouse
            def on_move(x, y):
                global stop_listen
                if stop_listen:
                    mouseListener.stop()
                template = mouse_action_template()
                template['event'] = 'move'
                template['location']['x'] = int(x)
                template['location']['y'] = int(y)
                # normalized for DPI/resolution independence
                try:
                    template['location']['nx'] = float(x) / float(sw)
                    template['location']['ny'] = float(y) / float(sh)
                except Exception:
                    pass
                file.writelines(json.dumps(template) + "\n")
                file.flush()

            # click mouse
            def on_click(x, y, button, pressed):
                global stop_listen
                if stop_listen:
                    mouseListener.stop()
                template = mouse_action_template()
                template['event'] = 'click'
                template['target'] = button.name
                template['action'] = pressed
                template['location']['x'] = int(x)
                template['location']['y'] = int(y)
                try:
                    template['location']['nx'] = float(x) / float(sw)
                    template['location']['ny'] = float(y) / float(sh)
                except Exception:
                    pass
                file.writelines(json.dumps(template) + "\n")
                file.flush()

            # scroll mouse
            def on_scroll(x, y, x_axis, y_axis):
                global stop_listen
                if stop_listen:
                    mouseListener.stop()
                template = mouse_action_template()
                template['event'] = 'scroll'
                template['location']['x'] = x_axis
                template['location']['y'] = y_axis
                file.writelines(json.dumps(template) + "\n")
                file.flush()

            with mouse.Listener(on_move=on_move, on_click=on_click, on_scroll=on_scroll) as mouseListener:
                mouseListener.join()                


######################################################################
# Executing
######################################################################
class KeyboardActionExecute(threading.Thread):

    def __init__(self, file_name='keyboard.action'):
        super().__init__()
        self.file_name = file_name

    def run(self):
        global execute_time_keyboard
        global stop_execute_keyboard
        while execute_time_keyboard >= 0:
            if stop_execute_keyboard:
                return
            
            with open(self.file_name, 'r', encoding='utf-8') as file:
                keyboard_exec = KeyBoardController()
                line = file.readline()
                while line:
                    obj = json.loads(line)
                    if obj['name'] == 'keyboard':
                        if obj['event'] == 'press':
                            keyboard_exec.press(KeyCode.from_vk(obj['vk']))
                            time.sleep(0.01)
                        elif obj['event'] == 'release':
                            keyboard_exec.release(KeyCode.from_vk(obj['vk']))
                            time.sleep(0.01)
                    line = file.readline()
            execute_time_keyboard = execute_time_keyboard - 1
            if execute_time_keyboard == 0:
                stop_execute_keyboard = True

class MouseActionExecute(threading.Thread):

    def __init__(self, file_name='mouse.action'):
        super().__init__()
        self.file_name = file_name

    def run(self):
        global execute_time_mouse
        global stop_execute_mouse
        while execute_time_mouse >= 0:
            if stop_execute_mouse:
                return
            
            with open(self.file_name, 'r', encoding='utf-8') as file:
                mouse_exec = MouseController()
                # playback-time screen size
                cw, ch = get_screen_size()
                rw, rh = cw, ch
                line = file.readline()
                while line:
                    obj = json.loads(line)
                    if obj.get('name') == 'meta':
                        try:
                            rw = int(obj['screen']['w'])
                            rh = int(obj['screen']['h'])
                        except Exception:
                            rw, rh = cw, ch
                    elif obj['name'] == 'mouse':
                        if obj['event'] == 'move':
                            lx = obj['location'].get('x')
                            ly = obj['location'].get('y')
                            nx = obj['location'].get('nx')
                            ny = obj['location'].get('ny')
                            if nx is not None and ny is not None and rw and rh:
                                tx = int(round(float(nx) * float(cw)))
                                ty = int(round(float(ny) * float(ch)))
                            else:
                                tx = int(lx)
                                ty = int(ly)
                            mouse_exec.position = (tx, ty)
                            time.sleep(0.01)
                        elif obj['event'] == 'click':
                            # ensure we click at recorded coordinates
                            lx = obj['location'].get('x')
                            ly = obj['location'].get('y')
                            nx = obj['location'].get('nx')
                            ny = obj['location'].get('ny')
                            if nx is not None and ny is not None and rw and rh:
                                tx = int(round(float(nx) * float(cw)))
                                ty = int(round(float(ny) * float(ch)))
                            else:
                                tx = int(lx)
                                ty = int(ly)
                            try:
                                mouse_exec.position = (tx, ty)
                            except Exception:
                                pass

                            if obj['action']:
                                if obj['target'] == 'left':
                                    mouse_exec.press(Button.left)
                                else:
                                    mouse_exec.press(Button.right)
                            else:
                                if obj['target'] == 'left':
                                    mouse_exec.release(Button.left)
                                else:
                                    mouse_exec.release(Button.right)
                            time.sleep(0.01)
                        elif obj['event'] == 'scroll':
                            mouse_exec.scroll(obj['location']['x'], obj['location']['y'])
                            time.sleep(0.01)
                    line = file.readline()
            execute_time_mouse = execute_time_mouse - 1
            if execute_time_mouse == 0:
                stop_execute_mouse = True                
                
                
######################################################################
# Controller
######################################################################
class ListenController(threading.Thread):
    
    def __init__(self):
        super().__init__()

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
                startListenerBtn['text'] = 'Start recording'
                startListenerBtn['state'] = 'normal'
                startExecuteBtn['state'] = 'normal'
                keyboardListener.stop()

        with keyboard.Listener(on_release=on_release) as keyboardListener:
            keyboardListener.join()

class ExecuteController(threading.Thread):
    
    def __init__(self):
        super().__init__()

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

        # Reset UI and states once everything is done
        can_start_listening = True
        can_start_executing = True
        startExecuteBtn['text'] = 'Start replaying'
        startListenerBtn['state'] = 'normal'
        startExecuteBtn['state'] = 'normal'
        keyboardListener.stop()

            
######################################################################
# GUI
######################################################################
if __name__ == '__main__':
    
    can_start_listening = True
    can_start_executing = True
    execute_time_keyboard = 0
    execute_time_mouse = 0
    stop_execute_keyboard = True
    stop_execute_mouse = True
    stop_listen = False
    
    root = tkinter.Tk()
    root.title('Quick Macro - Sean Zou')
    root.geometry('400x270')
    root.resizable(0,0)

    # recording
    # time to record
    listenerStartLabel = tkinter.Label(root, text='Record countdown')
    listenerStartLabel.place(x=100, y=10, width=120, height=20)
    
    listenCountDown = tkinter.IntVar()
    listenCountDown.set(3)
    
    listenerStartEdit = tkinter.Entry(root, textvariable=listenCountDown)
    listenerStartEdit.place(x=220, y=10, width=60, height=20)
    
    listenerTipLabel = tkinter.Label(root, text='s')
    listenerTipLabel.place(x=280, y=10, width=20, height=20)

    # start recording
    startListenerBtn = tkinter.Button(root, text="Start recording", command=lambda: command_adapter('listen'))
    startListenerBtn.place(x=100, y=45, width=200, height=30)

    # replaying
    # time to replay
    executeEndLabel = tkinter.Label(root, text='Replay countdown')
    executeEndLabel.place(x=100, y=85, width=120, height=20)
    
    executeCountDown = tkinter.IntVar()
    executeCountDown.set(3)
    
    executeEndEdit = tkinter.Entry(root, textvariable=executeCountDown)
    executeEndEdit.place(x=220, y=85, width=60, height=20)
    
    
    executeTipLabel = tkinter.Label(root, text='s')
    executeTipLabel.place(x=280, y=85, width=20, height=20)

    # times for replaying
    playCountLabel = tkinter.Label(root, text='Repeat Times')
    playCountLabel.place(x=100, y=115, width=120, height=20)
    
    playCount = tkinter.IntVar()
    playCount.set(1)
    
    playCountEdit = tkinter.Entry(root, textvariable=playCount)
    playCountEdit.place(x=220, y=115, width=60, height=20)

    playCountTipLabel = tkinter.Label(root, text='#')
    playCountTipLabel.place(x=280, y=115, width=20, height=20)

    # start replaying
    startExecuteBtn = tkinter.Button(root, text="Start replaying", command=lambda: command_adapter('execute'))
    startExecuteBtn.place(x=100, y=145, width=200, height=30)
    
    # if record mouse
    isRecordMouse = tkinter.BooleanVar()
    isRecordMouse.set(False)
    
    recordMouseCheckbox = tkinter.Checkbutton(root, text='record mouse', variable=isRecordMouse)
    recordMouseCheckbox.place(x=80, y=200, width=120, height=20)
    
    # if record keyboard
    isRecordKeyboard = tkinter.BooleanVar()
    isRecordKeyboard.set(True)
    
    recordKeyboardCheckbox = tkinter.Checkbutton(root, text='record keyborad', variable=isRecordKeyboard)
    recordKeyboardCheckbox.place(x=200, y=200, width=120, height=20)
    
    # if replay mouse
    isReplayMouse = tkinter.BooleanVar()
    isReplayMouse.set(False)
    
    replayMouseCheckbox = tkinter.Checkbutton(root, text='replay mouse', variable=isReplayMouse)
    replayMouseCheckbox.place(x=80, y=225, width=120, height=20)
    
    # if replay keyboard
    isReplayKeyboard = tkinter.BooleanVar()
    isReplayKeyboard.set(True)
    
    replayKeyboardCheckbox = tkinter.Checkbutton(root, text='replay keyboard', variable=isReplayKeyboard)
    replayKeyboardCheckbox.place(x=200, y=225, width=120, height=20)
    
    # run
    root.mainloop()
    

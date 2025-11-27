from core.recorder import Recorder
from typing import Callable

class RecordingController:
    def __init__(self, state, init_new_action_file: Callable, select_current_action_in_dropdown: Callable, update_ui_for_state: Callable, logger=None, on_started=None, on_stopped=None):
        self.state = state
        self.init_new_action_file = init_new_action_file
        self.select_current_action_in_dropdown = select_current_action_in_dropdown
        self.update_ui_for_state = update_ui_for_state
        self.logger = logger
        self.on_started = on_started
        self.on_stopped = on_stopped

    def start(self):
        if not (self.state.can_start_listening and self.state.can_start_executing):
            return
        self.init_new_action_file()
        self.select_current_action_in_dropdown()
        try:
            self.state.ev_stop_listen.clear()
        except Exception:
            pass
        self.update_ui_for_state('recording')
        try:
            self.state.current_recorder = Recorder(self.state.action_file_name, self.state.ev_stop_listen)
            self.state.current_recorder.start()
        except Exception:
            pass
        self.state.can_start_listening = False
        self.state.can_start_executing = False
        if callable(self.logger):
            self.logger(f"Start recording -> {self.state.action_file_name}")
        if callable(self.on_started):
            self.on_started()

    def stop(self):
        try:
            self.state.ev_stop_listen.set()
        except Exception:
            pass
        self.state.can_start_listening = True
        self.state.can_start_executing = True
        self.update_ui_for_state('idle')
        if callable(self.logger):
            self.logger("Recording stopped")
        if callable(self.on_stopped):
            self.on_stopped()

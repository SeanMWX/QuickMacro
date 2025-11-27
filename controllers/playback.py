from typing import Callable
from core.replayer import Replayer

class PlaybackController:
    def __init__(self, state, update_ui_for_state: Callable, release_all_inputs: Callable, start_monitor: Callable, on_progress: Callable, on_loop_start: Callable, logger=None, on_started=None, on_stopped=None):
        self.state = state
        self.update_ui_for_state = update_ui_for_state
        self.release_all_inputs = release_all_inputs
        self.start_monitor = start_monitor
        self.on_progress = on_progress
        self.on_loop_start = on_loop_start
        self.logger = logger
        self.on_started = on_started
        self.on_stopped = on_stopped

    def start(self, path, repeat_count, infinite=False, use_rel=False, rel_gain=1.0, rel_auto=True, total_ms=0):
        if not (self.state.can_start_listening and self.state.can_start_executing):
            if callable(self.logger):
                self.logger("Skip start replay (busy)")
            return
        self.state.action_file_name = path
        try:
            if infinite:
                self.state.ev_infinite_replay.set()
            else:
                self.state.ev_infinite_replay.clear()
        except Exception:
            pass
        self.state.execute_time_keyboard = repeat_count
        self.state.execute_time_mouse = repeat_count
        try:
            self.state.ev_stop_execute_keyboard.clear()
            self.state.ev_stop_execute_mouse.clear()
        except Exception:
            pass
        try:
            self.release_all_inputs()
        except Exception:
            pass
        try:
            self.start_monitor(repeat_count, total_ms/1000.0)
        except Exception:
            pass
        self.update_ui_for_state('replaying')
        try:
            self.state.current_replayer = Replayer(path, self.state.ev_stop_execute_keyboard, self.state.ev_stop_execute_mouse, self.state.ev_infinite_replay, repeat_count, use_rel, rel_gain, rel_auto, self.on_progress, self.on_loop_start)
            self.state.current_replayer.start()
        except Exception:
            pass
        if callable(self.logger):
            self.logger(f"Replay started -> {path}")
        if callable(self.on_started):
            self.on_started()

    def stop(self):
        try:
            self.state.ev_stop_execute_keyboard.set()
            self.state.ev_stop_execute_mouse.set()
        except Exception:
            pass
        self.state.can_start_listening = True
        self.state.can_start_executing = True
        self.update_ui_for_state('idle')
        if callable(self.logger):
            self.logger("Replay stop requested")
        if callable(self.on_stopped):
            self.on_stopped()

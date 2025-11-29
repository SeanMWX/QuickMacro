import threading
import time
import os
from pynput import keyboard


def _safe_set(ev):
    try:
        ev.set()
    except Exception:
        pass


class ExecuteController(threading.Thread):
    def __init__(self, state, ui_refs=None, command_adapter=None, release_all_inputs=None, on_finished=None):
        super().__init__()
        self.daemon = True
        self.state = state
        self.ui_refs = ui_refs
        self.command_adapter = command_adapter
        self.release_all_inputs = release_all_inputs
        self.on_finished = on_finished

    def run(self):
        keyboardListener = keyboard.Listener(on_release=lambda key: None)
        keyboardListener.start()

        while not (self.state.ev_stop_execute_keyboard.is_set() and self.state.ev_stop_execute_mouse.is_set()):
            time.sleep(0.05)

        try:
            if callable(self.release_all_inputs):
                self.release_all_inputs()
        except Exception:
            pass
        try:
            self.ui_refs.log_event("Replay stopped")
        except Exception:
            pass

        self.state.can_start_listening = True
        self.state.can_start_executing = True
        try:
            self.ui_refs.update_ui_for_state('idle')
        except Exception:
            pass

        try:
            base_name = os.path.basename(self.state.action_file_name) if self.state.action_file_name else ''
            if not self.state.restarting_flag:
                self.ui_refs.mark_finished()
            elif self.state.restarting_flag and self.state.pending_main_action:
                if base_name.lower() != 'restart.action' and not self.state.restart_running:
                    def _start_restart():
                        try:
                            self.state.restart_running = True
                            self.state.action_file_name = os.path.join('actions', 'restart.action')
                            self.ui_refs.actionFileVar.set('restart.action')
                            try:
                                # reset monitor timers for restart run
                                self.state.monitor_loop_start_ts = time.monotonic()
                                self.state.dungeon_start_ts = time.monotonic()
                                self.ui_refs.update_monitor_labels()
                            except Exception:
                                pass
                            self.state.ev_stop_execute_keyboard.clear()
                            self.state.ev_stop_execute_mouse.clear()
                            self.state.skip_run_increment = True
                            if callable(self.command_adapter):
                                self.command_adapter('execute')
                            self.state.skip_run_increment = False
                            # schedule force switch back after 13s regardless of restart completion
                            try:
                                if self.state.restart_back_job:
                                    self.ui_refs.root.after_cancel(self.state.restart_back_job)
                                def _force_resume():
                                    if self.state.restart_running:
                                        _safe_set(self.state.ev_stop_execute_keyboard)
                                        _safe_set(self.state.ev_stop_execute_mouse)
                                        # mark as if restart finished to trigger resume_main path
                                        self.state.restarting_flag = True
                                self.state.restart_back_job = self.ui_refs.root.after(13000, _force_resume)
                            except Exception:
                                pass
                        except Exception:
                            pass
                    self.ui_refs.root.after(0, _start_restart)
                elif base_name.lower() == 'restart.action' and self.state.restart_running:
                    def _resume_main():
                        try:
                            main_path = self.state.pending_main_action
                            if main_path and (not os.path.exists(main_path)):
                                cand = os.path.join('actions', os.path.basename(main_path))
                                if os.path.exists(cand):
                                    main_path = cand
                            self.state.action_file_name = main_path or ''
                            if main_path:
                                self.ui_refs.actionFileVar.set(os.path.basename(main_path))
                            if self.state.pending_main_playcount not in (None, ''):
                                self.ui_refs.playCount.set(int(self.state.pending_main_playcount))
                            self.state.pending_main_action = None
                            self.state.pending_main_playcount = None
                            self.state.restarting_flag = False
                            self.state.restart_running = False
                            try:
                                if self.state.restart_back_job:
                                    self.ui_refs.root.after_cancel(self.state.restart_back_job)
                            except Exception:
                                pass
                            try:
                                # reset monitor timers for resumed main run
                                self.state.monitor_loop_start_ts = time.monotonic()
                                self.state.dungeon_start_ts = None
                                self.ui_refs.update_monitor_labels()
                            except Exception:
                                pass
                            self.state.ev_stop_execute_keyboard.clear()
                            self.state.ev_stop_execute_mouse.clear()
                            self.state.skip_run_increment = True
                            if callable(self.command_adapter):
                                self.command_adapter('execute')
                            self.state.skip_run_increment = False
                        except Exception:
                            pass
                    self.ui_refs.root.after(0, _resume_main)
        except Exception:
            pass
        try:
            if callable(self.on_finished):
                self.on_finished()
        except Exception:
            pass
        keyboardListener.stop()

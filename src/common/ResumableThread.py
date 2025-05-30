import threading
from abc import abstractmethod


class ResumableThread(threading.Thread):
    def __init__(self, target=None, group=None, name=None,
                 args=(), kwargs=None, *, daemon=None):
        if target is None:
            target = self.perform

        super().__init__(group=group, target=target, name=name,
                         args=args, kwargs=kwargs, daemon=daemon)
        self.paused = False
        self.terminated = False
        self.daemon = True
        self.pause_condition = threading.Condition(threading.Lock())
        self.completed_event = threading.Event()
        self.is_running: bool = False

    @abstractmethod
    def perform(self):
        pass

    def run(self):
        self.is_running = True
        self._target(*self._args, **self._kwargs)

    def pause(self):
        self.paused = True

    def resume(self):
        with self.pause_condition:
            self.paused = False
            self.pause_condition.notify()

    def is_completed(self):
        return not self.is_running and self.completed_event.is_set()

    def is_running_code(self):
        return self.is_running and not self.completed_event.is_set()

    def mark_completed(self):
        self.is_running = False
        self.completed_event.set()

    def terminate(self):
        self.terminated = True
        with self.pause_condition:
            self.paused = False
            self.pause_condition.notify()

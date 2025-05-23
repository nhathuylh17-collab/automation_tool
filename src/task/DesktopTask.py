import time
from abc import ABC
from logging import Logger
from typing import Callable, Any

import psutil
import pyautogui
import pygetwindow as gw
from pygetwindow import Win32Window
from pywinauto import Application, WindowSpecification

from src.common.Stack import Stack
from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class DesktopTask(AutomatedTask, ABC):

    def __init__(self,
                 settings: dict[str, str],
                 callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self._window_title_stack: Stack[str] = Stack[str]()
        self._app: Application = None
        self._window: WindowSpecification = None

    def _kill_processes(self, process_name: str):
        """
        Kill all processes with the given name.
        Returns True if at least one process was killed, False otherwise.
        """
        pids: list[int] = self._get_matching_processes(process_name)
        if len(pids) == 0:
            print(f"No processes found with name '{process_name}'.")
            return False

        process: psutil.Process = None
        for pid in pids:
            try:
                process = psutil.Process(pid)
                # process.terminate()  # Send SIGTERM (graceful termination)
                # print(f"Terminated process '{process_name}' with PID: {pid}")
                # # Optional: Wait briefly to ensure termination
                # process.wait(timeout=3)
                process.kill()
            except psutil.NoSuchProcess:
                print(f"Process with PID {pid} no longer exists.")
            except psutil.AccessDenied:
                print(f"Access denied to terminate process with PID {pid}. Try running as administrator/root.")
            except psutil.TimeoutExpired:
                print(f"Process with PID {pid} did not terminate in time. Forcing kill...")
                process.kill()  # Send SIGKILL (forceful termination)
                print(f"Forced kill of process with PID {pid}.")
            except Exception as e:
                print(f"Error terminating process with PID {pid}: {e}")

        return True

    def _get_matching_processes(self, process_name: str) -> list[int]:
        """
        Check if a process with the given name is running.
        Returns True if found, False otherwise.
        """
        pids: list[int] = []
        for process in psutil.process_iter(['name']):

            try:
                # Compare process name (case-insensitive)
                if process.info['name'].lower().startswith(process_name.lower()):
                    pids.append(process.pid)

            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                # Skip processes that can't be accessed or no longer exist
                continue

        return pids

    def _is_current_window_having_title(self, expected_title: str) -> bool:
        window: Win32Window = gw.getActiveWindow()

        if str(window.title).__contains__(expected_title):
            return True

        return False

    def _wait_for_window(self, title: str, max_attempt: int = 20):
        current_attempt: int = 0

        while current_attempt < max_attempt:

            window_titles: list[str] = gw.getAllTitles()

            for window_title in window_titles:
                if window_title.__contains__(title):
                    gw.getWindowsWithTitle(window_title)[0].activate()
                    return window_title

            current_attempt += 1
            time.sleep(1)
        raise Exception('Can not find out the asked window {}'.format(title))

    def _hotkey_then_close_current_window(self, *args: Any) -> WindowSpecification:
        self._window_title_stack.pop()
        current_window_title: str = self._window_title_stack.peek()
        return self.__hotkey_then_activate_window_by_title(current_window_title, args)

    def _hotkey_then_open_new_window(self, window_title: str, *args: Any) -> WindowSpecification:
        if window_title is None or len(window_title) == 0:
            raise Exception('Invalid window title')

        self._window_title_stack.append(window_title)
        new_window_title: str = self._window_title_stack.peek()
        return self.__hotkey_then_activate_window_by_title(new_window_title, args)

    def __hotkey_then_activate_window_by_title(self, window_title, args) -> WindowSpecification:
        counter: int = 0
        while not self._is_current_window_having_title(window_title):

            if counter > 10:
                raise Exception(
                    f"Time is over for waiting {window_title} appear by the hotkey combination {args}")

            pyautogui.hotkey(*args)
            self.sleep()
            counter += 1

        gw.getWindowsWithTitle(window_title)[0].activate()
        self._window = self._app.window(title=self._window_title_stack.peek())
        return self._window

    def _close_windows_util_reach_first_gscc(self):
        return self._close_windows_with_window_title_stack(self._window_title_stack)

    def _close_windows_with_window_title_stack(self, window_title_stack: list[str]):
        logger: Logger = get_current_logger()

        for i in range(window_title_stack.__len__() - 1, 0, -1):
            window_title = window_title_stack[i]

            try:
                app = Application().connect(title=window_title, timeout=1)

                window = app.window(title=window_title)

                if window.exists(timeout=1):
                    window.close()
                    logger.debug(f"Window with title '{window_title}' has been closed.")
                else:
                    logger.debug(f"Window with title '{window_title}' does not exist.")

                window_title_stack.pop()

            except BaseException as e:
                logger.debug(f"An error occurred: {e}")
                i += 1

        # self._wait_for_window(window_title_stack[0])

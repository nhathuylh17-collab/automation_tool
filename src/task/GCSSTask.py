import subprocess
from abc import abstractmethod
from logging import Logger

import pyautogui
from pywinauto import Application, WindowSpecification
from pywinauto.controls.win32_controls import ComboBoxWrapper

from src.common.ProcessUtil import get_matching_processes, kill_processes
from src.common.ThreadLocalLogger import get_current_logger
from src.task.DesktopTask import DesktopTask


class GCSSTask(DesktopTask):

    def mandatory_settings(self) -> list[str]:
        return ['gcss_profile_name']

    def automate(self):
        try:
            self._pre_actions()
            self.automate_gcss()
            self._post_actions()
        finally:
            kill_processes('GCSS')

    @abstractmethod
    def automate_gcss(self):
        pass

    def _pre_actions(self):
        # check the current windows whether it contains the GCSS already
        # only open when the result is No
        # If yes then go back to the initial stage/page of the GCSS
        if len(get_matching_processes('GCSS')) > 0:
            kill_processes('GCSS')

        self._open_the_gcss()

        self._wait_for_window('Select User Profile')
        self._window_title_stack.append('Select User Profile')

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        ComboBox_maintain_payment: ComboBoxWrapper = self._window.children(class_name="ComboBox")[0]
        ComboBox_maintain_payment.select(self._settings['gcss_profile_name'])
        pyautogui.hotkey('tab')
        self.sleep()
        pyautogui.hotkey('enter')
        self._window_title_stack.clear()

    def _post_actions(self):
        if len(get_matching_processes('GCSS')) > 0:
            kill_processes('GCSS')

    def _open_the_gcss(self):
        logger: Logger = get_current_logger()

        exe_path = r"C:\Program Files (x86)\GCSS\PROD_A\GCSSExport.exe"
        argument = r"-wsnaddr =//gcssexport1.gls.dk.eur.crb.apmoller.net:15000"

        try:
            # Run the command asynchronously with Popen
            process = subprocess.Popen([exe_path, argument],
                                       stdout=subprocess.PIPE,
                                       stderr=subprocess.PIPE,
                                       text=True)
            logger.info(f"Command started with PID {process.pid}")
            # Continue without waiting for the process to finish
        except FileNotFoundError:
            logger.error("Error: The executable file was not found. Please check the path.")
        except Exception as e:
            logger.error(f"An unexpected error occurred: {e}")

import time
from datetime import datetime
from logging import Logger
from typing import Callable

import pyautogui
import pygetwindow as gw
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ProcessUtil import get_matching_processes
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.GCSSTask import GCSSTask


class GCSS_SPIR(GCSSTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.current_status_excel_col_index = None
        self.excel_provider: ExcelReaderProvider = None
        self.current_worksheet = None
        # self.current_status_excel_col_index: int = 0
        self.current_status_excel_row_index: int = 2

    def mandatory_settings(self) -> list[str]:
        desktop_mandatory_keys: list[str] = super().mandatory_settings()
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'excel.shipment']
        mandatory_keys.extend(desktop_mandatory_keys)
        return mandatory_keys

    def automate_gcss(self):
        logger: Logger = get_current_logger()

        self.excel_provider: ExcelReaderProvider = XlwingProvider()
        path_to_excel = self._settings['excel.path']
        workbook = self.excel_provider.get_workbook(path=path_to_excel)

        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        self.current_worksheet = self.excel_provider.get_worksheet(workbook, sheet_name)

        shipments: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                     self._settings['excel.sheet'],
                                                                     self._settings['excel.shipment'])

        self.current_element_count = 0
        self.total_element_size = len(shipments)

        self._wait_for_window('Pending Tray')
        self._window_title_stack.append('Pending Tray')

        for i, shipment in enumerate(shipments):
            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    logger.info("Currently pause")
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            print('START_FLOW current gcss process count {}'.format(len(get_matching_processes('GCSS'))))
            if len(get_matching_processes('GCSS')) == 0:
                self._pre_actions()

            self._wait_for_window('Pending Tray')

            # return to code
            logger.info("Start process shipment " + shipment)

            try:

                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")

                pyautogui.hotkey('ctrl', 'o')
                pyautogui.typewrite(shipment)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')

                # handle for adhoc shipment
                try:
                    self._wait_for_window(shipment)
                except:
                    self.handle_invalid_window(shipment, workbook)

                    self._wait_for_window(shipment)

                    status_column_B, status_column_C = self.process_on_each_shipment_adhoc(shipment)

                    self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                        2, status_column_B)
                    self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                        3, status_column_C)
                    self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                        4,
                                                        current_timestamp)

                    self.excel_provider.save(workbook)

                    logger.info("Done with shipment " + shipment)

                    self.current_status_excel_row_index += 1
                    self.current_element_count += 1
                    continue

                status_column_B, status_column_C = self.process_on_each_shipment(shipment)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2, status_column_B)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3, status_column_C)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)

                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)

                logger.info("Done with shipment " + shipment)

                print(
                    'END_FLOW current gcss process count {}'.format(len(get_matching_processes('GCSS'))))
                time.sleep(3)

            except SkipToNextShipment as e:
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")

                status_column_B = e.status_b
                status_column_C = e.status_c

                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2, status_column_B)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3, status_column_C)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1

                self.excel_provider.save(workbook)
                continue

            except Skipnoactivity as e:
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")

                status_column_B = e.status_b
                status_column_C = e.status_c

                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2, status_column_B)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3, status_column_C)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1

                self.excel_provider.save(workbook)
                continue

            except BaseException as e:
                logger.info(f'Cannot handle shipment {shipment}. \n {e} \nMoving to next shipment')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")

                if len(get_matching_processes('GCSS')) == 0:
                    self._pre_actions()
                else:
                    self._close_windows_util_reach_first_gscc()

                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Error - Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Cannot handle shipment {}, please check manual'.format(shipment))
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.excel_provider.save(workbook)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                continue

        try:
            self.excel_provider.save(workbook)
        except BaseException as e:
            logger.debug("Cannot save excel file")
        finally:
            self.excel_provider.close(workbook)
            self.excel_provider.quit_session()

    def process_on_each_shipment(self, shipment):
        logger: Logger = get_current_logger()
        window_normal_shipment: str = self._wait_for_window(shipment)
        self._window_title_stack.append(window_normal_shipment)
        gw.getWindowsWithTitle(window_normal_shipment)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        while True:
            pyautogui.hotkey('ctrl', 'k')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            if list_views.__len__() == 2:
                break
            self.sleep()

        list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")[1]
        items = list_views.items()

        if not items:
            logger.info('Not found any activity in {}, skip to next shipment'.format(shipment))
            self._close_windows_util_reach_first_gscc()
            raise Skipnoactivity(status_b="Skip", status_c="No activity found")

        runner = 0
        array = [None for _ in range(8)]
        for item in list_views.items():
            array[runner] = item.text()

            if runner != 7:
                runner = runner + 1
                continue

            runner = 0

            if 'LOAD' in array[3] or 'DISCHARG' in array[3]:
                return self.into_activity_shipment(shipment)

            else:
                logger.info(f"Not found Load or Discharge in shipment {shipment}. Skipping to next shipment.")
                self._close_windows_util_reach_first_gscc()
                raise SkipToNextShipment(status_b="Skip", status_c="Shipment not Load or Discharge")

    def process_on_each_shipment_adhoc(self, shipment):
        logger: Logger = get_current_logger()

        window_adhoc: str = self._wait_for_window(shipment)
        self._window_title_stack.append(window_adhoc)
        gw.getWindowsWithTitle(window_adhoc)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        while True:
            pyautogui.hotkey('ctrl', 'k')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            if list_views.__len__() == 2:
                break
            self.sleep()

        list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")[1]
        items = list_views.items()

        if not items:
            logger.info('Not found any activity in {}, skip to next shipment'.format(shipment))
            self._close_windows_util_reach_first_gscc()
            raise Skipnoactivity(status_b="Skip", status_c="No activity found")

        runner = 0
        array = [None for _ in range(8)]

        for item in list_views.items():

            array[runner] = item.text()

            if runner != 7:
                runner = runner + 1
                continue

            runner = 0
            if 'LOAD' in array[3] or 'DISCHARG' in array[3]:
                return self.into_activity_shipment(shipment)

            else:
                logger.info(f"Not found 'Load' in shipment {shipment}. Skipping to next shipment.")
                self._close_windows_util_reach_first_gscc()
                raise SkipToNextShipment(status_b="Skip", status_c="Shipment not Load or Discharge")

    def into_activity_shipment(self, shipment):
        logger: Logger = get_current_logger()

        while True:
            pyautogui.hotkey('ctrl', 't')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")

            if list_views.__len__() == 1:
                break
            self.sleep()

        runner = 0
        capture_tasks = False
        tasks_to_capture = 0

        array = [None for _ in range(6)]
        list_of_activity_plan_open: list[_listview_item] = []
        list_of_activity_plan_closed: list[_listview_item] = []
        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        status_column_B = "Done"
        status_column_C = "Successfully closed"

        for item in listview_activity.items():
            array[runner] = item

            if runner != 5:
                runner = runner + 1
                continue

            runner = 0

            # Find 'OPS (EQUIPMENT PICKUP)' and start capturing the next 2 tasks
            if array[0].text().startswith('OPS (EQUIPMENT PICKUP)'):
                capture_tasks = True
                tasks_to_capture = 3

            # If capturing, take the next two tasks (regardless of their status)
            if capture_tasks is True and tasks_to_capture > 0:

                if array[0].text().startswith('OPS (') and (array[4].text() == 'Open' or array[4].text() == ''):
                    list_of_activity_plan_open.append(array[0])

                if array[0].text().startswith('OPS (') and array[4].text() == 'Closed':
                    list_of_activity_plan_closed.append(array[0])

                tasks_to_capture -= 1

                if tasks_to_capture == 0:
                    capture_tasks = False
        self.sleep()

        for activity_plan in list_of_activity_plan_open:
            activity_plan.select()
            pyautogui.hotkey('alt', 'l')
            activity_plan.deselect()
            self.sleep()

        # Recheck the status of all tasks after attempting to close
        array = [None for _ in range(6)]
        runner = 0
        capture_tasks = False

        list_of_activity_send_invoice: list[_listview_item] = []
        list_of_activity_send_invoice_closed: list[_listview_item] = []
        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        for item in listview_activity.items():
            array[runner] = item

            if runner != 5:
                runner += 1
                continue

            runner = 0

            # Find 'Send Prepaid Invoice Request' and check status
            if array[0].text().startswith('Send Prepaid Invoice Request'):
                capture_tasks = True

            if capture_tasks is True:

                if array[0].text().startswith('Send Prepaid Invoice Request') and (
                        array[4].text() == 'Open' or array[4].text() == ''):
                    list_of_activity_send_invoice.append(array[0])

                if array[0].text().startswith('Send Prepaid Invoice Request') and array[4].text() == 'Closed':
                    list_of_activity_send_invoice_closed.append(array[0])
                    status_column_B = 'Closed before'
                    status_column_C = f"By {array[2].text()}"

        if list_of_activity_send_invoice_closed:
            self._close_windows_util_reach_first_gscc()
            return status_column_B, status_column_C

        if list_of_activity_send_invoice:
            status_column_B, status_column_C = self.handle_task_prepaid()
            return status_column_B, status_column_C

        while True:
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            pyautogui.hotkey('ctrl', 't')
            if list_views.__len__() == 1:
                break
            self.sleep()

        self._close_windows_util_reach_first_gscc()
        return status_column_B, status_column_C

    def handle_invalid_window(self, shipment: str, workbook):
        logger: Logger = get_current_logger()

        self._wait_for_window('Invalid Booking Number')

        pyautogui.hotkey('enter')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('shift', 'tab')
        pyautogui.hotkey('down')
        pyautogui.hotkey('alt', 'k')
        return

    def handle_task_prepaid(self):
        logger: Logger = get_current_logger()
        max_retries = 3
        retry_count = 0

        while retry_count < max_retries:
            self.sleep()
            pyautogui.hotkey('alt', 'i')
            self.sleep()
            pyautogui.hotkey('q')
            self.sleep()
            pyautogui.hotkey('p')
            self.sleep()

            while True:
                list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
                pyautogui.hotkey('ctrl', 't')
                if list_views.__len__() == 1:
                    break
                self.sleep()

            # Recheck the status of all tasks after attempting to close
            array = [None for _ in range(6)]
            runner = 0
            capture_tasks = False
            open_task = 0

            listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

            for item in listview_activity.items():
                array[runner] = item

                if runner != 5:
                    runner += 1
                    continue

                runner = 0

                # Find 'Send Prepaid Invoice Request' and check status
                if array[0].text().startswith('Send Prepaid Invoice Request'):
                    capture_tasks = True

                if capture_tasks is True:
                    if array[0].text().startswith('Send Prepaid Invoice Request') and (
                            array[4].text() == 'Open' or array[4].text() == ''):
                        open_task += 1

            # If no open tasks found, exit the retry loop
            if open_task == 0:
                logger.info("Prepaid task closed successfully")
                return "Done", "Successfully closed"

            retry_count += 1
            logger.info(f"Prepaid task remains open, retrying ({retry_count}/{max_retries})")

        # If max retries reached and task still open, log the failure
        logger.info("Max retries reached, prepaid task could not be closed")
        return "Cannot close shipment", "Shipment remains open"


class SkipToNextShipment(Exception):
    def __init__(self, status_b: str = "Skip", status_c: str = "Shipment not Load or Discharge)"):
        self.status_b = status_b
        self.status_c = status_c
        super().__init__()


class Skipnoactivity(Exception):
    def __init__(self, status_b: str = "Skip", status_c: str = "No activity found"):
        self.status_b = status_b
        self.status_c = status_c
        super().__init__()

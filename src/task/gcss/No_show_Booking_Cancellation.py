from datetime import datetime
from logging import Logger
from typing import Callable

import pyautogui
import pygetwindow as gw
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ProcessUtil import get_matching_processes
from src.common.ThreadLocalLogger import get_current_logger
from src.common.exception.SkipEquipmentMatched import SkipEquipmentMatched
from src.common.exception.SkipOpsTaskClosed import SkipOpsTaskClosed
from src.common.exception.SkipRTDI import SkipRTDI
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.GCSSTask import GCSSTask


class No_show_Booking_Cancellation(GCSSTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.current_status_excel_col_index = None
        self.excel_provider: ExcelReaderProvider = None
        self.current_worksheet = None
        # self.current_status_excel_col_index: int = 0
        self.current_status_excel_row_index: int = 2

    def mandatory_settings(self) -> list[str]:
        desktop_mandatory_keys: list[str] = super().mandatory_settings()
        mandatory_keys: list[str] = ['excel.path',
                                     'excel.sheet',
                                     'excel.shipment',
                                     'excel.status.cell']
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
                self.sleep()

            self._wait_for_window('Pending Tray')

            logger.info("Start process shipment " + shipment)

            try:
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")

                pyautogui.hotkey('ctrl', 'o')
                pyautogui.hotkey('shift', 'tab')
                pyautogui.typewrite('Shipment')
                pyautogui.hotkey('tab')
                pyautogui.typewrite(shipment)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')
                self.sleep()
                self._wait_for_window(shipment)

                status_column_B, status_column_C = self.process_on_each_shipment(shipment)
                # try to save excel if shipment can be handled
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    status_column_B)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    status_column_C)
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)

                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                workbook = self.excel_provider.save(workbook)
                logger.info("Done with shipment " + shipment)

                print(
                    'END_FLOW current gcss process count {}'.format(len(get_matching_processes('GCSS'))))
                self.sleep()

            except SkipEquipmentMatched:
                logger.error(f'Face Equipment Matched {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Equipment Matched')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipOpsTaskClosed:
                logger.error(f'Face Ops Task Closed {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Ops Task Closed')
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

                self._post_actions()

                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Error - Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Cannot handle shipment {}, please check manual'.format(
                                                        shipment))
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)

                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

        try:
            self.excel_provider.save(workbook)
        except BaseException as e:
            logger.debug("Cannot save excel file")
        finally:
            self.excel_provider.close(workbook)
            self.excel_provider.quit_session()

    def process_on_each_shipment(self, shipment):

        self.sleep()
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

        return self.into_activity_shipment(shipment)

    def into_activity_shipment(self, shipment):
        logger: Logger = get_current_logger()

        # Shipment linked/matched.....
        self._check_equipment_matching(shipment)

        # OPS Task Opening
        self._checking_ops_task_open(shipment)
        status_column_B = "abc"
        status_column_C = "123"

        return status_column_B, status_column_C

    def _check_equipment_matching(self, shipment):
        logger: Logger = get_current_logger()

        # Shipment linked/matched.....
        while True:
            pyautogui.hotkey('ctrl', 'h')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")

            if list_views.__len__() == 1:
                break
            self.sleep()

        def find_editable_combobox(control):
            if control.class_name() == "ComboBox" and control.control_id() == 50001:
                # Kiểm tra xem ComboBox có chứa Edit không
                for child in control.children():
                    if child.class_name() == "Edit" and child.control_id() == 1001:
                        return control
            for child in control.children():
                result = find_editable_combobox(child)
                if result:
                    return result
            return None

        # Tìm ComboBox trong toàn bộ hierarchy
        target_combobox = find_editable_combobox(self._window)

        if target_combobox:
            for child in target_combobox.children():
                if child.class_name() == "Edit" and child.control_id() == 1001:
                    child.type_keys("Equipment Matching")
                    self.sleep()

            pyautogui.hotkey('alt', 'a')

            self.sleep()

        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        if len(listview_activity.items()) > 0:
            logger.info('Shipment {} has equipment matched'.format(shipment))
            raise SkipEquipmentMatched

        if len(listview_activity.items()) == 0:
            logger.info('Shipment {} does not have equipment matched'.format(shipment))

    def _checking_ops_task_open(self, shipment):
        logger: Logger = get_current_logger()

        while True:
            pyautogui.hotkey('ctrl', 't')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")

            if list_views.__len__() == 1:
                break
            self.sleep()

        runner = 0
        array = [None for _ in range(6)]

        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        for item in listview_activity.items():
            array[runner] = item

            if runner != 5:
                runner = runner + 1
                continue

            runner = 0

            # Find 'OPS (.... OUT)' - should be OPENED
            logger.info('Checking OPS Task')
            if array[0].text().startswith('OPS ('):
                if array[0].text().endswith('OUT)'):

                    if array[4].text() == 'Open' or array[4].text() == '':
                        logger.info('{} Opening'.format(array[0].text()))

                    if array[4].text() == 'Closed':
                        logger.info('{} is closed by {}'.format(array[0].text(), array[2].text()))
                        raise SkipOpsTaskClosed
                    self.sleep()

            # Find RTDI - should be OPEN
            logger.info('Checking RTDI')
            if array[0].text().startswith('Receive Transport Document Instructions'):
                if array[4].text() == 'Open' or array[4].text() == '':
                    logger.info('{} Opening'.format(array[0].text()))

                if array[4].text() == 'Closed':
                    logger.info('{} is closed by {}'.format(array[0].text(), array[2].text()))
                    raise SkipRTDI

            # Find Export Haulage - should be OPEN
            logger.info('Checking Export Haulage')

            # Find Confirm Booking - should be CLOSED
            if array[0].text().startswith('Receive Transport Document Instructions'):
                if array[4].text() == 'Open' or array[4].text() == '':
                    logger.info('{} Opening'.format(array[0].text()))

                if array[4].text() == 'Closed':
                    logger.info('{} is closed by {}'.format(array[0].text(), array[2].text()))

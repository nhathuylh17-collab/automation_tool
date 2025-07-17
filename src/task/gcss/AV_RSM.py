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
from src.common.exception.SkipTPDOC import SkipTPDOC
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.GCSSTask import GCSSTask


class AV_RSM(GCSSTask):

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

            except SkipTPDOC:
                logger.error(f'Face an TP doc exception {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip TPDoc')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Have TPDoc')
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

        while True:
            pyautogui.hotkey('ctrl', 't')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")

            if list_views.__len__() == 1:
                break
            self.sleep()

            # Hàm tìm ComboBox có Edit bên trong

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
                    child.type_keys("Documentation")
                    self.sleep()

        (status_column_B, status_column_C,
         list_of_activity_plan_seal, list_of_activity_plan_seal_closed) = self.get_status_at_row_rsm(shipment)

        # cover case TPDOC - more than 1 row Seal Mismatch is closed before
        if len(list_of_activity_plan_seal) > 1 or len(list_of_activity_plan_seal_closed) > 1:
            logger.debug('{} has TP Doc'.format(shipment))
            self._close_windows_util_reach_first_gscc()
            raise SkipTPDOC

        # cover case TPDOC - 1 is opened and 1 is closed - total 2 row
        if len(list_of_activity_plan_seal) == 1 and len(list_of_activity_plan_seal_closed) == 1:
            logger.debug('{} has TP Doc'.format(shipment))
            self._close_windows_util_reach_first_gscc()
            raise SkipTPDOC

        # Return "Closed by [user]" if already closed
        if list_of_activity_plan_seal_closed:
            self._close_windows_util_reach_first_gscc()
            return 'Closed before', status_column_C

        # normal shipment
        for plan_seal in list_of_activity_plan_seal:
            self.sleep()
            plan_seal.select()
            self.sleep()
            pyautogui.hotkey('alt', 'L')
            self.sleep()
            plan_seal.deselect()
            self.sleep()

        # recheck after trying to close shipment
        status_column_B, status_column_C, list_of_activity_plan_seal, list_of_activity_plan_seal_closed = self.get_status_at_row_rsm(
            shipment)

        if list_of_activity_plan_seal:
            logger.info('{} is still open'.format(shipment))
            return 'Cannot close shipment', 'Shipment remains open'

        if list_of_activity_plan_seal_closed:
            logger.info('{} is now closed'.format(shipment))
            return 'Done', 'Successfully closed'

        self._close_windows_util_reach_first_gscc()
        return status_column_B, status_column_C

    def get_status_at_row_rsm(self, shipment):
        logger: Logger = get_current_logger()
        runner = 0
        capture_tasks = False

        self.sleep()
        processing_cells: list[_listview_item] = [None] * 6
        self.sleep()

        list_of_activity_plan_seal: list[_listview_item] = []
        list_of_activity_plan_seal_closed: list[_listview_item] = []

        status_column_B = 'Skip'
        status_column_C = 'Cannot found RSM'

        self.sleep()
        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        for item in listview_activity.items():
            processing_cells[runner] = item

            if runner != 5:
                runner = runner + 1
                continue

            runner = 0

            if processing_cells[0].text().startswith('Resolve Seal Mismatch'):
                capture_tasks = True

            if capture_tasks is True:

                checking_row_title = processing_cells[0].text()
                checking_row_status = processing_cells[4].text()
                person = processing_cells[2].text()

                if checking_row_title.startswith('Resolve Seal Mismatch') and (
                        checking_row_status == 'Open' or checking_row_status == ''):
                    logger.info('Seal Mismatch is Open now')
                    list_of_activity_plan_seal.append(processing_cells[0])

                if checking_row_title.startswith(
                        'Resolve Seal Mismatch') and checking_row_status == 'Closed':
                    list_of_activity_plan_seal_closed.append(processing_cells[0])
                    logger.info('Seal Mismatch is closed before by {}'.format(person))
                    status_column_B = 'Closed'
                    status_column_C = f"By {person}"

            processing_cells = [None] * 6

        if not capture_tasks:
            logger.info('No RSM tasks found for shipment {}'.format(shipment))
            status_column_B = 'Skip'
            status_column_C = 'Cannot found RSM row'

        return status_column_B, status_column_C, list_of_activity_plan_seal, list_of_activity_plan_seal_closed

import time
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

            self._wait_for_window('Pending Tray')

            logger.info("Start process shipment " + shipment)

            try:
                #     try to interface and open shipment
                pyautogui.hotkey('ctrl', 'o')
                pyautogui.hotkey('shift', 'tab')
                pyautogui.typewrite('Shipment')
                pyautogui.hotkey('tab')
                pyautogui.typewrite(shipment)
                pyautogui.hotkey('tab')
                pyautogui.hotkey('enter')
                self.sleep()
                self._wait_for_window(shipment)

                status = self.process_on_each_shipment(shipment)
                # try to save excel if shipment can be handled
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    status)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                workbook = self.excel_provider.save(workbook)
                logger.info("Done with shipment " + shipment)

                print(
                    'END_FLOW current gcss process count {}'.format(len(get_matching_processes('GCSS'))))
                time.sleep(3)

            except SkipTPDOC:
                logger.error(f'Face an TP doc exception {shipment}')
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Have TPDoc')
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except BaseException as e:

                logger.info(f'Cannot handle shipment {shipment}. \n {e} \nMoving to next shipment')

                if len(get_matching_processes('GCSS')) == 0:
                    self._pre_actions()
                else:
                    self._close_windows_util_reach_first_gscc()

                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Cannot handle shipment {}, please check manual'.format(
                                                        shipment))
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

        self.sleep()
        window_normal_shipment: str = self._wait_for_window(shipment)
        self._window_title_stack.append(window_normal_shipment)
        gw.getWindowsWithTitle(window_normal_shipment)[0].activate()  # recheck chỗ này

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
        runner = 0
        capture_tasks = False

        self.sleep()
        array = [None for _ in range(6)]
        self.sleep()

        list_of_activity_plan_seal: list[_listview_item] = []
        list_of_activity_plan_seal_closed: list[_listview_item] = []
        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        status = "Done"

        for item in listview_activity.items():
            array[runner] = item

            if runner != 5:
                runner = runner + 1
                continue

            runner = 0

            if array[0].text().startswith('Resolve Seal Mismatch'):
                capture_tasks = True
            if capture_tasks is True:

                if array[0].text().startswith('Resolve Seal Mismatch') and (
                        array[4].text() == 'Open' or array[4].text() == ''):
                    logger.info('Seal Mismatch is Open now')
                    list_of_activity_plan_seal.append(array[0])

                if array[0].text().startswith('Resolve Seal Mismatch') and array[4].text() == 'Closed':
                    list_of_activity_plan_seal_closed.append(array[0])
                    logger.info('Seal Mismatch is closed before by {}'.format(array[2].text()))
                    status = f"Closed by {array[2].text()}"

        # cover case TPDOC - more than 1 row Seal Mismatch is closed before
        if len(list_of_activity_plan_seal) > 1 or len(list_of_activity_plan_seal_closed) > 1:
            self._close_windows_util_reach_first_gscc()
            raise SkipTPDOC()

        # cover case TPDOC - 1 is opened and 1 is closed - total 2 row
        if len(list_of_activity_plan_seal) == 1 and len(list_of_activity_plan_seal_closed) == 1:
            self._close_windows_util_reach_first_gscc()
            raise SkipTPDOC()

        # Return "Closed by [user]" if already closed
        if list_of_activity_plan_seal_closed:
            self._close_windows_util_reach_first_gscc()
            return status

        # normal shipment
        for plan_seal in list_of_activity_plan_seal:
            max_attempts = 3

            for attempt in range(1, max_attempts + 1):
                try:
                    plan_seal.select()
                    pyautogui.hotkey('alt', 'L')
                    self.sleep()
                    # Check status after attempt
                    listview_activity = self._window.children(class_name="SysListView32")[0]
                    runner = 0
                    array = [None for _ in range(6)]
                    for item in listview_activity.items():
                        array[runner] = item
                        if runner != 5:
                            runner = runner + 1
                            continue
                        runner = 0
                        if array[0].text() == plan_seal.text() and array[0].text().startswith(
                                'Resolve Customs Data Quality Issues'):
                            if array[4].text() not in ['Open', '']:
                                break  # Success, move to next activity_plan
                    else:
                        # If loop completes without breaking, check final attempt
                        if attempt == max_attempts:
                            status = "Cannot close shipment"
                            logger.info(
                                f"Failed to process activity plan {plan_seal.text()} after {max_attempts} attempts")
                            return status  # Exit early if one activity fails

                except Exception as e:
                    logger.error(f"Error during attempt {attempt} for activity plan {plan_seal.text()}: {e}")
                    self.sleep()
                    if attempt == max_attempts:
                        status = "Cannot close shipment"
                        logger.info(
                            f"Failed to process activity plan {plan_seal.text()} after {max_attempts} attempts")
                        return status  # Exit early if one activity fails

            while True:
                list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
                pyautogui.hotkey('ctrl', 't')
                if list_views.__len__() == 1:
                    break
                self.sleep()

        self._close_windows_util_reach_first_gscc()
        return status

    def select_menu_item(self, menu_item_name):
        """
        Chọn một mục trong thanh menu chính (menu header).
        menu_item_name: Tên của mục menu cần chọn (ví dụ: 'Manifest', 'File', 'Edit', v.v.).
        """
        logger: Logger = get_current_logger()
        try:
            # Truy cập thanh menu của cửa sổ
            self._window = self._window_spec.wrapper_object()
            self._window.menu_select(menu_item_name)
            self.sleep()
            logger.info('selected')

        except Exception as e:
            menu = self._window.menu()
            menu_items = menu.items()
            for item in menu_items:
                logger.debug(f"  - {item.text()}")
                logger.info('Cannot select')

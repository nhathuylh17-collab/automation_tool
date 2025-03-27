from logging import Logger
from typing import Callable

import pyautogui
import pygetwindow as gw
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.DesktopTask import DesktopTask


class AV_RCDQI(DesktopTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.current_status_excel_col_index = None
        self.excel_provider: ExcelReaderProvider = None
        self.current_worksheet = None
        # self.current_status_excel_col_index: int = 0
        self.current_status_excel_row_index: int = 5

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'excel.shipment', 'excel.status.cell']
        return mandatory_keys

    def automate(self):
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

        # col, row = extract_row_col_from_cell_pos_format(self._settings['excel.status.cell'])
        # self.current_status_excel_col_index: int = int(self.get_letter_position(col))
        # self.current_status_excel_row_index: int = int(row)

        self._wait_for_window('Pending Tray')
        self._window_title_stack.append('Pending Tray')

        self.current_element_count = 0
        self.total_element_size = len(shipments)

        for i, shipment in enumerate(shipments):

            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    logger.info("Currently pause")
                    self.pause_condition.wait()

                if self.terminated is True:
                    return
            self._wait_for_window('Pending Tray')
            self._window_title_stack.append('Pending Tray')

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
                self._window_title_stack.append(shipment)
                try:
                    #     try to handle shipment
                    self.process_on_each_shipment(shipment)
                    try:
                        # try to save excel if shipment can be handled
                        self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                            2,
                                                            'Done')
                        self.current_status_excel_row_index += 1
                        self.current_element_count += 1
                        self.excel_provider.save(workbook)
                        logger.info("Done with shipment " + shipment)

                    except Exception as e:
                        print("Error handle with shipment " + shipment)

                except SkipToNextShipment:
                    try:
                        # try to save excel and skip shipment
                        self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                            2,
                                                            'Skip')
                        self.current_status_excel_row_index += 1
                        self.current_element_count += 1
                        self.excel_provider.save(workbook)
                    except Exception as e:
                        print("Skip " + shipment)

                except SkipTPDOC:
                    try:
                        # try to save excel and skip shipment
                        self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                            2,
                                                            'Have TP Doc')
                        self.current_status_excel_row_index += 1
                        self.current_element_count += 1
                        self.excel_provider.save(workbook)
                    except Exception as e:
                        print("Have TP Doc " + shipment)

            except Exception:
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Cannot handle shipment {}, please check manual'.format(shipment))
                self.excel_provider.save(workbook)
                logger.info(f'Cannot handle shipment {shipment}. Moving to next shipment')

                self._wait_for_window("Pending Tray")
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                continue

        try:
            self.excel_provider.save(workbook)
        except Exception as e:
            logger.debug("Cannot save excel file")
        finally:
            # release file excel tránh bị cannot run in backround
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

        self.into_activity_shipment(shipment)

    def into_activity_shipment(self, shipment):
        logger: Logger = get_current_logger()
        self._wait_for_window(shipment)
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
        list_of_activity_plan: list[_listview_item] = []
        list_of_activity_plan_close: list[_listview_item] = []
        self.sleep()
        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        for item in listview_activity.items():
            array[runner] = item

            if runner != 5:
                runner = runner + 1
                continue

            runner = 0

            if array[0].text().startswith('Resolve Customs Data Quality Issues'):
                capture_tasks = True

            if capture_tasks is True:

                if array[0].text().startswith('Resolve Customs Data Quality Issues') and array[4].text() == 'Open':
                    logger.info('Data Quality is Open now')
                    list_of_activity_plan.append(array[0])

                if array[0].text().startswith('Resolve Customs Data Quality Issues') and array[4].text() == 'Closed':
                    list_of_activity_plan_close.append(array[0])
                    logger.info('Data Quality is closed before by {}'.format(array[2]))

        # cover IF we have TPDOC - more than 1 row has Resolve Customs Data Quality Issues
        if len(list_of_activity_plan) > 1 or len(list_of_activity_plan_close) > 1:
            logger.info('{} has TP Doc'.format(shipment))
            pyautogui.hotkey('alt', 'e')
            self.sleep()
            pyautogui.hotkey('left')
            self.sleep()
            pyautogui.hotkey('c')
            self.sleep()
            self._wait_for_window('Pending Tray')
            raise SkipTPDOC

        # TPDOC - 1 row open and 1 row closed
        if len(list_of_activity_plan) == 1 and len(list_of_activity_plan_close) == 1:
            pyautogui.hotkey('alt', 'e')
            self.sleep()
            pyautogui.hotkey('left')
            self.sleep()
            pyautogui.hotkey('c')
            self.sleep()
            self._wait_for_window('Pending Tray')
            raise SkipTPDOC

        # normal shipment
        for activity_plan in list_of_activity_plan:
            activity_plan.select()
            pyautogui.hotkey('alt', 'h')
            self.sleep()
            pyautogui.hotkey('left')
            self.sleep()
            # self.select_menu_item("Manifest")
            pyautogui.hotkey('down')
            self.sleep()
            pyautogui.hotkey('right')
            self.sleep()
            pyautogui.hotkey('down')
            pyautogui.hotkey('enter')
            activity_plan.deselect()
            self.sleep()

        # self.select_menu_item("File")
        pyautogui.hotkey('alt', 'e')
        self.sleep()
        pyautogui.hotkey('left')
        self.sleep()
        pyautogui.hotkey('c')
        self.sleep()
        self._wait_for_window('Pending Tray')

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
                logger.info('cannot select')


class SkipToNextShipment(Exception):
    pass


class SkipTPDOC(Exception):
    pass

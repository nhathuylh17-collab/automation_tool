from datetime import datetime
from logging import Logger
from typing import Callable

import pyautogui
import pygetwindow as gw
from pygetwindow import Win32Window
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item, _treeview_element, TreeViewWrapper
from pywinauto.controls.win32_controls import ComboBoxWrapper, ButtonWrapper, EditWrapper

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ProcessUtil import get_matching_processes
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.GCSSTask import GCSSTask


class AV_CID(GCSSTask):

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

                # In case CNEE SCV is not maintained yet
            except SkipToNextShipment_novessel:
                logger.error(f'No have any vessel information at {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Not found vessel')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)

                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipToNextShipment_noCnee:
                logger.error(f'Not found information of Consignee and Notify Party at {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Not found Consignee')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipToNextShipment_noFirstNotify:
                logger.error(f'Not found information of Consignee and Notify Party at {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Not found First Notify Party')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipToNextShipment_NotfoundtermCollect:
                logger.error(f'Not found term Collect at {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Not found term Collect')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipToNextShipment_ValidationFailed:
                logger.info(
                    f'Vadidation Failed {shipment}. \nMoving to next shipment')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                self._post_actions()
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Validation failed when try to add Invoice and Credit Party'.format(
                                                        shipment))
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipToNextShipment_CannotCompleteCollect:
                logger.info(
                    f'Vadidation Failed {shipment}. \nMoving to next shipment')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                self._post_actions()
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Cannot Complete Collect'.format(
                                                        shipment))
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except Exception as e:
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

        """Loop for each shipment in this method.
        Interface with 2nd window (Ctrl R - to get CNEE name)
        3rd window (Ctrl G - Maintain Pricing),
        4th window (Maintain Invoice Details).
        Return Payment Term, Invoice party index with CNEE name and Shipment in file Excel"""

        logger: Logger = get_current_logger()

        GCSS_Shipment_MSL_Active_Title: str = self._wait_for_window(shipment)
        self._window_title_stack.append(GCSS_Shipment_MSL_Active_Title)
        gw.getWindowsWithTitle(GCSS_Shipment_MSL_Active_Title)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        self.sleep()
        pyautogui.hotkey("ctrl", "k")

        # check vessel
        self.check_status_vessel(shipment)

        # check parties
        self.into_parties_tab(shipment)

        # Get Cnee Name and First Notify Party in 2nd window
        consignee_name: str = self.get_listview_text('Consignee')

        self.select_consignee()
        cnee_edit_element: EditWrapper = self._window.children(class_name="Edit")[14]
        cnee_scv_no: str = cnee_edit_element.texts()[0]

        first_notify_party: str = self.get_listview_text('First Notify Party')

        if consignee_name == '':
            raise SkipToNextShipment_noCnee

        if first_notify_party == '':
            raise SkipToNextShipment_noFirstNotify

        invoice_party_name: str = self.get_listview_text(search_text_required='Invoice Party',
                                                         search_text_additional=consignee_name)
        credit_party_name: str = self.get_listview_text(search_text_required='Credit Party',
                                                        search_text_additional=consignee_name)

        if invoice_party_name == '' or credit_party_name == '':
            logger.info('Invoice and Credit parties not found, processing to update these parties')
            self.adding_invoice_and_credit_parties(shipment)
            logger.info('Updated Invoice and Credit Parties')

        # 'Get Invoice Tab'
        self._wait_for_window(shipment)

        self.into_freight_and_pricing_tab()
        self.sleep()

        self._wait_for_window(shipment)
        self.into_maintain_pricing_tab()

        GCSS_window_maintain: str = self._wait_for_window('Maintain Pricing and Invoicing')
        self._window_title_stack.append(GCSS_window_maintain)
        gw.getWindowsWithTitle(GCSS_window_maintain)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        # 'Open and interface with 3rd window - Maintain Pricing and Invoicing Window'
        self.sleep()
        self.count_and_choose_all_item_payment_term_collect()
        self.sleep()

        # click button Modify in tab 2nd - Invoice tab
        buttons: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_modify: ButtonWrapper = buttons[10]
        button_modify.click()

        # Into Maintain Invoice Details
        GCSS_window_maintain: str = self._wait_for_window('Maintain Invoice Details')
        self._window_title_stack.append(GCSS_window_maintain)
        gw.getWindowsWithTitle(GCSS_window_maintain)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        ComboBox_maintain_payment: ComboBoxWrapper = self._window.children(class_name="ComboBox")[0]
        ComboBox_maintain_payment.select('Collect')
        logger.info('Choose Collect')
        self.sleep()

        # when choosing this payment term, it will automate show a dialog Information
        dialog_title_infor = 'Information'
        dialog_window = self._app.window(title=dialog_title_infor)
        try:
            dialog_window.wait(timeout=5, wait_for='visible')
            dialog_window.type_keys('{ENTER}')
            logger.info('Press enter')
        except Exception:
            pyautogui.hotkey('tab')

        ComboBox_maintain_invoice_party: ComboBoxWrapper = self._window.children(class_name="ComboBox")[1]
        item_invoices: list[str] = ComboBox_maintain_invoice_party.item_texts()

        runner: int = 0
        for invoice in item_invoices:

            if cnee_scv_no in invoice:
                ComboBox_maintain_invoice_party.select(runner)
                logger.info('Selecting Invoice and Credit Party')

                dialog_title_qs = 'Question'
                dialog_window = self._app.window(title=dialog_title_qs)
                dialog_window.wait(timeout=10, wait_for='visible')
                dialog_window.type_keys('{ENTER}')

                break

            runner += 1

        self.sleep()
        ComboBox_maintain_collect_business: ComboBoxWrapper = self._window.children(class_name="ComboBox")[3]
        ComboBox_maintain_collect_business.select('Maersk Bangkok (Bangkok)')
        logger.info('Choose Maersk bangkok')
        self.sleep()

        ComboBox_maintain_printable_freight_line: ComboBoxWrapper = self._window.children(class_name="ComboBox")[4]
        ComboBox_maintain_printable_freight_line.select('Yes')
        logger.info('Choose Yes')
        self.sleep()

        # Click OK button in 4th window and window will be auto closed
        number_of_titles_before = len(gw.getAllTitles())
        pyautogui.hotkey('alt', 'k')

        # while len(gw.getAllTitles()) == number_of_titles_before:
        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_ok: ButtonWrapper = list_btn[0]
        button_ok.click()

        number_of_titles_after = len(gw.getAllTitles())

        if number_of_titles_after > number_of_titles_before:
            self._window_title_stack.append('Validation failed')
            raise SkipToNextShipment_ValidationFailed

        self._window_title_stack.pop()
        self._wait_for_window('Maintain Pricing and Invoicing')
        self._window_title_stack.peek()

        collect_details_collect_status: EditWrapper = self._window.children(class_name="Edit")[3]

        try_times = 0
        while True:
            if try_times > 4:
                raise SkipToNextShipment_CannotCompleteCollect

            pyautogui.hotkey('alt', 't')
            if collect_details_collect_status.texts()[0] == 'Yes':
                break

            try_times += 1
            self.sleep()

        self._close_windows_util_reach_first_gscc()
        return "Complete", "Done"

    def count_and_choose_all_item_payment_term_collect(self) -> int:
        """"
            return the number of item payment term collect have been clicked
        """
        logger: Logger = get_current_logger()

        list_views: ListViewWrapper = self._window.children(class_name="SysListView32")[1]

        count_item: int = 0

        pyautogui.keyDown('ctrl')
        for item in list_views.items():
            item: _listview_item
            if str(item.text()).__contains__('Collect'):
                item.select()
                count_item += 1

        pyautogui.keyUp('ctrl')
        if count_item == 0:
            raise SkipToNextShipment_NotfoundtermCollect

        logger.info('Total have {} row Collect'.format(count_item))
        return count_item

    def into_freight_and_pricing_tab(self):
        while True:
            pyautogui.hotkey('ctrl', 'g')

            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            if list_views.__len__() == 6:
                break

            self.sleep()

    def into_maintain_pricing_tab(self):
        while True:
            pyautogui.hotkey('alt', 'y')
            if self._is_current_window_having_title('Maintain Pricing and Invoicing') is True:
                self._window_title_stack.peek()
                get_current_logger().info('Get window Maintain')
                break
            # list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            self.sleep()

    def into_parties_tab(self, shipment):

        while True:
            self._wait_for_window(shipment)

            pyautogui.hotkey('ctrl', 'r')
            list_views_parties: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            tree_views_parties: list[TreeViewWrapper] = self._window.children(class_name="SysTreeView32")

            if len(list_views_parties) == 2 and len(tree_views_parties) == 0:
                get_current_logger().info('Going to Parties Panel')
                break

            self.sleep()

    def get_listview_text(self, search_text_required: str, search_text_additional: str = None) -> str:
        logger: Logger = get_current_logger()

        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        runner = 0
        processing_cells: list[_listview_item] = [None] * 2
        result_text = ''

        for item in listview_activity.items():
            processing_cells[runner] = item

            if runner != 1:
                runner = runner + 1
                continue

            runner = 0
            # Check both conditions if search_text_1 is provided
            if search_text_additional and processing_cells[0].text().startswith(search_text_additional) and \
                    processing_cells[1].text().startswith(search_text_required):

                processing_cells[0].select()

                result_text = processing_cells[0].text()
                logger.info(f'{search_text_required} found, result is {result_text}')
                return result_text
            # Check only search_text_2 if search_text_1 is not provided
            elif not search_text_additional and processing_cells[1].text().startswith(search_text_required):

                processing_cells[0].select()

                result_text = processing_cells[0].text()
                logger.info(f'{search_text_required} is {result_text}')
                return result_text
            processing_cells = [None] * 2

        logger.info(
            f'Not found {search_text_required}' + (f' with {search_text_additional}' if search_text_additional else ''))
        return result_text

    def adding_invoice_and_credit_parties(self, shipment):

        self.get_listview_text('Consignee')
        cnee_edit_element: EditWrapper = self._window.children(class_name="Edit")[14]
        cnee_scv_no: str = cnee_edit_element.texts()[0]

        self._window = self._hotkey_then_open_new_window('Party Details', 'alt', 'a')

        self._window = self._hotkey_then_open_new_window('Customer Search', 'alt', 'c')

        self.sleep()

        def __find_editable_combobox(control):
            if control.class_name() == "ComboBox" and control.control_id() == 50001:
                for child in control.children():
                    if child.class_name() == "Edit" and child.control_id() == 1001:
                        return control
            for child in control.children():
                result = __find_editable_combobox(child)
                if result:
                    return result
            return None

        # Tìm ComboBox trong toàn bộ hierarchy
        target_combobox = __find_editable_combobox(self._window)

        if target_combobox:
            for child in target_combobox.children():
                if child.class_name() == "Edit" and child.control_id() == 1001:
                    child.type_keys("Customer ID")
                    self.sleep()

        pyautogui.hotkey('tab')
        self.sleep()
        pyautogui.write(cnee_scv_no)
        self.sleep()
        self._hotkey_then_close_current_window('enter')
        self.sleep()

        list_views: ListViewWrapper = self._window.children(class_name="SysListView32")[0]
        try:
            pyautogui.keyDown('ctrl')
            for item in list_views.items():
                if item.text() == 'Invoice Party':
                    item.select()
                    continue
                if item.text() == 'Credit Party':
                    item.select()
                    continue
        finally:
            pyautogui.keyUp('ctrl')

        # click btn >>
        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_right: ButtonWrapper = list_btn[3]
        button_right.click()

        # Click OK button to close Party Details
        pyautogui.hotkey('alt', 'k')

        self.sleep()

        current_window_title: Win32Window = gw.getActiveWindow().title

        if current_window_title == 'Validation failed':
            raise SkipToNextShipment_ValidationFailed

        self._window_title_stack.pop()
        self._wait_for_window(shipment)
        self._window_title_stack.peek()

    def check_status_vessel(self, shipment):
        logger: Logger = get_current_logger()
        self._wait_for_window(shipment)
        self._window_title_stack.append(shipment)

        while True:
            self.sleep()
            pyautogui.hotkey('ctrl', 'k')
            self.sleep()

            tree_views: list[TreeViewWrapper] = self._window.children(class_name="SysTreeView32")
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")

            if tree_views.__len__() == 1 and list_views.__len__() == 2:
                break
            self.sleep()

        tree_view: TreeViewWrapper = tree_views[0]
        root_items: list[_treeview_element] = tree_view.roots()

        found_mvs = False
        for target_root in root_items:
            target_text = target_root.select().text()

            if "MVS" in target_text:
                found_mvs = True
                logger.info('Has MVS')
                children = target_root.children()
                break
        if not found_mvs:
            raise SkipToNextShipment_novessel
        eta_found = False

        for child in target_root.children():
            item_text = child.text()
            if item_text.startswith("ETA: ") and len(item_text) > 5:  # Check if "ETD: " has data after it
                eta_found = True
                logger.info(f"Vessel checked ETA for shipment {shipment}")
                break

        if not eta_found:
            logger.info(f"No valid ETA found for shipment {shipment}")
            raise SkipToNextShipment_novessel()

    def print_all_controls(self, control, depth=0):
        """Recursively print all controls and their children."""
        try:
            # Print the current control's details
            print("  " * depth + f"Class Name: {control.class_name()}, "
                                 f"Text: {control.window_text()}, "
                                 f"Control ID: {control.control_id()}")

            # Recursively print all children of the current control
            for child in control.children():
                self.print_all_controls(child, depth + 1)
        except Exception as e:
            print("  " * depth + f"Error accessing control: {e}")

    def select_consignee(self):
        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        runner = 0
        processing_cells: list[_listview_item] = [None] * 2
        result_text = ''

        for item in listview_activity.items():
            processing_cells[runner] = item

            if runner != 1:
                runner = runner + 1
                continue

            runner = 0
            # Check both conditions if search_text_1 is provided
            if processing_cells[1].text().startswith('Consignee'):
                processing_cells[0].select()
                return
            return


class SkipToNextShipment_novessel(Exception):
    pass


class SkipToNextShipment_noCnee(Exception):
    pass


class SkipToNextShipment_noFirstNotify(Exception):
    pass


class SkipToNextShipment_ValidationFailed(Exception):
    pass


class SkipToNextShipment_CannotCompleteCollect(Exception):
    pass


class SkipToNextShipment_NotfoundtermCollect(Exception):
    pass

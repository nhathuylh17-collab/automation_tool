from datetime import datetime
from logging import Logger
from typing import Callable
from typing import Optional

import pyautogui
import pygetwindow as gw
from pygetwindow import Win32Window
from pywinauto import Application, WindowSpecification
from pywinauto.controls.common_controls import ListViewWrapper, _listview_item, TreeViewWrapper
from pywinauto.controls.uia_controls import ButtonWrapper, ComboBoxWrapper
from pywinauto.controls.win32_controls import EditWrapper

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ProcessUtil import get_matching_processes
from src.common.ThreadLocalLogger import get_current_logger
from src.common.exception.SkipETD import SkipETD
from src.common.exception.SkipEmail import SkipEmail
from src.common.exception.SkipEquipmentMatched import SkipEquipmentMatched
from src.common.exception.SkipPlaceofReceipt import SkipPlaceofReceipt
from src.common.exception.SkipToNextShipment_ValidationFailed import SkipToNextShipment_ValidationFailed
from src.common.exception.SkipToNextShipment_noBookedBy import SkipToNextShipment_noBookedBy
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

            except SkipETD:
                logger.error(f'ETD not passed - {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Recheck ETD')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipEmail:
                logger.error(f'Not found email to sent Cancel Booking - {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Not found email')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    4,
                                                    current_timestamp)
                self.current_status_excel_row_index += 1
                self.current_element_count += 1
                self.excel_provider.save(workbook)
                continue

            except SkipPlaceofReceipt:
                logger.error(f'Place of receipt not in Vietnam - {shipment}')
                current_timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
                # try to save excel and skip shipment
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    2,
                                                    'Skip')
                self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                    3,
                                                    'Place of Receipt not in Vietnam')
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
        GCSS_Shipment_MSL_Active_Title: str = self._wait_for_window(shipment)
        self._window_title_stack.append(GCSS_Shipment_MSL_Active_Title)
        gw.getWindowsWithTitle(GCSS_Shipment_MSL_Active_Title)[0].activate()

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
        self._wait_for_window(shipment)

        while True:
            pyautogui.hotkey('ctrl', 'k')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            if list_views.__len__() == 2:
                break
            self.sleep()

        logger.info('Checking product delivery')
        type_of_product_delivery: str = self._check_product_delivery(shipment)
        self.sleep()

        logger.info('Checking place of receipt')
        # self._check_place_of_receipt(shipment)
        self.sleep()

        logger.info('Checking ETD')
        self._check_ETD(shipment)
        self.sleep()

        logger.info('Checking equipment matched or linked')
        self._check_equipment_matching(shipment)
        self.sleep()

        while True:
            pyautogui.hotkey('ctrl', 'k')
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")
            if list_views.__len__() == 2:
                break
            self.sleep()

        if type_of_product_delivery == 'SPOT':
            logger.info('Handling SPOT shipment')
            self._handle_spot_shipment(shipment)
            status_column_B = "Done"
            status_column_C = "Canceled SPOT Booking"
            return status_column_B, status_column_C

        if type_of_product_delivery == 'Normal':
            logger.info('Handling normal shipment')
            self._handle_normal_shipment(shipment)
            status_column_B = "Done"
            status_column_C = "Canceled Normal Booking"
            return status_column_B, status_column_C

    def _check_ETD(self, shipment):
        logger: Logger = get_current_logger()
        # TreeView processing to extract ETD from first (P) node
        tree_views: list[TreeViewWrapper] = self._window.children(class_name="SysTreeView32")
        if not tree_views:
            logger.error("No TreeView found in the window")
            return None

        tree_view: TreeViewWrapper = tree_views[0]
        logger.debug("Found TreeView control")

        # Get all root nodes
        root_nodes = tree_view.roots()
        if not root_nodes:
            logger.error("No root nodes found in TreeView")
            return None

        # Find the first root node starting with "(P)"
        first_p_node = root_nodes[0]

        # Search for ETD node among children of the first (P) node
        etd_node_data = None
        for child in first_p_node.children():
            child_text = child.text()
            if child_text.startswith("ETD:"):
                control_id = child.control_id() if hasattr(child, 'control_id') else "N/A"
                automation_id = child.automation_id() if hasattr(child, 'automation_id') else "N/A"
                etd_node_data = {
                    "text": child_text,
                    "control_id": control_id,
                    "automation_id": automation_id
                }
                logger.debug(
                    f"Found ETD: Text = {child_text}, Control ID = {control_id}, Automation ID = {automation_id}")
                break

        if not etd_node_data:
            logger.error("No ETD node found under the first (P) node")
            self._close_windows_util_reach_first_gscc()
            raise SkipETD

        try:
            etd_text = etd_node_data["text"]

            etd_date_str = etd_text.replace("ETD: ", "").rstrip("*")

            etd_date = datetime.strptime(etd_date_str, "%m/%d/%Y %I:%M:%S %p")
            logger.info(f"Parsed ETD date: {etd_date}")

            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

            self.sleep()
            # Compare ETD with today
            if etd_date.date() < today.date():
                logger.info(
                    f"ETD {etd_date.date()} passed, continuing with shipment {shipment}")
                pyautogui.hotkey('ctrl', 'k')

                return etd_node_data

            else:
                logger.info(
                    f"ETD {etd_date.date()} is on or after today {today.date()}, skipping shipment {shipment}")
                self._close_windows_util_reach_first_gscc()
                raise SkipETD(f"ETD {etd_date.date()} is on or after today {today.date()}")

        except ValueError as e:
            logger.error(f"Failed to parse ETD date for shipment {shipment}: {e}")
            return None

    def _check_product_delivery(self, shipment):
        logger: Logger = get_current_logger()
        controls = self._window.descendants()
        for control in controls:
            class_name = control.class_name()
            text = control.window_text()
            control_id = control.control_id()

            if class_name == 'Edit' and control_id == 50003:
                if ('SPOT' in text) or ('Spot' in text):
                    logger.info('this is SPOT shipment')
                    return 'SPOT'
                else:
                    logger.info('Checked product - not SPOT shipment')
                    return 'Normal'

    def _check_place_of_receipt(self, shipment):
        logger: Logger = get_current_logger()
        controls = self._window.descendants()
        for control in controls:
            class_name = control.class_name()
            text = control.window_text()
            control_id = control.control_id()

            if class_name == 'Edit' and control_id == 50007:
                if 'Vietnam' in text:
                    logger.info('Checked place of receipt {}'.format(text))
                else:
                    logger.info('This shipment not in Vietnam')
                    self._close_windows_util_reach_first_gscc()
                    raise SkipPlaceofReceipt

    def _check_equipment_matching(self, shipment):
        logger: Logger = get_current_logger()
        self.sleep()

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

        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_ok: ButtonWrapper = list_btn[1]
        button_ok.click()

        self.sleep()

        listview_activity: ListViewWrapper = self._window.children(class_name="SysListView32")[0]

        if len(listview_activity.items()) > 0:
            logger.info('Shipment {} has equipment matched'.format(shipment))
            self._close_windows_util_reach_first_gscc()
            raise SkipEquipmentMatched

        if len(listview_activity.items()) == 0:
            logger.info('Shipment {} does not have equipment matched'.format(shipment))

    def _handle_spot_shipment(self, shipment):
        logger: Logger = get_current_logger()
        self._wait_for_window(shipment)

        while True:
            self.sleep()
            pyautogui.hotkey('ctrl', 'm')
            self.sleep()

            tree_views: list[TreeViewWrapper] = self._window.children(class_name="SysTreeView32")
            list_views: list[ListViewWrapper] = self._window.children(class_name="SysListView32")

            if tree_views.__len__() == 1 and list_views.__len__() == 0:
                break
            self.sleep()

        self.into_properties_booking_tab(shipment)

        self.into_parties_tab(shipment)

        booked_by_name: str = self.get_listview_text('Booked By')

        self.select_booked_by()

        booked_by_element: EditWrapper = self._window.children(class_name="Edit")[14]
        booked_by_scv_no: str = booked_by_element.texts()[0]

        if booked_by_name == '':
            raise SkipToNextShipment_noBookedBy

        invoice_party_name: str = self.get_listview_text(search_text_required='Invoice Party',
                                                         search_text_additional=booked_by_name)
        credit_party_name: str = self.get_listview_text(search_text_required='Credit Party',
                                                        search_text_additional=booked_by_name)

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

        # 'Open and interface with 3rd window - Maintain Pricing and Invoicing Window
        # ___________________today'
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

            if booked_by_scv_no in invoice:
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

    def _handle_normal_shipment(self, shipment):
        logger: Logger = get_current_logger()
        self._wait_for_window(shipment)

        logger.info('Canceling normal booking {}'.format(shipment))
        self.into_cancel_booking_tab()

        GCSS_window_cancel_booking: str = self._wait_for_window('Cancel Booking')
        self._window_title_stack.append(GCSS_window_cancel_booking)
        gw.getWindowsWithTitle(GCSS_window_cancel_booking)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        def find_editable_combobox_in_cancel_booking_tab(control):
            if control.class_name() == "ComboBox" and control.control_id() == 50001:
                for child in control.children():
                    if child.class_name() == "Edit" and child.control_id() == 1001:
                        return control
            for child in control.children():
                result = find_editable_combobox_in_cancel_booking_tab(child)
                if result:
                    return result
            return None

        target_combobox = find_editable_combobox_in_cancel_booking_tab(self._window)

        if target_combobox:
            for child in target_combobox.children():
                if child.class_name() == "Edit" and child.control_id() == 1001:
                    child.type_keys("Cargo not ready")
                    self.sleep()

        self._select_recipients_email()

        message_normal_shipment: str = 'Cargo no show - Unmaterialized booking past ETD without equipment pick up yet.'
        message_box = self._window.child_window(control_id=50008, class_name="Edit")
        message_box.set_edit_text(message_normal_shipment)

        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_ok: ButtonWrapper = list_btn[2]
        button_ok.click()
        self.sleep()

        window_cancel_booking: str = self._wait_for_window('Warning')
        self._window_title_stack.append(window_cancel_booking)
        gw.getWindowsWithTitle(window_cancel_booking)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_ok: ButtonWrapper = list_btn[0]
        button_ok.click()
        self.sleep()

        self._window_title_stack.pop()
        self._window_title_stack.pop()
        self._wait_for_window(shipment)
        logger.info('Canceled Booking')
        self._close_windows_util_reach_first_gscc()

    def into_cancel_booking_tab(self):
        while True:
            pyautogui.hotkey('alt', 's')
            self.sleep()

            pyautogui.hotkey('a')
            self.sleep()

            if self._is_current_window_having_title('Cancel Booking') is True:
                self._window_title_stack.peek()
                get_current_logger().info('Get window Cancel Booking')
                break
            self.sleep()

    def _select_recipients_email(self):
        logger: Logger = get_current_logger()
        # TreeView processing...
        tree_views: list[TreeViewWrapper] = self._window.children(class_name="SysTreeView32")
        if not tree_views:
            logger.error("No TreeView found in the window")
            return None

        tree_view: TreeViewWrapper = tree_views[1]  # Hoặc [0] tùy UI thực tế
        logger.debug("Found TreeView control")

        root_nodes = tree_view.roots()
        if not root_nodes:
            logger.error("No root nodes found in TreeView")
            return None

        first_node = root_nodes[0]

        email_node = None
        try:
            first_node.expand()
            first_node.ensure_visible()
            for child_level1 in first_node.children():
                target_node = self._expand_to_target(child_level1, 'E-mail')
                if target_node:
                    email_node = target_node
                    break
        except Exception as e:
            logger.error(f"Error during expansion: {e}")
            self._close_windows_util_reach_first_gscc()
            raise SkipEmail

        if not email_node:
            self._close_windows_util_reach_first_gscc()
            raise SkipEmail

        # Lấy thông tin email
        email_text = email_node.text()
        control_id = email_node.control_id() if hasattr(email_node, 'control_id') else "N/A"
        automation_id = email_node.automation_id() if hasattr(email_node, 'automation_id') else "N/A"

        email = {
            "text": email_text,
            "control_id": control_id,
            "automation_id": automation_id
        }
        logger.debug(f"Found email: Text = {email_text}, Control ID = {control_id}, Automation ID = {automation_id}")

        email_node.select()

        if not email:
            logger.error("No ETD node found under the first (P) node")
            self._close_windows_util_reach_first_gscc()
            raise SkipEmail

    def _check_treeview_elements(self, shipment):
        logger: Logger = get_current_logger()
        # Locate the TreeView control
        tree_views: list[TreeViewWrapper] = self._window.children(class_name="SysTreeView32")
        if not tree_views:
            logger.error("No TreeView found in the window")
            return

        tree_view: TreeViewWrapper = tree_views[0]  # Assume first TreeView
        logger.info("Found TreeView control")

        # Get root nodes
        root_nodes = tree_view.roots()
        if not root_nodes:
            logger.error("No root nodes found in TreeView")
            return

        # Recursive function to traverse and extract node details
        def traverse_tree_node(node, depth=0):
            # Get node properties
            node_text = node.text()
            control_id = node.control_id() if hasattr(node, 'control_id') else "N/A"
            automation_id = node.automation_id() if hasattr(node, 'automation_id') else "N/A"
            node_rect = node.rectangle()
            logger.info(
                f"{'  ' * depth}Node: Text = {node_text}, Control ID = {control_id}, Automation ID = {automation_id}, Rectangle = {node_rect}")

            # Get child nodes
            children = node.children()
            for child in children:
                traverse_tree_node(child, depth + 1)

        # Traverse all root nodes
        logger.info("Traversing TreeView nodes")
        for root_node in root_nodes:
            traverse_tree_node(root_node)

        # Optional: Get all nodes in a flat list
        all_nodes = tree_view.get_all_items()
        node_data = [{"text": node.text(), "control_id": node.control_id() if hasattr(node, 'control_id') else "N/A"}
                     for node in all_nodes]
        logger.info(f"TreeView Nodes: {node_data}")

    def _expand_to_target(self, node: TreeViewWrapper, target_text: str) -> Optional[TreeViewWrapper]:
        logger: Logger = get_current_logger()
        try:
            node.expand()
            logger.debug(f"Expanded node: {node.text()}")
            node.ensure_visible()  # Đảm bảo visible trong view
        except Exception as e:
            logger.warning(f"Failed to expand node {node.text()}: {e}")
            return None

        for child in node.children():
            child_text = child.text()
            if target_text in child_text:
                child.ensure_visible()
                return child  # Tìm thấy target
            # Đệ quy nếu child có sub-children tiềm năng
            sub_target = self._expand_to_target(child, target_text)
            if sub_target:
                return sub_target
        return None

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

    def into_properties_booking_tab(self, shipment):
        # Click button "Properties"
        logger: Logger = get_current_logger()
        logger.info('Handling Booking Properties Tab')

        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_ok: ButtonWrapper = list_btn[11]
        button_ok.click()

        self.sleep()

        window_booking_properties = self._wait_for_window('Booking Properties - {}'.format(shipment))
        self._window_title_stack.append('Booking Properties - {}'.format(shipment))
        gw.getWindowsWithTitle(window_booking_properties)[0].activate()

        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        list_combobox: list[ComboBoxWrapper] = self._window.children(class_name="ComboBox")

        TPD_chanels_combobox = list_combobox[3]
        for child in TPD_chanels_combobox.children():
            if child.class_name() == "Edit" and child.control_id() == 1001:
                child.type_keys("Failed EDI")
                self.sleep()

        Document_group_combobox = list_combobox[5]
        for child in Document_group_combobox.children():
            if child.class_name() == "Edit" and child.control_id() == 1001:
                child.type_keys("MSL TPDoc Sea Waybill Shipped")
                self.sleep()

        Document_type_combobox = list_combobox[6]
        for child in Document_type_combobox.children():
            if child.class_name() == "Edit" and child.control_id() == 1001:
                child.type_keys("Shipped on Board")
                self.sleep()

        list_btn: list[ButtonWrapper] = self._window.children(class_name="Button")
        button_ok: ButtonWrapper = list_btn[8]
        button_ok.click()
        self.sleep()

        self._window_title_stack.pop()
        self._wait_for_window(shipment)
        gw.getWindowsWithTitle(self._wait_for_window(shipment))[0].activate()
        self._app: Application = Application().connect(title=self._window_title_stack.peek())
        self._window: WindowSpecification = self._app.window(title=self._window_title_stack.peek())

        logger.info('Done setting {} in Booking Properties panel'.format(shipment))

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

    def select_booked_by(self):
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
            if processing_cells[1].text().startswith('Booked By'):
                processing_cells[0].select()
                return
            return

    def adding_invoice_and_credit_parties(self, shipment):

        self.get_listview_text('Booked By')
        booked_by_element: EditWrapper = self._window.children(class_name="Edit")[14]
        booked_by_scv_no: str = booked_by_element.texts()[0]

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
        pyautogui.write(booked_by_scv_no)
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

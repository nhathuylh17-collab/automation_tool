import os
from logging import Logger
from typing import Callable, Optional

from PyPDF2 import PdfMerger

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


# noinspection PyPackageRequirements
class PDFCombine_BillandCont(AutomatedTask):
    doc_to_bill: dict[str, str] = {}

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

        self.excel_provider: ExcelReaderProvider = None
        self.current_worksheet = None
        self.current_status_excel_row_index: int = 2

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'folder_bill.folder', 'folder_doc.folder',
                                     'folder_output.folder', 'excel.column.bill',
                                     'excel.column.cont']
        return mandatory_keys

    def automate(self):
        logger: Logger = get_current_logger()
        self.excel_provider: ExcelReaderProvider = XlwingProvider()

        path_to_excel = self._settings['excel.path']
        workbook = self.excel_provider.get_workbook(path=path_to_excel)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        self.current_worksheet = self.excel_provider.get_worksheet(workbook, sheet_name)

        folder_bill = self._settings['folder_bill.folder']
        folder_doc = self._settings['folder_doc.folder']
        folder_output = self._settings['folder_output.folder']

        os.makedirs(folder_output, exist_ok=True)

        col_bill = self._settings['excel.column.bill']
        col_cont = self._settings['excel.column.cont']

        bill_ids: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                    self._settings['excel.sheet'],
                                                                    self._settings[
                                                                        'excel.column.bill'])

        cont_ids: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                    self._settings['excel.sheet'],
                                                                    self._settings['excel.column.cont'])

        if len(bill_ids) == 0:
            logger.error('Input booking id list is empty ! Please check again')

        if len(bill_ids) != len(cont_ids):
            raise Exception("Please check your input data length of bills, type_bill are not equal")

        self.total_element_size = len(bill_ids)
        self.current_element_count = 0

        for idx, (bill_id, cont_id) in enumerate(zip(bill_ids, cont_ids), start=2):
            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            logger.info("Processing: {}".format(bill_id))

            status_column_C = self.combine_bill_and_cont(
                bill_id=bill_id,
                cont_id=cont_id,
                folder_bill=folder_bill,
                folder_doc=folder_doc,
                output_folder=folder_output,
                logger=logger
            )

            self.excel_provider.change_value_at(self.current_worksheet, self.current_status_excel_row_index,
                                                3,
                                                status_column_C)
            self.current_element_count += 1
            self.current_status_excel_row_index += 1
            self.excel_provider.save(workbook)

        # Close the Excel workbook
        self.excel_provider.close(workbook)
        self.excel_provider.quit_session()
        logger.info('Done')

    def combine_bill_and_cont(self, bill_id: str, cont_id: str, folder_bill: str, folder_doc: str,
                              output_folder: str, logger: Logger):
        """
        Gộp file Bill (trước) + Cont (sau) → output file tên = cont_id.pdf
        """

        logger: Logger = get_current_logger()
        merger = PdfMerger()
        combined = False

        # find Bill
        bill_path = self._find_pdf_recursive(folder_bill, bill_id)
        if bill_path:
            if not self._safe_append_pdf(merger, bill_path, "Bill", logger):
                logger.error("Skip, file {} error".format(bill_path))
            else:
                combined = True
        else:
            logger.warning("Cannot find file {}".format(bill_id))

        # find Cont
        cont_path = self._find_pdf_recursive(folder_doc, cont_id)
        if cont_path:
            if not self._safe_append_pdf(merger, cont_path, "Cont", logger):
                logger.error("Skip, file {} error".format(cont_path))
            else:
                combined = True
        else:
            logger.warning("Cannot found file {}".format(cont_id))

        # combine
        if combined:
            output_path = os.path.join(output_folder, f"{cont_id}.pdf")
            try:
                # Xóa file cũ nếu tồn tại
                if os.path.exists(output_path):
                    os.remove(output_path)
                with open(output_path, 'wb') as f:
                    merger.write(f)
                logger.info("Done")
                status_column_C = 'Done'

            except Exception as e:
                logger.error("Cannot combine")
                status_column_C = 'Cannot combine'

        else:
            logger.error("Cannot combine, cannot find bill or cont {}".format(cont_id))
            status_column_C = 'Cannot find bill or cont'

        merger.close()

        return status_column_C

    def _find_pdf_recursive(self, root_folder: str, prefix: str) -> Optional[str]:
        """
        Tìm file PDF đầu tiên (đệ quy) có tên bắt đầu bằng `prefix` trong `root_folder`.
        Trả về đường dẫn đầy đủ hoặc None nếu không tìm thấy.
        """
        logger: Logger = get_current_logger()
        for dirpath, _, filenames in os.walk(root_folder):
            for f in filenames:
                if f.lower().endswith('.pdf') and f.startswith(prefix):
                    full_path = os.path.join(dirpath, f)
                    return full_path
        return None

    def _safe_append_pdf(self, merger: PdfMerger, file_path: str, file_type: str, logger: Logger) -> bool:
        """
        """
        try:
            if not os.path.exists(file_path):
                logger.error("File not exits")
                return False
            if os.path.getsize(file_path) == 0:
                logger.error("File error")
                return False

            merger.append(file_path)
            logger.info("Added {}".format(file_path))
            return True

        except Exception as e:
            logger.error("Error when reading file")
            return False

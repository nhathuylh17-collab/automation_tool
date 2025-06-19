import os
from logging import Logger
from typing import Callable, Dict

from src.common.FileUtil import get_excel_data_in_column_start_at_row
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


class PDF_rename(AutomatedTask):
    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'folder_docs.folder', 'old_names', 'new_names',
                                     'file_type']
        return mandatory_keys

    def clear_worksheet(self, worksheet):
        excel_reader: ExcelReaderProvider = XlwingProvider()
        used_range = worksheet.used_range
        used_range.clear_contents()

    def read_excel_mapping(self, excel_reader: ExcelReaderProvider, workbook, sheet_name: str) -> Dict[str, str]:
        """
        Read Excel file to create a mapping from old PDF names (column A) to new PDF names (column B).

        Args:
            excel_reader: Instance of ExcelReaderProvider to read Excel data.
            workbook: The open Excel workbook.
            sheet_name: Name of the sheet to read.

        Returns:
            Dictionary mapping old file names to new file names.
        """
        logger: Logger = get_current_logger()
        worksheet = excel_reader.get_worksheet(workbook, sheet_name)

        # Read column A (old names) and column B (new names), starting from row 2 (assuming row 1 is header)
        old_names: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                     self._settings['excel.sheet'],
                                                                     self._settings['old_names'])
        new_names: list[str] = get_excel_data_in_column_start_at_row(self._settings['excel.path'],
                                                                     self._settings['excel.sheet'],
                                                                     self._settings['new_names'])
        file_type = self._settings['file_type']
        type = file_type.lower().lstrip('.')

        if not old_names or not new_names:
            logger.error("Excel file contains no valid data in columns A or B.")
            raise ValueError("Excel file contains no valid data in columns A or B.")

        if len(old_names) != len(new_names):
            logger.error("Mismatch between old and new name counts in Excel.")
            raise ValueError("Mismatch between old and new name counts in Excel.")

        # Ensure names have .pdf extension
        mapping = {}
        for old_name, new_name in zip(old_names, new_names):
            if old_name and new_name:  # Skip empty or None values
                old_name = str(old_name).strip()
                new_name = str(new_name).strip()
                # dang bi loi cho nay, chi rename khi maping, voi nhung file chua dc map tool loi
                old_name = old_name if old_name.lower().endswith(f".{type}") else f"{old_name}.{type}"
                new_name = new_name if new_name.lower().endswith(f".{type}") else f"{new_name}.{type}"
                mapping[old_name] = new_name

        logger.info(f"Loaded {len(mapping)} name mappings from Excel sheet: {sheet_name}")
        return mapping

    def automate(self):
        logger: Logger = get_current_logger()

        excel_reader: ExcelReaderProvider = XlwingProvider()

        path_to_excel_contain_pdfs_content = self._settings['excel.path']

        workbook = excel_reader.get_workbook(path=path_to_excel_contain_pdfs_content)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']

        # Read the name mapping from Excel
        try:
            name_mapping = self.read_excel_mapping(excel_reader, workbook, sheet_name)
        except Exception as e:
            logger.error(f"Failed to read Excel mapping: {str(e)}")
            excel_reader.close(workbook=workbook)
            excel_reader.quit_session()
            raise

        path_to_docs = self._settings['folder_docs.folder']

        renamed_count = 0

        for root, dirs, files in os.walk(path_to_docs):
            if self.terminated is True:
                return

            with self.pause_condition:

                while self.paused:
                    self.pause_condition.wait()

                if self.terminated is True:
                    return

            for current_pdf in files:
                if not current_pdf.lower().endswith(".pdf"):
                    continue

                current_pdf_path = os.path.join(root, current_pdf)

                if current_pdf in name_mapping:
                    new_file_name = name_mapping[current_pdf]
                    new_file_path = os.path.join(root, new_file_name)

                    # Check if new file name already exists
                    if os.path.exists(new_file_path):
                        logger.warning(f"New file name already exists, skipping: {new_file_name}")
                        continue

                    try:
                        os.rename(current_pdf_path, new_file_path)
                        logger.info(f"Renamed {current_pdf} to {new_file_name}")
                        renamed_count += 1
                    except Exception as e:
                        logger.error(f"Failed to rename {current_pdf} to {new_file_name}: {str(e)}")
                else:
                    logger.warning(f"No mapping found for file: {current_pdf}")

        logger.info(f"Renaming complete. Successfully renamed {renamed_count} files.")

        excel_reader.close(workbook=workbook)
        excel_reader.quit_session()
        logger.info('Closed excel file - Done')

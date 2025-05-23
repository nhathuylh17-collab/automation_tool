import os
from logging import Logger

import xlwings as xw
from xlwings import App

from src.common.ProcessUtil import kill_processes
from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider.ExcelReaderProvider import ExcelReaderProvider


class XlwingProvider(ExcelReaderProvider):

    def __init__(self):
        self._app: App = xw.App(visible=False)
        self.name_to_workbook: dict[str, object] = {}

    def get_workbook(self, path: str):
        # Open an existing workbook
        if self.name_to_workbook.get(path):
            return self.name_to_workbook.get(path)

        wb = self._app.books.open(path)
        self.name_to_workbook[path] = wb
        return wb

    def get_worksheet(self, workbook, sheet_name: str):
        ws = workbook.sheets[sheet_name]
        return ws

    def change_value_at(self, worksheet, row, column, value):
        worksheet.range(row, column).value = value
        return True

    def get_value_at(self, worksheet, row, column):
        return worksheet.range((row, column)).value

    def delete_contents(self, worksheet, start_cell, end_cell):
        worksheet.range(start_cell + ":" + end_cell).clear_contents()
        return True

    def save(self, workbook):
        logger: Logger = get_current_logger()

        try:

            workbook.save()
            logger.info(f'Save the workbook successfully')
            return workbook

        except BaseException as e:
            logger.error(f'Can not save the workbook, exception details: {e}')
            kill_processes('excel')

            current_path = workbook.fullname
            temp_path = current_path + 'temp'

            workbook.save(temp_path)
            os.remove(current_path)
            os.rename(temp_path, current_path)
            return self.get_workbook(path=current_path)

    def close(self, workbook):
        path_to_workbook: str = workbook.fullname
        workbook.close()
        if self.name_to_workbook.get(path_to_workbook) is not None:
            del self.name_to_workbook[path_to_workbook]

    def quit_session(self):
        self._app.quit()

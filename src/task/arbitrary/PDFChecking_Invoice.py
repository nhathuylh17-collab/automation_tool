import os
from logging import Logger
from typing import Callable

import pdfplumber
from pdfplumber import PDF

from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


class PDFChecking_Invoice(AutomatedTask):
    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'folder_docs.folder']
        return mandatory_keys

    def clear_worksheet(self, worksheet):
        excel_reader: ExcelReaderProvider = XlwingProvider()
        used_range = worksheet.used_range
        used_range.clear_contents()

    def automate(self):
        logger: Logger = get_current_logger()

        excel_reader: ExcelReaderProvider = XlwingProvider()

        path_to_excel_contain_pdfs_content = self._settings['excel.path']
        workbook = excel_reader.get_workbook(path=path_to_excel_contain_pdfs_content)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        worksheet = excel_reader.get_worksheet(workbook, sheet_name)
        self.clear_worksheet(worksheet)

        path_to_docs = self._settings['folder_docs.folder']
        pdf_counter: int = 2

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

                try:
                    pdf: PDF = pdfplumber.open(current_pdf_path)
                except Exception as e:
                    logger.error(f"Failed to open {current_pdf_path}: {e}")
                    continue

                logger.info("File name : {} PDF counter  = {}".format(current_pdf, pdf_counter))
                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter, column=1, value=current_pdf)

                current_page_in_current_pdf = 2

                for number, pageText in enumerate(pdf.pages):
                    raw_text = pageText.extract_text()
                    clean_text = raw_text.replace("\x00", "").replace("=", "")
                    line: str
                    PDF_lengh: list[str] = clean_text.splitlines()

                    runner: int = 0
                    while runner < len(PDF_lengh):
                        current = PDF_lengh[runner]

                        # case 1: Yusen
                        row_yusen = 0
                        yusen_HDGTGT = "HÓA ĐƠN GIÁ TRỊ GIA TĂNG Ký hiệu"
                        yusent_VAT = "VAT INVOICE Số HĐ"

                        if PDF_lengh[row_yusen].startswith("CHI NHÁNH CÔNG TY TNHH YUSEN LOGISTICS"):
                            def extract_after_colon(line: str) -> str:
                                if ':' in line:
                                    return line.split(':', 1)[1].strip()
                                return ""

                            if (runner + 1 < len(PDF_lengh) and
                                    current.startswith(yusen_HDGTGT) and
                                    PDF_lengh[runner + 1].startswith(yusent_VAT)):

                                logger.info('Processing for Yusen - {}'.format(current_pdf))

                                part1 = extract_after_colon(current)
                                part2 = extract_after_colon(PDF_lengh[runner + 1])
                                if part1 and part2:
                                    invoice_code = f"{part1}/{part2}"
                                    logger.info("found: {}".format(invoice_code))
                                    excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter,
                                                                 column=current_page_in_current_pdf, value=invoice_code)

                            if (runner + 1 < len(PDF_lengh) and
                                    PDF_lengh[runner + 1].startswith("Tổng cộng (Total amount):")):
                                bill_number = PDF_lengh[runner]

                                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter,
                                                             column=current_page_in_current_pdf + 1,
                                                             value="YASV/" and bill_number)

                                current_page_in_current_pdf += 1

                        # case 2: OVERSEAS
                        row_oversea = 1
                        overseaexpress_HDGTGT = "OVERSEAS EXPRESS CONSOLIDATORS"
                        overseaexpress_VAT = "HÓA ĐƠN GIÁ TRỊ GIA TĂNG"

                        if PDF_lengh[row_oversea].startswith("OVERSEAS EXPRESS"):
                            if (runner + 5 < len(PDF_lengh) and
                                    current.startswith(overseaexpress_HDGTGT) and
                                    PDF_lengh[runner + 5].startswith(overseaexpress_VAT)):

                                logger.info('Processing for OVERSEAS EXPRESS - {}'.format(current_pdf))

                                def extract_code(line: str) -> str:
                                    line = line.strip()

                                    if "Ký hiệu (Serial No.):" in line:
                                        return line.split("Ký hiệu (Serial No.):", 1)[1].strip()

                                    if line.startswith(overseaexpress_VAT):
                                        after_prefix = line[len(overseaexpress_VAT):].strip()
                                        if after_prefix:
                                            return after_prefix
                                        if ':' in line:
                                            return line.split(':', 1)[1].strip()
                                    if ':' in line:
                                        return line.split(':', 1)[1].strip()

                                    return ""

                                part1 = extract_code(current)
                                part2 = extract_code(PDF_lengh[runner + 5])

                                if part1 and part2:
                                    invoice_code = f"{part1}/{part2}"
                                    logger.info("found: {}".format(invoice_code))
                                    excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter,
                                                                 column=current_page_in_current_pdf, value=invoice_code)
                                current_page_in_current_pdf += 1

                        # case 3: Super speed
                        row_superspeed = 1
                        SUPERSPEED_HDGTGT = "HÓA ĐƠN GIÁ TRỊ GIA TĂNG Ký hiệu"
                        SUPERSPEED_VAT = "(VAT Invoice) Số"

                        if PDF_lengh[row_superspeed].startswith("SUPER SPEED LOGISTICS JOINT STOCK COMPANY"):
                            if (runner + 1 < len(PDF_lengh) and
                                    current.startswith(SUPERSPEED_HDGTGT) and
                                    PDF_lengh[runner + 1].startswith(SUPERSPEED_VAT)):

                                logger.info('Processing for Supper Speed - {}'.format(current_pdf))

                                def extract_code_super_speed(line: str) -> str:
                                    line = line.strip()
                                    if ':' in line:
                                        return line.split(':', 1)[1].strip()

                                    return ""

                                part1 = extract_code_super_speed(current)
                                part2 = extract_code_super_speed(PDF_lengh[runner + 1])

                                if part1 and part2:
                                    invoice_code = f"{part1}/{part2}"
                                    logger.info("found: {}".format(invoice_code))
                                    excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter,
                                                                 column=current_page_in_current_pdf, value=invoice_code)
                                current_page_in_current_pdf += 1

                            # runner = runner + 1

                        # case 4: FDI
                        row_fdi = 1
                        FDI_HDGTGT = "Ký hiệu (Serial)"
                        FDI_VAT = "(VAT INVOICE) Số (No)"

                        if PDF_lengh[row_fdi].startswith("GIAO NHẬN HÀNG HÓA F.D.I"):
                            if (runner + 1 < len(PDF_lengh) and
                                    current.startswith(FDI_HDGTGT) and
                                    PDF_lengh[runner + 1].startswith(FDI_VAT)):

                                logger.info('Processing for  F.D.I - {}'.format(current_pdf))

                                def extract_code_fdi(line: str) -> str:
                                    line = line.strip()
                                    if ':' in line:
                                        return line.split(':', 1)[1].strip()

                                    return ""

                                part1 = extract_code_fdi(current)
                                part2 = extract_code_fdi(PDF_lengh[runner + 1])

                                if part1 and part2:
                                    invoice_code = f"{part1}/{part2}"
                                    logger.info("found: {}".format(invoice_code))
                                    excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter,
                                                                 column=current_page_in_current_pdf, value=invoice_code)
                                current_page_in_current_pdf += 1

                        # case 5: WAN HAI (VIETNAM) LTD
                        row_wanhai = 1
                        Wanhai_HDGTGT = "HÓA ĐƠN GIÁ TRỊ GIA TĂNG Ký hiệu (Serial)"
                        Wanhai_VAT = "(VAT INVOICE) Số (Number)"

                        if PDF_lengh[row_wanhai].startswith("WAN HAI (VIETNAM) LTD"):
                            if (runner + 1 < len(PDF_lengh) and
                                    current.startswith(Wanhai_HDGTGT) and
                                    PDF_lengh[runner + 1].startswith(Wanhai_VAT)):

                                logger.info('Processing for Supper Speed - {}'.format(current_pdf))

                                def extract_code_super_speed(line: str) -> str:
                                    line = line.strip()
                                    if ':' in line:
                                        return line.split(':', 1)[1].strip()

                                    return ""

                                part1 = extract_code_super_speed(current)
                                part2 = extract_code_super_speed(PDF_lengh[runner + 1])

                                if part1 and part2:
                                    invoice_code = f"{part1}/{part2}"
                                    logger.info("found: {}".format(invoice_code))
                                    excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter,
                                                                 column=current_page_in_current_pdf,
                                                                 value=invoice_code)
                                current_page_in_current_pdf += 1
                        runner = runner + 1

                excel_reader.save(workbook=workbook)
                pdf_counter += 1

        excel_reader.close(workbook=workbook)
        excel_reader.quit_session()

        logger.info('Closed excel file - We done')

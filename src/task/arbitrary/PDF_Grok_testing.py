import os
from logging import Logger
from typing import Callable

import pdfplumber
from pdfplumber import PDF

from src.common.ThreadLocalLogger import get_current_logger
from src.excel_reader_provider import ExcelReaderProvider
from src.excel_reader_provider.XlwingProvider import XlwingProvider
from src.task.AutomatedTask import AutomatedTask


class PDF_Grok_testing(AutomatedTask):
    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['excel.path', 'excel.sheet', 'folder_docs.folder']
        return mandatory_keys

    def clear_worksheet(self, worksheet):
        excel_reader: ExcelReaderProvider = XlwingProvider()
        used_range = worksheet.used_range
        used_range.clear_contents()

    def extract_specific_fields(self, text: str) -> dict:
        """Trích xuất Consignee, Shipper, FCR Number, GW, CBM dựa trên tiêu đề và loại bỏ dữ liệu không mong muốn."""
        result = {
            "Shipper": None,
            "Consignee": None,
            "FCR Number": None,
            "Gross Weight": None,
            "CBM": None
        }

        lines = text.splitlines()
        full_text = text.lower()

        # Tìm Shipper (lấy toàn bộ đến "INCHEON, KOREA" hoặc trước "PORT AND COUNTRY OF ORIGIN")
        shipper_lines = []
        for i, line in enumerate(lines):
            if "shipper" in line.lower():
                start = i + 1
                while start < len(lines):
                    current_line = lines[start].strip()
                    current_line_lower = current_line.lower()
                    # Dừng nếu gặp "port and country", "consignee", "notify party", hoặc "receipt no"
                    if any(keyword in current_line_lower for keyword in
                           ["port and country", "consignee", "notify party", "receipt no"]):
                        break
                    if current_line and "r e c e i p t n o" not in current_line_lower:  # Loại bỏ "R E C E IP T N O"
                        shipper_lines.append(current_line)
                        # Kiểm tra nếu gặp "incheon, korea" để lấy hết thông tin
                        if "incheon, korea" in current_line_lower:
                            break
                    start += 1
                break
        if shipper_lines:
            # Loại bỏ dòng trống và ghép lại, giới hạn đến trước "PORT AND COUNTRY OF ORIGIN"
            result["Shipper"] = "\n".join(line for line in shipper_lines if line).strip()

        # Tìm Consignee (lấy toàn bộ đến "ZIP CODE: 34396**" hoặc trước "NOTIFY PARTY")
        consignee_lines = []
        for i, line in enumerate(lines):
            if "consignee" in line.lower():
                start = i + 1
                while start < len(lines):
                    current_line = lines[start].strip()
                    current_line_lower = current_line.lower()
                    # Dừng nếu gặp "notify party", "vessel", "maersk", hoặc "forwarder"
                    if any(keyword in current_line_lower for keyword in
                           ["notify party", "vessel", "maersk", "forwarder"]):
                        break
                    if current_line:
                        consignee_lines.append(current_line)
                    # Kiểm tra nếu có "zip code" để lấy hết thông tin
                    if "zip code" in current_line_lower:
                        start += 1
                        while start < len(lines) and lines[start].strip() and "notify party" not in lines[
                            start].lower():
                            if lines[start].strip() and "logistics & services" not in lines[start].lower():
                                consignee_lines.append(lines[start].strip())
                            start += 1
                        break
                    start += 1
                break
        if consignee_lines:
            # Loại bỏ "Logistics & Services" nếu lẫn vào
            result["Consignee"] = "\n".join(
                line for line in consignee_lines if line and "logistics & services" not in line.lower()).strip()

        # Tìm FCR Number
        for line in lines:
            if "receipt no." in line.lower():
                fcr = line.split("Receipt No.:")[-1].strip()
                if "squ" in fcr.lower():
                    fcr = fcr.replace("SQu", "SGN").replace("squ", "SGN")
                result["FCR Number"] = fcr
                break

        # Tìm Gross Weight và CBM
        for line in lines:
            if ("total:" in line.lower() and "cartons" in line.lower()) or (
                    "g.t.s" in line.lower() and "kgs" in line.lower()):
                parts = line.split()
                for j, part in enumerate(parts):
                    if ("cartons" in parts[j].lower() or "g.t.s" in parts[j].lower()) and j + 2 < len(parts):
                        try:
                            # Kiểm tra GW (Kgs) và CBM
                            if parts[j + 1].replace(".", "").isdigit() and parts[j + 2].replace(".", "").isdigit():
                                gw = parts[j + 1]
                                cbm = parts[j + 2]
                                result["Gross Weight"] = f"{gw} Kgs"
                                result["CBM"] = f"{cbm} Cbm"
                                break
                        except ValueError:
                            continue
                if result["Gross Weight"] and result["CBM"]:
                    break

        return result

    def automate(self):
        logger: Logger = get_current_logger()

        excel_reader: ExcelReaderProvider = XlwingProvider()

        path_to_excel_contain_pdfs_content = self._settings['excel.path']
        workbook = excel_reader.get_workbook(path=path_to_excel_contain_pdfs_content)
        logger.info('Loading excel files')

        sheet_name: str = self._settings['excel.sheet']
        worksheet = excel_reader.get_worksheet(workbook, sheet_name)
        self.clear_worksheet(worksheet)

        # Ghi tiêu đề cột
        headers = ["File Name", "FCR Number", "Shipper", "Consignee", "Gross Weight", "CBM"]
        for col, header in enumerate(headers, start=1):
            excel_reader.change_value_at(worksheet=worksheet, row=1, column=col, value=header)

        path_to_docs = self._settings['folder_docs.folder']
        pdf_counter: int = 1

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

                logger.info(f"File name: {current_pdf} PDF counter = {pdf_counter}")

                # Ghi tên file
                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter + 1, column=1, value=current_pdf)

                # Trích xuất toàn bộ text từ PDF
                full_text = ""
                for page in pdf.pages:
                    raw_text = page.extract_text()
                    if raw_text:
                        full_text += raw_text + "\n"

                # Lấy thông tin cụ thể
                extracted_data = self.extract_specific_fields(full_text)

                # Ghi dữ liệu vào Excel
                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter + 1, column=2,
                                             value=extracted_data["FCR Number"])
                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter + 1, column=3,
                                             value=extracted_data["Shipper"])
                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter + 1, column=4,
                                             value=extracted_data["Consignee"])
                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter + 1, column=5,
                                             value=extracted_data["Gross Weight"])
                excel_reader.change_value_at(worksheet=worksheet, row=pdf_counter + 1, column=6,
                                             value=extracted_data["CBM"])

                pdf_counter += 1

        excel_reader.save(workbook=workbook)
        excel_reader.close(workbook=workbook)
        excel_reader.quit_session()
        logger.info('Closed excel file - Done')

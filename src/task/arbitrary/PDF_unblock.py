import os
from concurrent.futures import ThreadPoolExecutor
from logging import Logger
from typing import Callable, Tuple, List

from pikepdf import Pdf, PasswordError, PdfError

from src.common.ThreadLocalLogger import get_current_logger
from src.task.AutomatedTask import AutomatedTask


class PDF_unblock(AutomatedTask):
    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self.logger: Logger = get_current_logger()
        # Danh sách mật khẩu cơ bản + phức tạp
        self.common_passwords: List[str] = [
            '', '123456', 'password', 'admin', '1234', 'test',
            'Password123!', 'Admin@2023', 'welcome123', 'user@pass',
            'SecurePass2025!', 'Test123456!', 'admin2025', 'pass@word1',
            'MyLongPassword@123', 'TestPassword!2025', 'Secure@Pass123',
            'AdminPassword!2025', 'Welcome@2025', 'User123!Pass'
        ]
        # Số lần thử lại tối đa cho mỗi file
        self.max_retries = 20
        # Đường dẫn đến wordlist (nếu có)
        self.wordlist_path = 'wordlist.txt'

        # Nạp thêm mật khẩu từ wordlist nếu tồn tại
        if os.path.exists(self.wordlist_path):
            with open(self.wordlist_path, 'r', encoding='utf-8') as f:
                self.common_passwords.extend(line.strip() for line in f if line.strip())
            self.logger.info(
                f"Loaded {len(self.common_passwords) - len(['', '123456', 'password', 'admin', '1234', 'test'])} additional passwords from wordlist.")

    def mandatory_settings(self) -> list[str]:
        return ['folder_docs.folder']

    def _decrypt_pdf(self, input_path: str, output_path: str) -> Tuple[bool, str]:
        """Attempt to decrypt a single PDF file with pikepdf, retry up to max_retries times."""
        attempt = 0
        while attempt < self.max_retries:
            attempt += 1
            try:
                # Thử mở file mà không cần mật khẩu trước
                with Pdf.open(input_path) as pdf:
                    if not pdf.is_encrypted:
                        pdf.save(output_path)
                        return True, f"Copied unencrypted {os.path.basename(input_path)}"
                    # Thử từng mật khẩu
                    for password in self.common_passwords:
                        try:
                            with Pdf.open(input_path, password=password) as pdf:
                                pdf.save(output_path)
                                return True, f"Decrypted {os.path.basename(input_path)} with password '{password}' after {attempt} attempts"
                        except PasswordError:
                            continue
                    return False, f"Failed to decrypt {os.path.basename(input_path)} - no matching password after {attempt} attempts"
            except PdfError as e:
                if attempt == self.max_retries:
                    return False, f"PDF error for {os.path.basename(input_path)} after {attempt} attempts: {str(e)}"
                self.logger.warning(
                    f"Attempt {attempt}/{self.max_retries} failed for {os.path.basename(input_path)}: {str(e)}. Retrying...")
            except Exception as e:
                if attempt == self.max_retries:
                    return False, f"Unexpected error for {os.path.basename(input_path)} after {attempt} attempts: {str(e)}"
                self.logger.warning(
                    f"Attempt {attempt}/{self.max_retries} failed for {os.path.basename(input_path)}: {str(e)}. Retrying...")
        return False, f"Failed to decrypt {os.path.basename(input_path)} after {self.max_retries} attempts"

    def automate(self):
        folder_path = self._settings['folder_docs.folder']
        self.logger.info(f"Starting PDF unblock process for folder: {folder_path}")

        output_folder = os.path.join(folder_path, "unlocked_pdfs")
        os.makedirs(output_folder, exist_ok=True)
        self.logger.info(f"Created output folder: {output_folder}")

        # Thu thập danh sách file PDF
        pdf_files = [
            (os.path.join(folder_path, filename),
             os.path.join(output_folder, f"unlocked_{filename}"))
            for filename in os.listdir(folder_path)
            if filename.lower().endswith('.pdf')
        ]

        if not pdf_files:
            self.logger.info("No PDF files found in the folder.")
            return

        self.logger.info(f"Found {len(pdf_files)} PDF files to process. Loaded {len(self.common_passwords)} passwords.")

        # Xử lý song song với ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=4) as executor:  # Điều chỉnh số workers tùy theo CPU
            futures = {executor.submit(self._decrypt_pdf, input_path, output_path): (input_path, output_path)
                       for input_path, output_path in pdf_files}

            # Xử lý kết quả khi hoàn thành
            for future in futures:
                input_path, output_path = futures[future]
                success, message = future.result()
                if success:
                    self.logger.info(message)
                else:
                    self.logger.warning(message)

        self.logger.info("PDF unblock process completed")

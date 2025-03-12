import logging

from PyQt5.QtWidgets import QTextEdit

from src.common.ThreadLocalLogger import get_current_logger


class TextBoxLoggingHandler(logging.Handler):
    def __init__(self, textbox: QTextEdit):
        super().__init__()
        self.textbox = textbox

    def emit(self, record):
        msg = self.format(record)
        # Đảm bảo thread-safe khi cập nhật UI
        self.textbox.append(msg)  # QTextEdit có phương thức append để thêm text
        self.textbox.verticalScrollBar().setValue(
            self.textbox.verticalScrollBar().maximum())  # Tự động scroll xuống cuối


class CustomLogFormatter(logging.Formatter):
    def format(self, record):
        # Lấy thời gian (chỉ giờ:phút:giây), loại bỏ mili giây
        time = record.asctime.split(',')[0]  # Lấy phần "YYYY-MM-DD HH:MM:SS" rồi tách HH:MM:SS
        time_part = time.split()[-1]  # Lấy phần "HH:MM:SS"
        module = record.module  # Tên module (ví dụ: "AutomatedTask")
        message = record.message  # Thông điệp log (ví dụ: "Run in headless mode")
        return f"{time_part}   :   {message}"


def setup_textbox_logger(textbox: QTextEdit):
    thread_local_logger: logging.Logger = get_current_logger()
    logging_handler: TextBoxLoggingHandler = TextBoxLoggingHandler(textbox)
    formatter: CustomLogFormatter = CustomLogFormatter()  # Sử dụng formatter tùy chỉnh
    logging_handler.setFormatter(formatter)
    thread_local_logger.addHandler(logging_handler)

import os
import sys
from typing import Dict, Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QTextEdit, QProgressBar, QFrame, QMessageBox, QScrollArea, QSplitter)

# Import các module cần thiết từ dự án
from src.common.FileUtil import load_key_value_from_file_properties, persist_settings_to_file
from src.common.ReflectionUtil import create_task_instance
from src.common.ThreadLocalLogger import get_current_logger
from src.gui.TextBoxLoggingHandler import setup_textbox_logger
from src.gui.UITaskPerformingStates import UITaskPerformingStates
from src.observer.Event import Event
from src.observer.EventBroker import EventBroker
from src.observer.EventHandler import EventHandler
from src.observer.PercentChangedEvent import PercentChangedEvent
from src.setup.packaging.path.PathResolvingService import PathResolvingService
from src.task.AutomatedTask import AutomatedTask


# Class để quản lý thread cho các task, chạy task trong background
class TaskThread(QThread):
    progress_updated = pyqtSignal(float, str)  # Tín hiệu cập nhật tiến trình và tên task

    def __init__(self, task: AutomatedTask, logger_setup):
        super().__init__()
        self.task = task
        self.logger_setup = logger_setup

    def run(self):
        if self.task:
            self.task.start()
            self.logger_setup()  # Cấu hình logger cho task


# Class triển khai EventHandler để xử lý sự kiện, sử dụng composition
class EventHandlerImpl(EventHandler):
    def __init__(self, gui_app=None):
        self.gui_app = gui_app  # Tham chiếu đến GUIApp để delegate logic xử lý sự kiện

    def handle_incoming_event(self, event: Event) -> None:
        if isinstance(event, PercentChangedEvent):
            if self.gui_app and self.gui_app.automated_task is None:
                self.gui_app.logger.error(
                    f'The PercentChangedEvent for {event.task_name} but no task in action in GUI app')
                return

            current_task_name = (type(self.gui_app.automated_task).__name__
                                 if self.gui_app and self.gui_app.automated_task else None)
            if current_task_name and event.task_name != current_task_name:
                self.gui_app.logger.warning(
                    f'The PercentChangedEvent for {event.task_name} is not match with the current task {current_task_name}')
                return

            if self.gui_app:
                self.gui_app.progress_bar.setValue(round(event.current_percent))
                self.gui_app.progress_bar.setFormat(f"{current_task_name} {event.current_percent}%")


# Class triển khai UITaskPerformingStates để quản lý trạng thái task, sử dụng composition
class UITaskPerformingStatesImpl(UITaskPerformingStates):
    def __init__(self, gui_app=None):
        self.gui_app = gui_app  # Tham chiếu đến GUIApp để delegate các phương thức trạng thái

    def get_ui_settings(self) -> Dict[str, str]:
        return self.gui_app.get_ui_settings() if self.gui_app else {}

    def set_ui_settings(self, new_ui_setting_values: Dict[str, str]) -> Dict[str, str]:
        return (self.gui_app.set_ui_settings(new_ui_setting_values)
                if self.gui_app else {})

    def get_task_name(self) -> str:
        return self.gui_app.get_task_name() if self.gui_app else ""

    def get_task_instance(self) -> Optional[AutomatedTask]:
        return self.gui_app.get_task_instance() if self.gui_app else None


# Class chính cho giao diện GUI, sử dụng PyQt5 với composition thay vì inheritance
class GUIApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # Sử dụng composition cho EventHandler và UITaskPerformingStates
        self.event_handler = EventHandlerImpl(self)  # Instance để xử lý sự kiện
        self.ui_task_states = UITaskPerformingStatesImpl(self)  # Instance để quản lý trạng thái task

        # Đăng ký làm observer cho sự kiện PercentChangedEvent
        EventBroker.get_instance().subscribe(topic=PercentChangedEvent.event_name, observer=self.event_handler)

        # Khởi tạo logger
        self.logger = get_current_logger()

        # Khởi tạo trạng thái task
        self.automated_task = None
        self.current_task_settings: Dict[str, str] = {}
        self.current_task_name = None
        self.is_task_currently_pause = False
        self.task_buttons = {}  # Lưu tất cả các nút task theo menu, dạng {menu: [task_buttons]}

        # Định nghĩa sidebar_menus như thuộc tính của class
        self.sidebar_menus = ["HomePage", "Website", "Desktop App", "Arbitrary"]

        # Cấu hình cửa sổ chính
        self.setWindowTitle("Maersk GSC VN Automation Toolkit")
        # Đặt kích thước cố định cho cửa sổ (1200x800, không mở rộng hết màn hình)
        self.resize(1200, 800)
        # self.setMinimumSize(1200, 800)  # Đặt kích thước tối thiểu để giữ cố định
        # self.setMaximumSize(1200, 800)  # Đặt kích thước tối đa để không mở rộng

        # Căn giữa cửa sổ trên màn hình
        screen = QApplication.desktop().screenGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

        # Xóa border viền xung quanh cửa sổ (không có frame viền)
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)

        # Tạo widget trung tâm và layout chính
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Header
        header = QFrame()
        header.setStyleSheet("background-color: #003E62; height: 50px; min-height: 50px; max-height: 50px;")
        header_layout = QHBoxLayout(header)
        header_layout.setAlignment(Qt.AlignLeft)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(0)

        resource_dir = PathResolvingService.get_instance().resolve('resource')
        logo_path = os.path.join(resource_dir, "img", "logo.png")
        if not os.path.exists(logo_path):
            logo_label = QLabel("Logo not found")
        else:
            logo_label = QLabel()
            pixmap = QPixmap(logo_path)
            logo_label.setPixmap(pixmap.scaled(200, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            logo_label.setMaximumHeight(60)
            logo_label.setMinimumHeight(60)
        header_layout.addWidget(logo_label)

        # Thêm spacer để đẩy các nút điều khiển về bên phải
        header_layout.addStretch()

        button_frame = QFrame()  # Tạo frame để chứa các nút, đảm bảo căn phải và đồng đều
        button_layout = QHBoxLayout(button_frame)
        button_layout.setSpacing(0)  # Không khoảng cách giữa các nút
        button_layout.setContentsMargins(0, 0, 0, 0)  # Không margin cho button_layout

        # Thêm các nút điều khiển (minimize, maximize, close) bên phải header, thiết kế giống hình ảnh
        minimize_button = QPushButton("-")
        minimize_button.setFont(QFont("Maersk Headline", 12, QFont.Bold))
        minimize_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF;
                padding: 0px 10px;
                border: none;
                border-radius: 0px;
                min-width: 30px;  /* Kích thước nút cố định */
                max-width: 30px;
            }
            QPushButton:hover {
                background-color: #42B0D5;
            }
            QPushButton:pressed {
                background-color: #1686BD;
            }
        """)
        minimize_button.clicked.connect(self.showMinimized)
        button_layout.addWidget(minimize_button)

        maximize_button = QPushButton("□")
        maximize_button.setFont(QFont("Maersk Headline", 12, QFont.Bold))
        maximize_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF; 
                padding: 0px 10px;
                border: none;
                border-radius: 0px;
                min-width: 30px;  /* Kích thước nút cố định */
                max-width: 30px;
            }
            QPushButton:hover {
                background-color: #42B0D5;
            }
            QPushButton:pressed {
                background-color: #1686BD;
            }
        """)
        maximize_button.clicked.connect(self.toggleMaximized)
        button_layout.addWidget(maximize_button)

        close_button = QPushButton("×")
        close_button.setFont(QFont("Maersk Headline", 12, QFont.Bold))  # Sử dụng font Arial, đậm, kích thước 12
        close_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF;  /* Màu đen cho icon */
                padding: 0px 10px;
                border: none;
                border-radius: 0px;
                min-width: 30px;  /* Kích thước nút cố định */
                max-width: 30px;
            }
            QPushButton:hover {
                background-color: #FF4444;  /* Màu đỏ khi hover, giống hình ảnh */
            }
            QPushButton:pressed {
                background-color: #CC3333;
            }
        """)
        close_button.clicked.connect(self.close)
        button_layout.addWidget(close_button)

        # Thêm frame chứa các nút vào header_layout (căn phải)
        header_layout.addWidget(button_frame)

        main_layout.addWidget(header)

        # Thêm khả năng di chuyển cửa sổ bằng cách kéo header
        self.header = header  # Lưu tham chiếu đến header
        self.drag_position = None

        # Sử dụng QSplitter để chia phần còn lại thành khu vực bên trái (3/10) và bên phải (7/10)
        splitter = QSplitter(Qt.Horizontal)

        # Khu vực bên trái (3/10 chiều rộng, chứa sidebar, cao 9/10 chiều cao)
        left_frame = QFrame()
        left_frame.setStyleSheet("background-color: #FFFFFF;")
        left_layout = QVBoxLayout(left_frame)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Sidebar với QScrollArea, thiết kế thanh cuộn dọc mảnh và hiện đại
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)  # Cho phép widget bên trong thay đổi kích thước
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: 0px;
            }
            QScrollArea QScrollBar:vertical {
                border: none;
                background: #F5F5F5;  /* Nền sáng, hiện đại */
                width: 6px;  /* Mỏng hơn, phù hợp với thiết kế hiện đại */
                margin: 0px 0px 0px 0px;
            }
            QScrollArea QScrollBar::handle:vertical {
                background: #C0C0C0;  /* Màu xám nhạt, giống hình mẫu */
                min-height: 20px;
                border-radius: 3px;  /* Bo góc mượt */
            }
            QScrollArea QScrollBar::handle:vertical:hover {
                background: #A9A9A9;  /* Màu tối hơn khi hover, tinh tế */
            }
            QScrollArea QScrollBar::add-line:vertical {
                height: 0px;  /* Loại bỏ nút xuống dưới */
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }
            QScrollArea QScrollBar::sub-line:vertical {
                height: 0px;  /* Loại bỏ nút lên trên */
                subcontrol-position: top;
                subcontrol-origin: margin;
            }
            QScrollArea QScrollBar::up-arrow:vertical, QScrollArea QScrollBar::down-arrow:vertical {
                background: none;
            }
            QScrollArea QScrollBar::add-page:vertical, QScrollArea QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # Vô hiệu hóa scrollbar ngang
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # Chỉ hiển thị scrollbar dọc khi cần
        self.sidebar = QFrame()
        self.sidebar.setStyleSheet("background-color: #FFFFFF; border: 0px;")
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setAlignment(Qt.AlignTop)
        self.sidebar_layout.setSpacing(5)  # Giảm khoảng cách giữa các nút cho nhìn gọn hơn
        scroll_area.setWidget(self.sidebar)
        left_layout.addWidget(scroll_area)

        # Thêm các nút menu vào sidebar, bám sát bên trái với style mới, đảm bảo text hiển thị full
        for menu in self.sidebar_menus:
            btn = QPushButton(menu)
            btn.setFont(QFont("Maersk Headline", 10))  # Giữ font size 10 để text fit tốt
            if menu in ["Website", "Desktop App", "Arbitrary"]:
                # Style cho Website, Desktop App, và Arbitrary, đảm bảo text không bị cắt
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #FFFFFF;
                        color: #6A6A6A;  /* Màu xám như trong ảnh */
                        padding: 8px 0 8px 10px;  /* Padding nhỏ hơn cho nhìn gọn */
                        border: none;
                        width: 100%;
                        text-align: left;
                        font-weight: normal;
                        min-width: 0;  /* Đảm bảo nút không cố định chiều rộng tối thiểu */
                        max-width: 100%;  /* Cho phép text mở rộng hết chiều rộng sidebar */
                        white-space: normal;  /* Cho phép text xuống dòng nếu cần */
                    }
                    QPushButton:hover {
                        background-color: #F5F5F5;  /* Nền nhạt khi hover */
                        color: #444444;  /* Màu xám đậm hơn khi hover */
                    }
                """)
            else:  # Style cho HomePage (giữ nguyên hoặc điều chỉnh nếu cần)
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #FFFFFF;
                        color: #003E62;
                        padding: 10px 0 10px 10px;
                        border: none;
                        width: 100%;
                        text-align: left;
                        font-weight: bold;  /* Giữ HomePage đậm hơn các nút khác */
                        min-width: 0;  /* Đảm bảo nút không cố định chiều rộng tối thiểu */
                        max-width: 100%;  /* Cho phép text mở rộng hết chiều rộng sidebar */
                        white-space: normal;  /* Cho phép text xuống dòng nếu cần */
                    }
                    QPushButton:hover {
                        color: #1686BD;
                        background-color: #E0E0E0;
                    }
                """)
            if menu == "HomePage":
                btn.clicked.connect(lambda checked, m=menu: self.handle_homepage(m))
            elif menu in ["Website", "Desktop App", "Arbitrary"]:
                btn.clicked.connect(lambda checked, m=menu: self.toggle_task_list(m, {"Website": "web",
                                                                                      "Desktop App": "desktop",
                                                                                      "Arbitrary": "arbitrary"}[m]))
            self.sidebar_layout.addWidget(btn)

        left_layout.addWidget(scroll_area)
        left_frame.setLayout(left_layout)

        # Khu vực bên phải (7/10 chiều rộng, chứa main content, cao 9/10 chiều cao)
        right_frame = QFrame()
        right_frame.setStyleSheet("background-color: #F0F0F0;")
        right_layout = QVBoxLayout(right_frame)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # Main content
        main_content = QFrame()
        main_content.setStyleSheet("background-color: #F0F0F0; border: 1px solid #D4D4D4; border-radius: 5px;")
        main_content_layout = QHBoxLayout(main_content)
        main_content_layout.setAlignment(Qt.AlignCenter)

        # Nội dung chính (settings frame với scroll bar, thiết kế thanh cuộn mảnh và hiện đại)
        content_frame = QFrame()
        content_frame.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D4D4D4; border-radius: 5px;")
        content_layout = QVBoxLayout(content_frame)

        # Khu vực hiển thị field với scroll bar
        scroll_area_content = QScrollArea()
        scroll_area_content.setWidgetResizable(True)
        scroll_area_content.setStyleSheet("""
            QScrollArea {
                border: 0px;
            }
            QScrollArea QScrollBar:vertical {
                border: none;
                background: #F5F5F5;  /* Nền sáng, hiện đại */
                width: 6px;  /* Mỏng hơn, phù hợp với thiết kế hiện đại */
                margin: 0px 0px 0px 0px;
            }
            QScrollArea QScrollBar::handle:vertical {
                background: #C0C0C0;  /* Màu xám nhạt, giống hình mẫu */
                min-height: 20px;
                border-radius: 3px;  /* Bo góc mượt */
            }
            QScrollArea QScrollBar::handle:vertical:hover {
                background: #A9A9A9;  /* Màu tối hơn khi hover, tinh tế */
            }
            QScrollArea QScrollBar::add-line:vertical {
                height: 0px;  /* Loại bỏ nút xuống dưới */
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }
            QScrollArea QScrollBar::sub-line:vertical {
                height: 0px;  /* Loại bỏ nút lên trên */
                subcontrol-position: top;
                subcontrol-origin: margin;
            }
            QScrollArea QScrollBar::up-arrow:vertical, QScrollArea QScrollBar::down-arrow:vertical {
                background: none;
            }
            QScrollArea QScrollBar::add-page:vertical, QScrollArea QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        scroll_area_content.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # Vô hiệu hóa scrollbar ngang
        scroll_area_content.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # Chỉ hiển thị scrollbar dọc khi cần
        self.settings_frame = QFrame()
        self.settings_frame.setStyleSheet(
            "background-color: #FFFFFF; border: 1px solid #D4D4D4; border-radius: 5px; padding: 10px;")
        self.settings_layout = QVBoxLayout(self.settings_frame)
        scroll_area_content.setWidget(self.settings_frame)
        content_layout.addWidget(scroll_area_content)

        main_content_layout.addWidget(content_frame)
        right_layout.addWidget(main_content)

        # Frame chứa các nút, progress bar, và textbox logger
        bottom_frame = QFrame()
        bottom_frame.setStyleSheet("background-color: #F0F0F0;")
        bottom_layout = QVBoxLayout(bottom_frame)

        # Frame chứa các nút
        self.button_frame = QFrame()
        self.button_frame.setStyleSheet("background-color: #F0F0F0;")
        button_layout = QHBoxLayout(self.button_frame)
        button_layout.setSpacing(10)

        self.perform_button = QPushButton("Perform")
        self.perform_button.setFont(QFont("Maersk Headline", 11))
        self.perform_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF;
                padding: 5px 15px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #42B0D5;
            }
            QPushButton:pressed {
                background-color: #1686BD;
            }
        """)
        self.perform_button.clicked.connect(self.handle_perform_button)
        button_layout.addWidget(self.perform_button)

        self.pause_button = QPushButton("Pause")
        self.pause_button.setFont(QFont("Maersk Headline", 11))
        self.pause_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF;
                padding: 5px 15px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #42B0D5;
            }
            QPushButton:pressed {
                background-color: #1686BD;
            }
        """)
        self.pause_button.clicked.connect(self.handle_pause_button)
        button_layout.addWidget(self.pause_button)

        self.reset_button = QPushButton("Reset")
        self.reset_button.setFont(QFont("Maersk Headline", 11))
        self.reset_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF;
                padding: 5px 15px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #EA5D4B;
            }
            QPushButton:pressed {
                background-color: #1686BD;
            }
        """)
        self.reset_button.clicked.connect(self.handle_reset_button)
        button_layout.addWidget(self.reset_button)

        bottom_layout.addWidget(self.button_frame)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setFont(QFont("Maersk Headline", 10))
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #D4D4D4;
                border: 1px solid #D4D4D4;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #42B0D5;
                border-radius: 5px;
            }
        """)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("None 0%")
        bottom_layout.addWidget(self.progress_bar)

        # Textbox logging với thanh cuộn dọc được thiết kế mảnh và hiện đại
        self.logging_textbox = QTextEdit()
        self.logging_textbox.setFont(QFont("Maersk Headline", 10))
        self.logging_textbox.setStyleSheet("""
            QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #D4D4D4;
                border-radius: 5px;
                padding: 5px;
                color: #363636;
                min-height: 200px;  # Tăng chiều cao tối thiểu
                min-width: 800px;   # Tăng chiều rộng tối thiểu
            }
            QTextEdit QScrollBar:vertical {
                border: none;
                background: #F5F5F5;  /* Nền sáng, hiện đại */
                width: 6px;  /* Mỏng hơn, phù hợp với thiết kế hiện đại */
                margin: 0px 0px 0px 0px;
            }
            QTextEdit QScrollBar::handle:vertical {
                background: #C0C0C0;  /* Màu xám nhạt, giống hình mẫu */
                min-height: 20px;
                border-radius: 3px;  /* Bo góc mượt */
            }
            QTextEdit QScrollBar::handle:vertical:hover {
                background: #A9A9A9;  /* Màu tối hơn khi hover, tinh tế */
            }
            QTextEdit QScrollBar::add-line:vertical {
                height: 0px;  /* Loại bỏ nút xuống dưới */
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }
            QTextEdit QScrollBar::sub-line:vertical {
                height: 0px;  /* Loại bỏ nút lên trên */
                subcontrol-position: top;
                subcontrol-origin: margin;
            }
            QTextEdit QScrollBar::up-arrow:vertical, QTextEdit QScrollBar::down-arrow:vertical {
                background: none;
            }
            QTextEdit QScrollBar::add-page:vertical, QTextEdit QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        self.logging_textbox.setReadOnly(True)
        self.logging_textbox.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # Hiển thị thanh cuộn khi cần
        bottom_layout.addWidget(self.logging_textbox)

        right_layout.addWidget(bottom_frame)

        # Cấu hình QSplitter
        splitter.addWidget(left_frame)
        splitter.addWidget(right_frame)
        splitter.setStretchFactor(0, 3)  # Khu vực bên trái chiếm 3/10 (30%)
        splitter.setStretchFactor(1, 7)  # Khu vực bên phải chiếm 7/10 (70%)
        splitter.setHandleWidth(1)
        splitter.setStyleSheet("QSplitter::handle { background-color: #D4D4D4; }")

        # Thêm splitter vào layout chính (dưới header)
        main_layout.addWidget(splitter)

        # Cấu hình logger cho textbox với formatter tùy chỉnh
        self.setup_custom_logger()

    def setup_custom_logger(self):
        """Cấu hình logger với formatter tùy chỉnh để hiển thị log theo định dạng: 'HH:MM:SS Module Message'."""
        setup_textbox_logger(self.logging_textbox)

    def get_ui_settings(self) -> Dict[str, str]:
        """Lấy các thiết lập giao diện hiện tại."""
        return self.current_task_settings

    def set_ui_settings(self, new_ui_setting_values: Dict[str, str]) -> Dict[str, str]:
        """Cập nhật các thiết lập giao diện."""
        self.current_task_settings = new_ui_setting_values
        return self.current_task_settings

    def get_task_name(self) -> str:
        """Lấy tên task hiện tại."""
        return self.current_task_name

    def get_task_instance(self) -> Optional[AutomatedTask]:
        """Lấy instance của task hiện tại."""
        return self.automated_task

    # Delegate phương thức xử lý sự kiện từ EventHandler
    def handle_incoming_event(self, event: Event) -> None:
        self.event_handler.handle_incoming_event(event)

    def closeEvent(self, event):
        """Xử lý khi cửa sổ chính được đóng."""
        persist_settings_to_file(self.current_task_name, self.current_task_settings)
        event.accept()

    def handle_homepage(self, menu: str):
        """Xử lý khi nhấn nút HomePage."""
        # Xóa nội dung cũ trong settings layout và ẩn các nút task con
        for i in reversed(range(self.settings_layout.count())):
            self.settings_layout.itemAt(i).widget().setParent(None)
        for menu_tasks in self.task_buttons.values():
            for btn in menu_tasks:
                btn.setVisible(False)
        # self.logger.info("Displaying HomePage.")
        self.current_task_name = None
        self.automated_task = None
        self.current_task_settings = {}

    def toggle_task_list(self, menu: str, dir_name: str):
        """Hiển thị hoặc ẩn danh sách các task từ thư mục riêng cho từng menu, giữ các menu khác nguyên, đảm bảo chỉ hiển thị tasks của thư mục tương ứng, và hỗ trợ toggle (hiển thị/ẩn) khi click lại nút."""
        # Debug để kiểm tra thư mục và file
        self.logger.debug(f"Toggling tasks for menu: {menu}, directory: {dir_name}")

        # Kiểm tra xem menu có task chưa, nếu có thì toggle hiển thị/ẩn
        if menu in self.task_buttons and self.task_buttons[menu]:
            are_tasks_visible = self.task_buttons[menu][0].isVisible() if self.task_buttons[menu] else False
            for btn in self.task_buttons[menu]:
                btn.setVisible(not are_tasks_visible)  # Toggle trạng thái hiển thị
            return

        # Xóa tất cả các task hiện có trong sidebar trước khi thêm mới, để tránh tích lũy hoặc nhầm lẫn
        for menu_tasks in self.task_buttons.values():
            for btn in menu_tasks:
                btn.setParent(None)  # Xóa hoàn toàn các nút cũ
        self.task_buttons.clear()  # Xóa toàn bộ task_buttons để bắt đầu lại

        # Đường dẫn đến thư mục task trong src/task/
        base_dir = os.path.dirname(__file__)
        task_dir = os.path.join(base_dir, '..', 'task', dir_name)
        self.logger.debug(f"Task directory path: {task_dir}")
        if not os.path.exists(task_dir):
            self.logger.error(f"Directory {task_dir} does not exist.")
            return

        # Lấy danh sách các file .py trong thư mục, loại bỏ __init__.py
        task_files = [f[:-3] for f in os.listdir(task_dir) if f.endswith('.py') and f != '__init__.py']
        self.logger.debug(f"Found task files: {task_files}")
        if not task_files:
            self.logger.info(f"No tasks found in {dir_name} directory.")
            return

        # Tạo và thêm các nút con cho từng task dưới nút menu tương ứng
        self.task_buttons[menu] = []  # Lưu danh sách các nút task cho menu này
        menu_index = self.sidebar_menus.index(menu)
        for task_name in task_files:
            task_btn = QPushButton(task_name)
            task_btn.setFont(QFont("Maersk Headline", 10))  # Giữ font size 10 để text fit tốt
            task_btn.setStyleSheet("""
                QPushButton {
                    background-color: #FFFFFF;
                    color: #6A6A6A;  /* Màu xám như nút menu chính */
                    padding: 5px 0 5px 20px;  /* Padding trái lớn hơn để thụt lề */
                    border: none;
                    width: 100%;
                    text-align: left;
                    font-weight: normal;
                    min-width: 0;  /* Đảm bảo nút không cố định chiều rộng tối thiểu */
                    max-width: 100%;  /* Cho phép text mở rộng hết chiều rộng sidebar */
                    white-space: normal;  /* Cho phép text xuống dòng nếu cần */
                }
                QPushButton:hover {
                    background-color: #F5F5F5;  /* Nền nhạt khi hover */
                    color: #444444;  /* Màu xám đậm hơn khi hover */
                }
            """)
            task_btn.clicked.connect(lambda checked, t=task_name: self.render_task_fields(t))
            task_btn.setVisible(True)  # Hiển thị nút task con ngay khi tạo
            self.task_buttons[menu].append(task_btn)
            self.sidebar_layout.insertWidget(menu_index + 1, task_btn)  # Thêm ngay dưới nút menu

        # Ẩn các task của các menu khác nếu chúng hiện đang hiển thị
        for other_menu in self.sidebar_menus:
            if other_menu != menu and other_menu in self.task_buttons:
                for btn in self.task_buttons[other_menu]:
                    btn.setVisible(False)

    def render_task_fields(self, task_name: str):
        """Hiển thị các field input cho task được chọn."""
        # Xóa nội dung cũ trong settings layout
        for i in reversed(range(self.settings_layout.count())):
            self.settings_layout.itemAt(i).widget().setParent(None)

        self.logger.info(f'Display fields for task {task_name}')

        # Tạo hoặc tải file settings cho task
        setting_file = os.path.join(PathResolvingService.get_instance().get_input_dir(),
                                    f'{task_name}.properties')
        if not os.path.exists(setting_file):
            with open(setting_file, 'w'):
                pass  # Tạo file mới nếu chưa tồn tại

        # Tải các giá trị thiết lập từ file
        input_setting_values: Dict[str, str] = load_key_value_from_file_properties(setting_file)
        input_setting_values['invoked_class'] = task_name

        # Thiết lập giá trị mặc định nếu chưa có
        if 'time.unit.factor' not in input_setting_values:
            input_setting_values['time.unit.factor'] = '1'
        if 'use.GUI' not in input_setting_values:
            input_setting_values['use.GUI'] = 'True'

        # Tạo instance của task
        self.automated_task = create_task_instance(input_setting_values, task_name,
                                                   lambda: self.setup_custom_logger())
        mandatory_settings = self.automated_task.mandatory_settings()
        mandatory_settings.extend(['invoked_class', 'time.unit.factor', 'use.GUI'])

        # Cập nhật thông tin task và settings
        self.current_task_name = task_name
        self.current_task_settings = {}
        for setting in mandatory_settings:
            initial_value = input_setting_values.get(setting, '')
            self.current_task_settings[setting] = initial_value

            # Tạo label và input cho mỗi thiết lập, loại bỏ hoàn toàn border để phẳng hơn
            setting_frame = QFrame()
            setting_frame.setStyleSheet("background-color: #FFFFFF;")  # Loại bỏ border, giữ nền trắng
            setting_layout = QHBoxLayout(setting_frame)
            setting_layout.setContentsMargins(0, 0, 0, 0)  # Loại bỏ margin để phẳng hơn
            setting_layout.setSpacing(5)  # Khoảng cách nhỏ giữa label và input

            label = QLabel(f"{setting}:")
            label.setFont(QFont("Maersk Headline", 10))
            label.setStyleSheet("color: #363636;")
            setting_layout.addWidget(label)

            from src.gui.UIComponentFactory import UIComponentFactory
            component = UIComponentFactory.get_instance(self).create_component(setting, initial_value, setting_frame)
            if component:
                # Loại bỏ viền của component (input) để phẳng hoàn toàn, giữ nền trắng
                component.setStyleSheet(
                    "border: none; background-color: #FFFFFF; padding: 5px;")  # Không viền, nền trắng, padding nhỏ
                setting_layout.addWidget(component)

            self.settings_layout.addWidget(setting_frame)

        self.automated_task.settings = self.current_task_settings

    def update_setting(self, setting: str, value: str):
        """Cập nhật giá trị của một thiết lập cụ thể."""
        self.current_task_settings[setting] = value

    def handle_perform_button(self):
        """Xử lý khi nút 'Perform' được nhấn."""
        if self.automated_task and self.automated_task.is_alive():
            QMessageBox.information(self, "Task Running", "Please terminate the current task before run a new one")
            return

        if not self.automated_task:
            self.automated_task = create_task_instance(self.current_task_settings, self.current_task_name,
                                                       lambda: self.setup_custom_logger())

        self.progress_bar.setValue(0)
        self.progress_bar.setFormat(f"{type(self.automated_task).__name__} 0%")

        # Chạy task trong thread riêng
        self.task_thread = TaskThread(self.automated_task, lambda: self.setup_custom_logger())
        self.task_thread.progress_updated.connect(self.update_progress)
        self.task_thread.start()

    def handle_pause_button(self):
        """Xử lý khi nút 'Pause' được nhấn."""
        if not self.automated_task or not self.automated_task.is_alive():
            return

        if self.is_task_currently_pause:
            self.automated_task.resume()
            self.pause_button.setText("Pause")
            self.is_task_currently_pause = False
        else:
            self.automated_task.pause()
            self.pause_button.setText("Resume")
            self.is_task_currently_pause = True

    def handle_reset_button(self):
        """Xử lý khi nút 'Reset' được nhấn."""
        # Kiểm tra nếu chưa có task nào được chạy hoặc chưa chọn task
        if not self.automated_task or not self.current_task_name:
            self.logger.info("No task is running or selected. Reset action skipped.")
            return  # Không làm gì nếu chưa có task hoặc chưa chọn task

        # Nếu có task nhưng chưa chạy (chỉ chọn task), reset về trạng thái ban đầu
        if self.automated_task and not self.automated_task.is_alive():
            self.logger.info(f"Resetting selected task: {self.current_task_name}")
            self.automated_task = None
            self.current_task_name = None
            self.current_task_settings = {}
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat("None 0%")
            if self.is_task_currently_pause:
                self.pause_button.setText("Pause")
                self.is_task_currently_pause = False
            return

        # Nếu task đang chạy, terminate nó
        if self.automated_task and self.automated_task.is_alive():
            self.logger.info(f"Terminating running task: {self.current_task_name}")
            self.automated_task.terminate()

        # Reset các trạng thái sau khi terminate
        if self.is_task_currently_pause:
            self.pause_button.setText("Pause")
            self.is_task_currently_pause = False

        self.automated_task = None
        self.current_task_name = None
        self.current_task_settings = {}
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("None 0%")

    def update_progress(self, percent: float, task_name: str):
        """Cập nhật tiến trình của task trên progress bar."""
        self.progress_bar.setValue(round(percent))
        self.progress_bar.setFormat(f"{task_name} {percent}%")

    def toggleMaximized(self):
        """Toggle giữa chế độ toàn màn hình và kích thước cố định 1200x800."""
        if self.isMaximized():
            self.showNormal()  # Khôi phục về kích thước ban đầu (1200x800)
        else:
            self.showMaximized()  # Chuyển sang chế độ toàn màn hình


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Maersk Headline", 10))
    window = GUIApp()
    window.show()
    sys.exit(app.exec_())

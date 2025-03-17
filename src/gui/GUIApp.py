import getpass
import os
import socket
import sys
from typing import Dict, Optional

import pythoncom
import win32com.client
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation, QEasingCurve, QRect
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QTextEdit, QProgressBar, QFrame, QMessageBox, QScrollArea, QSplitter, QCheckBox, QComboBox,
                             QGraphicsDropShadowEffect, QLineEdit, QTabWidget, QTreeWidgetItem, QTreeWidget, QStyle)

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


class TaskThread(QThread):
    progress_updated = pyqtSignal(float, str)

    def __init__(self, task: AutomatedTask, logger_setup):
        super().__init__()
        self.automated_task = task
        self.logger_setup = logger_setup
        self.com_initialized = False

    def run(self):
        try:
            pythoncom.CoInitialize()
            self.com_initialized = True
            if self.automated_task:
                self.automated_task.start()
                self.logger_setup()
        except Exception as e:
            self.logger.error(f"Error in task thread: {str(e)}")
        finally:
            if self.com_initialized:
                pythoncom.CoUninitialize()


class TaskFieldWorker(QThread):
    task_ready = pyqtSignal(str, object, object, object, str)
    error_occurred = pyqtSignal(str)

    def __init__(self, task_name: str, parent=None):
        super().__init__(parent)
        self.task_name = task_name
        self.com_initialized = False

    def run(self):
        try:
            if not self.com_initialized:
                try:
                    pythoncom.CoInitialize()
                    self.com_initialized = True
                except pythoncom.com_error:
                    pass

            setting_file = os.path.join(PathResolvingService.get_instance().get_input_dir(),
                                        f'{self.task_name}.properties')
            if not os.path.exists(setting_file):
                with open(setting_file, 'w'):
                    pass

            input_setting_values: Dict[str, str] = load_key_value_from_file_properties(setting_file)
            input_setting_values['invoked_class'] = self.task_name

            if 'time.unit.factor' not in input_setting_values:
                input_setting_values['time.unit.factor'] = '1'
            if 'use.GUI' not in input_setting_values:
                input_setting_values['use.GUI'] = 'False'

            automated_task = create_task_instance(input_setting_values, self.task_name,
                                                  lambda: self.parent().setup_custom_logger() if self.parent() else None)
            if automated_task is None:
                raise ValueError(f'Could not instantiate task {self.task_name}')

            mandatory_settings = automated_task.mandatory_settings()
            mandatory_settings.extend(['invoked_class', 'time.unit.factor', 'use.GUI'])

            task_menu = None
            if self.parent():
                for menu, task_buttons in self.parent().task_buttons.items():
                    for btn in task_buttons:
                        if btn.text() == self.task_name:
                            task_menu = menu
                            break
                    if task_menu:
                        break
                if not task_menu and self.task_name in self.parent().sidebar_menus:
                    task_menu = self.task_name

            current_task_settings = {}
            for setting in mandatory_settings:
                initial_value = input_setting_values.get(setting, '')
                current_task_settings[setting] = initial_value

            self.task_ready.emit(self.task_name, automated_task, current_task_settings, mandatory_settings, task_menu)

        except Exception as e:
            self.error_occurred.emit(f"Error preparing task {self.task_name}: {str(e)}")
        finally:
            if self.com_initialized:
                pythoncom.CoUninitialize()


class EventHandlerImpl(EventHandler):
    def __init__(self, gui_app=None):
        self.gui_app = gui_app

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


class UITaskPerformingStatesImpl(UITaskPerformingStates):
    def __init__(self, gui_app=None):
        self.gui_app = gui_app

    def get_ui_settings(self) -> Dict[str, str]:
        return self.gui_app.get_ui_settings() if self.gui_app else {}

    def set_ui_settings(self, new_ui_setting_values: Dict[str, str]) -> Dict[str, str]:
        return (self.gui_app.set_ui_settings(new_ui_setting_values)
                if self.gui_app else {})

    def get_task_name(self) -> str:
        return self.gui_app.get_task_name() if self.gui_app else ""

    def get_task_instance(self) -> Optional[AutomatedTask]:
        return self.gui_app.get_task_instance() if self.gui_app else None


class GUIApp(QMainWindow):
    def __init__(self):
        super().__init__()
        try:
            pythoncom.CoInitialize()
            self.com_initialized = True
        except pythoncom.com_error:
            self.com_initialized = False

        self.event_handler = EventHandlerImpl(self)
        self.ui_task_states = UITaskPerformingStatesImpl(self)
        EventBroker.get_instance().subscribe(topic=PercentChangedEvent.event_name, observer=self.event_handler)

        self.logger = get_current_logger()
        self.automated_task = None
        self.current_task_settings: Dict[str, str] = {}
        self.current_task_name = None
        self.is_task_currently_pause = False
        self.task_buttons = {}
        self.sidebar_menus = ["HomePage", "Website", "Desktop App", "Arbitrary", "Setting"]

        self.setWindowTitle("Maersk GSC VN Automation Toolkit")
        self.resize(1200, 800)

        screen = QApplication.desktop().screenGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

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

        header_layout.addStretch()

        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setSpacing(0)
        button_layout.setContentsMargins(0, 0, 0, 0)

        minimize_button = QPushButton("-")
        minimize_button.setFont(QFont("Maersk Headline", 12, QFont.Bold))
        minimize_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF;
                padding: 0px 10px;
                border: none;
                border-radius: 0px;
                min-width: 30px;
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
                min-width: 30px;
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
        close_button.setFont(QFont("Maersk Headline", 12, QFont.Bold))
        close_button.setStyleSheet("""
            QPushButton {
                background-color: #003E62;
                color: #FFFFFF;
                padding: 0px 10px;
                border: none;
                border-radius: 0px;
                min-width: 30px;
                max-width: 30px;
            }
            QPushButton:hover {
                background-color: #FF4444;
            }
            QPushButton:pressed {
                background-color: #CC3333;
            }
        """)
        close_button.clicked.connect(self.close)
        button_layout.addWidget(close_button)

        header_layout.addWidget(button_frame)
        main_layout.addWidget(header)

        self.header = header
        self.drag_position = None

        splitter = QSplitter(Qt.Horizontal)

        left_frame = QFrame()
        left_frame.setStyleSheet("background-color: #FFFFFF;")
        left_layout = QVBoxLayout(left_frame)
        left_layout.setContentsMargins(0, 0, 0, 0)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("""
            QScrollArea {
                border: 0px;
            }
            QScrollArea QScrollBar:vertical {
                border: none;
                background: #F5F5F5;
                width: 6px;
                margin: 0px 0px 0px 0px;
            }
            QScrollArea QScrollBar::handle:vertical {
                background: #C0C0C0;
                min-height: 20px;
                border-radius: 3px;
            }
            QScrollArea QScrollBar::handle:vertical:hover {
                background: #A9A9A9;
            }
            QScrollArea QScrollBar::add-line:vertical {
                height: 0px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }
            QScrollArea QScrollBar::sub-line:vertical {
                height: 0px;
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
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.sidebar = QFrame()
        self.sidebar.setStyleSheet("background-color: #FFFFFF; border: 0px;")
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setAlignment(Qt.AlignTop)
        self.sidebar_layout.setSpacing(5)
        scroll_area.setWidget(self.sidebar)
        left_layout.addWidget(scroll_area)

        for menu in self.sidebar_menus:
            btn = QPushButton(menu)
            btn.setFont(QFont("Maersk Headline", 10))
            if menu in ["Website", "Desktop App", "Arbitrary"]:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #FFFFFF;
                        color: #6A6A6A;
                        padding: 8px 0 8px 10px;
                        border: none;
                        width: 100%;
                        text-align: left;
                        font-weight: normal;
                        min-width: 0;
                        max-width: 100%;
                        white-space: normal;
                    }
                    QPushButton:hover {
                        background-color: #F5F5F5;
                        color: #444444;
                    }
                """)
            else:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #FFFFFF;
                        color: #003E62;
                        padding: 10px 0 10px 10px;
                        border: none;
                        width: 100%;
                        text-align: left;
                        font-weight: bold;
                        min-width: 0;
                        max-width: 100%;
                        white-space: normal;
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
            elif menu == "Setting":
                btn.clicked.connect(self.handle_settings)

            self.sidebar_layout.addWidget(btn)

        left_layout.addWidget(scroll_area)
        left_frame.setLayout(left_layout)

        right_frame = QFrame()
        right_frame.setStyleSheet("background-color: #F0F0F0;")
        right_layout = QVBoxLayout(right_frame)
        right_layout.setContentsMargins(0, 0, 0, 0)

        main_content = QFrame()
        main_content.setStyleSheet("background-color: #F0F0F0; border: 0px solid #D4D4D4; border-radius: 5px;")
        main_content_layout = QHBoxLayout(main_content)
        main_content_layout.setAlignment(Qt.AlignCenter)

        content_frame = QFrame()
        content_frame.setStyleSheet("background-color: #FFFFFF; border: 0px solid #D4D4D4; border-radius: 5px;")
        content_layout = QVBoxLayout(content_frame)

        scroll_area_content = QScrollArea()
        scroll_area_content.setWidgetResizable(True)
        scroll_area_content.setStyleSheet("""
            QScrollArea {
                border: 0px;
            }
            QScrollArea QScrollBar:vertical {
                border: none;
                background: #F5F5F5;
                width: 6px;
                margin: 0px 0px 0px 0px;
            }
            QScrollArea QScrollBar::handle:vertical {
                background: #C0C0C0;
                min-height: 20px;
                border-radius: 3px;
            }
            QScrollArea QScrollBar::handle:vertical:hover {
                background: #A9A9A9;
            }
            QScrollArea QScrollBar::add-line:vertical {
                height: 0px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }
            QScrollArea QScrollBar::sub-line:vertical {
                height: 0px;
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
        scroll_area_content.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area_content.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.settings_frame = QFrame()
        self.settings_frame.setStyleSheet(
            "background-color: #FFFFFF; border: 0px solid #D4D4D4; border-radius: 5px; padding: 10px;")
        self.settings_layout = QVBoxLayout(self.settings_frame)
        scroll_area_content.setWidget(self.settings_frame)
        content_layout.addWidget(scroll_area_content)

        main_content_layout.addWidget(content_frame)
        right_layout.addWidget(main_content)

        bottom_frame = QFrame()
        bottom_frame.setStyleSheet("background-color: #F0F0F0;")
        bottom_layout = QVBoxLayout(bottom_frame)

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

        self.logging_textbox = QTextEdit()
        self.logging_textbox.setFont(QFont("Maersk Headline", 10))
        self.logging_textbox.setStyleSheet("""
            QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #D4D4D4;
                border-radius: 5px;
                padding: 5px;
                color: #363636;
                min-height: 200px;
                min-width: 800px;
            }
            QTextEdit QScrollBar:vertical {
                border: none;
                background: #F5F5F5;
                width: 6px;
                margin: 0px 0px 0px 0px;
            }
            QTextEdit QScrollBar::handle:vertical {
                background: #C0C0C0;
                min-height: 20px;
                border-radius: 3px;
            }
            QTextEdit QScrollBar::handle:vertical:hover {
                background: #A9A9A9;
            }
            QTextEdit QScrollBar::add-line:vertical {
                height: 0px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }
            QTextEdit QScrollBar::sub-line:vertical {
                height: 0px;
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
        self.logging_textbox.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        bottom_layout.addWidget(self.logging_textbox)

        right_layout.addWidget(bottom_frame)

        splitter.addWidget(left_frame)
        splitter.addWidget(right_frame)

        # Set initial sizes to achieve 3:7 ratio (total width = 1200)
        splitter.setSizes([240, 960])  # 3 parts (360) for left, 7 parts (840) for right
        splitter.setStretchFactor(0, 2)  # Left frame: 3 parts
        splitter.setStretchFactor(1, 8)  # Right frame: 7 parts
        splitter.setHandleWidth(1)
        splitter.setStyleSheet("QSplitter::handle { background-color: #D4D4D4; }")

        main_layout.addWidget(splitter)

        self.setup_custom_logger()
        self.handle_homepage("HomePage")

    # Implement mouse event handlers for dragging the window via the header
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.header.geometry().contains(event.pos()):
            self.drag_position = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton and self.drag_position is not None:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_position = None
            event.accept()

    def setup_custom_logger(self):
        setup_textbox_logger(self.logging_textbox)

    def get_ui_settings(self) -> Dict[str, str]:
        return self.current_task_settings

    def set_ui_settings(self, new_ui_setting_values: Dict[str, str]) -> Dict[str, str]:
        self.current_task_settings = new_ui_setting_values
        return self.current_task_settings

    def get_task_name(self) -> str:
        return self.current_task_name

    def get_task_instance(self) -> Optional[AutomatedTask]:
        return self.automated_task

    def handle_incoming_event(self, event: Event) -> None:
        self.event_handler.handle_incoming_event(event)

    def closeEvent(self, event):
        persist_settings_to_file(self.current_task_name, self.current_task_settings)

        if hasattr(self, 'task_field_worker') and self.task_field_worker.isRunning():
            self.task_field_worker.quit()
            self.task_field_worker.wait()

        if self.com_initialized:
            pythoncom.CoUninitialize()
        event.accept()

    def toggle_task_list(self, menu: str, dir_name: str):
        self.logger.debug(f"Toggling tasks for menu: {menu}, directory: {dir_name}")

        if menu in self.task_buttons and self.task_buttons[menu]:
            are_tasks_visible = self.task_buttons[menu][0].isVisible() if self.task_buttons[menu] else False
            for btn in self.task_buttons[menu]:
                btn.setVisible(not are_tasks_visible)
            return

        for menu_tasks in self.task_buttons.values():
            for btn in menu_tasks:
                btn.setParent(None)
        self.task_buttons.clear()

        # if getattr(sys, 'frozen', False):
        #     base_dir = os.path.dirname(sys.executable)
        # else:
        #     base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        # task_dir = os.path.join(base_dir, 'src', 'task', dir_name)
        # Determine base directory
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)  # e.g., C:\automation_tool
        else:
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        # Set task_dir to the root task directory
        task_dir = os.path.join(base_dir, 'src', 'task', dir_name)

        self.logger.debug(f"Task directory path: {task_dir}")
        if not os.path.exists(task_dir):
            self.logger.error(f"Directory {task_dir} does not exist.")
            return

        task_files = [f[:-3] for f in os.listdir(task_dir) if f.endswith('.py') and f != '__init__.py']
        self.logger.debug(f"Found task files: {task_files}")
        if not task_files:
            self.logger.info(f"No tasks found in {dir_name} directory.")
            return

        self.task_buttons[menu] = []
        menu_index = self.sidebar_menus.index(menu)
        for task_name in task_files:
            task_btn = QPushButton(task_name)
            task_btn.setFont(QFont("Maersk Headline", 10))
            task_btn.setStyleSheet("""
                QPushButton {
                    background-color: #FFFFFF;
                    color: #6A6A6A;
                    padding: 5px 0 5px 20px;
                    border: 1px;
                    width: 100%;
                    text-align: left;
                    font-weight: normal;
                    min-width: 0;
                    max-width: 100%;
                    white-space: normal;
                }
                QPushButton:hover {
                    background-color: #F5F5F5;
                    color: #444444;
                }
            """)
            task_btn.clicked.connect(lambda checked, t=task_name: self.render_task_fields(t))
            task_btn.setVisible(True)
            self.task_buttons[menu].append(task_btn)
            self.sidebar_layout.insertWidget(menu_index + 1, task_btn)

        for other_menu in self.sidebar_menus:
            if other_menu != menu and other_menu in self.task_buttons:
                for btn in self.task_buttons[other_menu]:
                    btn.setVisible(False)

    def render_task_fields(self, task_name: str):
        self.clear_settings_layout()
        for i in reversed(range(self.settings_layout.count())):
            self.settings_layout.itemAt(i).widget().setParent(None)

        self.logger.debug(f'Display fields for task {task_name}')

        loading_label = QLabel("Loading task fields...")
        loading_label.setFont(QFont("Maersk Headline", 10))
        loading_label.setAlignment(Qt.AlignCenter)
        self.settings_layout.addWidget(loading_label)

        self.task_field_worker = TaskFieldWorker(task_name, parent=self)
        self.task_field_worker.task_ready.connect(self._update_task_fields)
        self.task_field_worker.error_occurred.connect(self._handle_task_field_error)
        self.task_field_worker.start()

    def _update_task_fields(self, task_name: str, automated_task: object, current_task_settings: object,
                            mandatory_settings: object, task_menu: str):
        try:
            for i in reversed(range(self.settings_layout.count())):
                self.settings_layout.itemAt(i).widget().setParent(None)

            automated_task = automated_task
            current_task_settings = dict(current_task_settings)
            mandatory_settings = list(mandatory_settings)

            self.current_task_name = task_name
            self.automated_task = automated_task
            self.current_task_settings = current_task_settings

            for setting in mandatory_settings:
                if setting == 'invoked_class':
                    continue
                if setting == 'use.GUI' and task_menu != "Website":
                    continue

                setting_frame = QFrame()
                setting_frame.setStyleSheet("background-color: #FFFFFF")
                setting_layout = QHBoxLayout(setting_frame)
                setting_layout.setContentsMargins(0, 0, 0, 0)
                setting_layout.setSpacing(2)

                label = QLabel(f"{setting}:")
                label.setFont(QFont("Maersk Headline", 10))
                label.setStyleSheet("color: #363636;")
                setting_layout.addWidget(label)

                from src.gui.UIComponentFactory import UIComponentFactory
                initial_value = current_task_settings.get(setting, '')
                component = UIComponentFactory.get_instance(self).create_component(setting, initial_value,
                                                                                   setting_frame)
                if not isinstance(component, QCheckBox):
                    component.setStyleSheet(
                        "border-bottom: 0.5px solid #D4D4D4; border-radius: 0;background-color: #FFFFFF; padding: 2px; ")
                setting_layout.addWidget(component)

                underline = QFrame()
                underline.setFrameShape(QFrame.HLine)
                underline.setFrameShadow(QFrame.Sunken)
                underline.setStyleSheet("background-color: #FF0000; height: 2px;")
                setting_layout.addWidget(underline)

                self.settings_layout.addWidget(setting_frame)

            self.automated_task.settings = self.current_task_settings
            self.update_button_states()
            self.logger.info(f'Task {task_name} is ready')

        except Exception as e:
            self.logger.error(f'Error updating task fields for {task_name}: {str(e)}', exc_info=True)
            QMessageBox.critical(self, "Error", f"An error occurred while rendering task {task_name}: {str(e)}")

    def _handle_task_field_error(self, error_message: str):
        self.logger.error(error_message, exc_info=True)
        QMessageBox.critical(self, "Error", error_message)

    def update_setting(self, setting: str, value: str):
        self.current_task_settings[setting] = value

    def handle_perform_button(self):
        if not self.current_task_name:
            self.logger.error("No task selected. Please select a task before performing.")
            return

        if self.automated_task and self.automated_task.is_alive():
            QMessageBox.information(self, "Task Running", "Please terminate the current task before running a new one")
            return

        if not self.automated_task:
            self.automated_task = create_task_instance(self.current_task_settings, self.current_task_name,
                                                       lambda: self.setup_custom_logger())

        self.progress_bar.setValue(0)
        self.progress_bar.setFormat(f"{type(self.automated_task).__name__} 0%")

        if hasattr(self, 'task_thread') and self.task_thread.isRunning():
            self.logger.warning("Waiting for previous task thread to finish...")
            self.task_thread.wait()

        self.task_thread = TaskThread(self.automated_task, lambda: None)
        self.task_thread.progress_updated.connect(self.update_progress)
        self.task_thread.finished.connect(lambda: self.on_task_finished())
        self.task_thread.start()

    def handle_pause_button(self):
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
        if not self.current_task_name:
            self.logger.info("No task is running or selected. Reset action skipped.")
            return

        if self.automated_task and self.automated_task.is_alive():
            self.logger.info(f"Terminating running task: {self.current_task_name}")
            self.automated_task.terminate()

        if hasattr(self, 'task_thread') and self.task_thread.isRunning():
            self.logger.info("Waiting for task thread to finish before reset...")
            self.task_thread.wait()

        if self.is_task_currently_pause:
            self.pause_button.setText("Pause")
            self.is_task_currently_pause = False

        self.automated_task = None
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("None 0%")
        self.logger.info('Reset task {}'.format(self.current_task_name))

    def on_task_finished(self):
        if self.is_task_currently_pause:
            self.pause_button.setText("Pause")
            self.is_task_currently_pause = False

    def update_progress(self, percent: float, task_name: str):
        self.progress_bar.setValue(round(percent))
        self.progress_bar.setFormat(f"{task_name} {percent}%")

    def toggleMaximized(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def scale_widget(self, widget, scale_factor):
        current_rect = widget.geometry()
        new_width = current_rect.width() * scale_factor
        new_height = current_rect.height() * scale_factor
        new_x = current_rect.x() - (new_width - current_rect.width()) / 2
        new_y = current_rect.y() - (new_height - current_rect.height()) / 2

        animation = QPropertyAnimation(widget, b"geometry")
        animation.setDuration(300)
        animation.setEasingCurve(QEasingCurve.OutQuad)
        animation.setStartValue(current_rect)
        animation.setEndValue(QRect(int(new_x), int(new_y), int(new_width), int(new_height)))
        animation.start()

    def handle_homepage(self, menu: str):
        self.clear_settings_layout()

        for i in reversed(range(self.settings_layout.count())):
            self.settings_layout.itemAt(i).widget().setParent(None)
        for menu_tasks in self.task_buttons.values():
            for btn in menu_tasks:
                btn.setVisible(False)

        self.logger.info("Displaying HomePage.")

        self.current_task_name = None
        self.automated_task = None
        self.current_task_settings = {}

        # Determine the base directory (installation directory when installed, script dir when not)
        if getattr(sys, 'frozen', False):  # Running as PyInstaller executable
            base_dir = os.path.dirname(sys.executable)  # Path to automation_tool.exe
        else:  # Running as script (e.g., during development)
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))

        release_notes_path = os.path.join(base_dir, 'release_notes')
        if not os.path.exists(release_notes_path):
            self.logger.error(f"Release notes directory not found at {release_notes_path}")
            release_notes_path = os.path.dirname(__file__)

        self.settings_frame.setLayout(QVBoxLayout())
        home_layout = self.settings_frame.layout()

        top_frame = QFrame()
        top_frame.setStyleSheet(
            "background-color: #FFFFFF; border: 0px solid #F0F0F0; border-radius: 5px; padding: 0px;")
        top_layout = QHBoxLayout(top_frame)
        top_layout.setSpacing(10)

        def get_versions():
            try:
                versions = []
                for f in os.listdir(release_notes_path):
                    if f.endswith('.txt'):
                        name = f.replace('.txt', '')
                        if name.startswith('v'):
                            try:
                                parts = name.replace('v', '').split('.')
                                if len(parts) == 3 and all(part.isdigit() for part in parts):
                                    versions.append(name)
                            except ValueError:
                                continue
                versions.sort(key=lambda x: [int(i) for i in x.replace('v', '').split('.')])
                return versions
            except OSError as e:
                self.logger.error(f"Error reading release notes directory: {str(e)}")
                return ['v1.0.0']

        versions = get_versions()
        if not versions:
            self.logger.warning("No version files found in release_notes directory.")
            versions = ['v1.0.0']

        latest_version = versions[-1] if versions else 'v1.0.0'
        default_message = self.get_message_for_version(latest_version, release_notes_path)
        formatted_default_message = self.format_message(default_message)

        fields = [
            ("Version", versions),
            ("Bug", "abc"),
            ("Dev", "HNL014"),
            ("Team", "Bespoke Automation Committee")
        ]

        for title, value in fields:
            field_frame = QFrame()
            field_layout = QVBoxLayout(field_frame)
            field_layout.setSpacing(5)

            title_label = QLabel(title)
            title_label.setFont(QFont("Maersk Headline", 11))
            title_label.setAlignment(Qt.AlignCenter)
            title_label.setStyleSheet("color: #003E62; text-align: center;")

            if title == "Version":
                version_combo = QComboBox()
                version_combo.setFont(QFont("Maersk Headline", 10))
                version_combo.setStyleSheet(f"""
                    QComboBox {{
                        background-color: #FFFFFF;
                        color: #003E62;
                        padding: 10px 15px;
                        border: 1px solid #D4D4D4;
                        border-radius: 5px;
                        min-width: 150px;
                        min-height: 80px;
                        max-width: 300px;
                        max-height: 160px;
                        text-align: center;
                    }}
                    QComboBox:hover {{
                        background-color: #003E62;
                        border: 0px solid #42B0D5;
                        color: #FFFFFF;
                        content: "▼";
                    }}
                    QComboBox::drop-down {{
                        border: none;
                        width: 0px;
                    }}
                    QComboBox QAbstractItemView {{
                        background-color: #FFFFFF;
                        border: 1px solid #D4D4D4;
                        selection-background-color: #42B0D5;
                    }}
                """)
                version_combo.addItems(value)
                version_combo.setCurrentText(latest_version)

                shadow = QGraphicsDropShadowEffect(version_combo)
                shadow.setBlurRadius(15)
                shadow.setXOffset(0)
                shadow.setYOffset(5)
                shadow.setColor(Qt.gray)
                version_combo.setGraphicsEffect(shadow)

                version_combo.enterEvent = lambda event: self.scale_widget(version_combo, scale_factor=2.0)
                version_combo.leaveEvent = lambda event: self.scale_widget(version_combo, scale_factor=1.0)

                def on_version_changed(index):
                    selected_version = version_combo.itemText(index)
                    message = self.get_message_for_version(selected_version, release_notes_path)
                    formatted_message = self.format_message(message)
                    description.setText(f'{formatted_message}')

                version_combo.currentIndexChanged.connect(on_version_changed)
                version_combo.setObjectName(f"VersionCombo_{title}")
                field_layout.addWidget(title_label)
                field_layout.addWidget(version_combo)
            else:
                button = QPushButton(value)
                button.setFont(QFont("Maersk Headline", 10))
                button.setStyleSheet(f"""
                    QPushButton {{
                        background-color: #FFFFFF;
                        color: #003E62;
                        padding: 10px 15px;
                        border: 1px solid #D4D4D4;
                        border-radius: 5px;
                        min-width: 150px;
                        min-height: 80px;
                        max-width: 300px;
                        max-height: 160px;
                        text-align: center;
                    }}
                    QPushButton:hover {{
                        background-color: #003E62;
                        border: 0px solid #42B0D5;
                        color: #FFFFFF;
                    }}
                """)
                shadow = QGraphicsDropShadowEffect(button)
                shadow.setBlurRadius(15)
                shadow.setXOffset(0)
                shadow.setYOffset(5)
                shadow.setColor(Qt.gray)
                button.setGraphicsEffect(shadow)

                button.enterEvent = lambda event: self.scale_widget(button, scale_factor=2.0)
                button.leaveEvent = lambda event: self.scale_widget(button, scale_factor=1.0)

                button.clicked.connect(lambda checked, t=title, v=value: self.logger.info(f"Clicked {t}: {v}"))
                button.setObjectName(f"Button_{title}")
                field_layout.addWidget(title_label)
                field_layout.addWidget(button)

            top_layout.addWidget(field_frame)

        home_layout.addWidget(top_frame)

        bottom_frame = QFrame()
        bottom_frame.setStyleSheet(
            "background-color: #FFFFFF; border: 0px solid #D4D4D4; border-radius: 5px; padding: 5px 5px 5px 10px; margin-top: 0px;")
        bottom_layout = QVBoxLayout(bottom_frame)
        bottom_layout.setAlignment(Qt.AlignLeft)

        description = QTextEdit()
        description.setFont(QFont("Maersk Text", 10))
        description.setStyleSheet(
            "color: #6A6A6A; background-color: #FFFFFF; border-left: none; padding: 5px;")
        description.setReadOnly(True)
        description.setAlignment(Qt.AlignLeft)
        description.document().setDefaultStyleSheet("""
            p {
                margin-top: 2px;
                margin-bottom: 2px;
            }
        """)
        description.setText(f'{formatted_default_message}')
        bottom_layout.addWidget(description)

        home_layout.addWidget(bottom_frame)

    def get_message_for_version(self, version, path):
        try:
            file_path = os.path.join(path, f"{version}.txt")
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read().strip()
            else:
                self.logger.warning(f"File {version}.txt not found in release_notes directory.")
                return "No message available"
        except Exception as e:
            self.logger.error(f"Error reading file {version}.txt: {str(e)}")
            return "Error reading message"

    def format_message(self, message):
        if not message:
            return ""
        lines = message.split('\n')
        formatted_lines = []
        for line in lines:
            if line.strip():
                formatted_lines.append(f"⚓ {line.strip()}")
            else:
                formatted_lines.append("")
        return '\n'.join(formatted_lines)

    def handle_settings(self):
        self.clear_settings_layout()

        self.clear_settings_layout()

        self.current_task_name = None
        self.automated_task = None
        self.current_task_settings = {}

        self.update_button_states()

        tab_widget = QTabWidget()
        tab_widget.setFont(QFont("Maersk Headline", 10))
        tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 0px solid #D4D4D4;
                background-color: #FFFFFF;
                border-radius: 5px;
                color: #363636;
            }
            QTabBar::tab {
                background-color: #F0F0F0;
                color: #363636;
                padding: 8px 15px;
                border: 1px solid #D4D4D4;
                border-bottom: none;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                min-width: 100px;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                color: #003E62;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background-color: #E0E0E0;
            }
        """)

        account_tab = QWidget()
        account_layout = QVBoxLayout(account_tab)

        account_list_label = QLabel("Account List")
        account_list_label.setFont(QFont("Maersk Headline"))
        account_list_label.setStyleSheet("color: #141414; font-size: 14px; font-weight: bold;")
        account_layout.addWidget(account_list_label)

        account_list_frame = QFrame()
        account_list_frame.setFont(QFont("Maersk Headline", 10))
        account_list_layout = QHBoxLayout(account_list_frame)
        account_list_layout.setSpacing(15)

        outlook_path = os.path.join("C:\\Users", getpass.getuser(), "AppData", "Local", "Microsoft", "Outlook")
        emails = []
        try:
            if os.path.exists(outlook_path):
                for filename in os.listdir(outlook_path):
                    if filename.lower().endswith(".ost"):
                        account_name = filename.replace(".ost", "").replace(".OST", "")
                        emails.append(account_name)
            else:
                self.logger.warning(f"Outlook directory not found at {outlook_path}")
                emails = ["Cannot retrieve email"]
        except Exception as e:
            self.logger.error(f"Error accessing Outlook directory: {str(e)}")
            emails = ["Cannot retrieve email"]

        for email in emails:
            account_button = QPushButton(email)
            account_button.setFont(QFont("Maersk Headline", 10))
            account_button.setStyleSheet("""
                QPushButton {
                    background-color: #42B0D5;
                    color: #FFFFFF;
                    padding: 5px 10px;
                    border: none;
                    border-radius: 5px;
                    min-width: 150px;
                    min-height: 20px;
                    text-align: center;
                }
                QPushButton:hover {
                    background-color: #003E62;
                    color: #FFFFFF;
                }
            """)
            account_list_layout.addWidget(account_button)
        account_list_layout.addStretch()
        account_layout.addWidget(account_list_frame)

        folder_tree_label = QLabel("Folder Tree")
        folder_tree_label.setFont(QFont("Maersk Headline"))
        folder_tree_label.setStyleSheet("color: #000000; font-size: 14px; font-weight: bold;")
        account_layout.addWidget(folder_tree_label)

        self.account_tree_widget = QTreeWidget()
        self.account_tree_widget.setHeaderLabels([""])
        self.account_tree_widget.setFont(QFont("Maersk Headline", 10))
        self.account_tree_widget.setStyleSheet("""
            QTreeWidget {
                background-color: #FFFFFF;
                border: none;
                border-radius: 5px;
                padding: 5px;
                min-height: 500px;
            }
            QTreeWidget::item {
                padding: 3px;
            }
            QTreeWidget::item:hover {
                background-color: #F0F0F0;
                color: #141414;
            }
        """)

        style = self.style()
        folder_icon = style.standardIcon(QStyle.SP_DirIcon)

        for email in emails:
            email_item = QTreeWidgetItem(self.account_tree_widget, [email])
            email_item.setIcon(0, folder_icon)

        def populate_folders_on_expand(item):
            if item.childCount() == 0:
                if item.parent() is None:
                    email = item.text(0)
                    if email in emails:
                        self._populate_folders_from_outlook(email, item)

        def on_item_clicked(item, column):
            if not item.isExpanded():
                item.setExpanded(True)
                populate_folders_on_expand(item)
            else:
                item.setExpanded(False)

        self.account_tree_widget.itemExpanded.connect(populate_folders_on_expand)
        self.account_tree_widget.itemClicked.connect(on_item_clicked)

        account_layout.addWidget(self.account_tree_widget)
        account_layout.addStretch()
        tab_widget.addTab(account_tab, "Account")

        icon_tab = QWidget()
        icon_layout = QVBoxLayout(icon_tab)
        icon_layout.addWidget(QLabel("Icon settings will go here"))
        icon_layout.addStretch()
        tab_widget.addTab(icon_tab, "Icon")

        update_tab = QWidget()
        update_layout = QVBoxLayout(update_tab)
        update_layout.addWidget(QLabel("Update settings will go here"))
        update_layout.addStretch()
        tab_widget.addTab(update_tab, "Update")

        check_mmd_tab = QWidget()
        check_mmd_layout = QVBoxLayout(check_mmd_tab)
        check_mmd_layout.setAlignment(Qt.AlignCenter)
        computer_name = socket.gethostname()
        is_mmd_device = computer_name.startswith("MMD")

        smiley_label = QLabel()
        smiley_pixmap = QPixmap("resource/img/smiley.png")
        if smiley_pixmap.isNull():
            smiley_label.setText("😊")
            smiley_font = QFont("Maersk Text", 32)
            smiley_label.setFont(smiley_font)
        else:
            smiley_label.setPixmap(smiley_pixmap.scaled(50, 50, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        smiley_label.setAlignment(Qt.AlignCenter)
        check_mmd_layout.addWidget(smiley_label)

        main_message = QLabel(
            "Great, you are a fortunate Blue Star" if is_mmd_device else "This is a TMD device, no MMD access")
        main_message.setFont(QFont("Maersk Headline", 15))
        main_message.setStyleSheet("color: #42B0D5;" if is_mmd_device else "color: #EA5D4B;")
        main_message.setAlignment(Qt.AlignCenter)
        check_mmd_layout.addWidget(main_message)

        sub_message = QLineEdit(
            f"[{computer_name}] is a MMD device." if is_mmd_device else f"[{computer_name}] is not an MMD device.")
        sub_message.setFont(QFont("Maersk Text", 10))
        sub_message.setStyleSheet("color: #6A6A6A;")
        sub_message.setAlignment(Qt.AlignCenter)
        sub_message.setReadOnly(True)
        check_mmd_layout.addWidget(sub_message)

        check_detail_button = QPushButton("Check Detail")
        check_detail_button.setFont(QFont("Maersk Headline", 10))
        check_detail_button.setStyleSheet("""
                QPushButton {
                    background-color: #003E62;
                    color: #FFFFFF;
                    padding: 8px 15px;
                    border: none;
                    border-radius: 5px;
                    min-width: 120px;
                }
                QPushButton:hover {
                    background-color: #42B0D5;
                }
                QPushButton:pressed {
                    background-color: #1686BD;
                }
            """)
        check_mmd_layout.addWidget(check_detail_button)

        self.detail_frame = QFrame()
        self.detail_frame.setStyleSheet(
            "background-color: #363636; border: 0px solid #D4D4D4; border-radius: 5px; padding: 3px;")
        self.detail_layout = QVBoxLayout(self.detail_frame)
        self.detail_layout.setAlignment(Qt.AlignLeft)
        self.detail_frame.setVisible(False)
        check_mmd_layout.addWidget(self.detail_frame)

        check_detail_button.clicked.connect(lambda: self.show_mmd_details(computer_name, is_mmd_device))

        check_mmd_layout.addStretch()
        tab_widget.addTab(check_mmd_tab, "Check MMD")

        draft_tab = QWidget()
        draft_layout = QVBoxLayout(draft_tab)
        general_settings = self.load_general_settings()
        default_path_frame = QFrame()
        default_path_layout = QHBoxLayout(default_path_frame)
        default_path_label = QLabel("Default Path: ")
        default_path_label.setStyleSheet("color: #363636;")
        default_path_layout.addWidget(default_path_label)
        default_path_input = QLineEdit()
        default_path_input.setStyleSheet(
            "border-bottom: 0.5px solid #D4D4D4; border-radius: 0; background-color: #FFFFFF; padding: 2px;")
        default_path_input.setPlaceholderText("Enter default path")
        default_path_input.setObjectName("default_path_input")
        default_path_input.setText(general_settings.get('default_path', ''))
        default_path_layout.addWidget(default_path_input)
        default_path_underline = QFrame()
        default_path_underline.setFrameShape(QFrame.HLine)
        default_path_underline.setFrameShadow(QFrame.Sunken)
        default_path_underline.setStyleSheet("background-color: #FF0000; height: 2px;")
        default_path_layout.addWidget(default_path_underline)
        draft_layout.addWidget(default_path_frame)
        log_level_frame = QFrame()
        log_level_layout = QHBoxLayout(log_level_frame)
        log_level_label = QLabel("Log Level: ")
        log_level_label.setStyleSheet("color: #363636;")
        log_level_layout.addWidget(log_level_label)
        log_level_combo = QComboBox()
        log_level_combo.addItems(["INFO", "DEBUG", "WARNING", "ERROR"])
        log_level_combo.setCurrentText(general_settings.get('log_level', 'INFO'))
        log_level_combo.setObjectName("log_level_combo")
        log_level_layout.addWidget(log_level_combo)
        log_level_underline = QFrame()
        log_level_underline.setFrameShape(QFrame.HLine)
        log_level_underline.setFrameShadow(QFrame.Sunken)
        log_level_underline.setStyleSheet("background-color: #FF0000; height: 2px;")
        log_level_layout.addWidget(log_level_underline)
        draft_layout.addWidget(log_level_frame)
        theme_frame = QFrame()
        theme_layout = QHBoxLayout(theme_frame)
        theme_label = QLabel(" Theme: ")
        theme_label.setStyleSheet("color: #363636;")
        theme_layout.addWidget(theme_label)
        theme_combo = QComboBox()
        theme_combo.addItems(["Light", "Dark"])
        theme_combo.setCurrentText(general_settings.get('theme', 'Light'))
        theme_combo.setObjectName("theme_combo")
        theme_layout.addWidget(theme_combo)
        theme_underline = QFrame()
        theme_underline.setFrameShape(QFrame.HLine)
        theme_underline.setFrameShadow(QFrame.Sunken)
        theme_underline.setStyleSheet("background-color: #FF0000; height: 2px;")
        theme_layout.addWidget(theme_underline)
        draft_layout.addWidget(theme_frame)
        save_button = QPushButton("Save Settings")
        save_button.clicked.connect(self.save_general_settings)
        draft_layout.addWidget(save_button)
        draft_layout.addStretch()
        tab_widget.addTab(draft_tab, "Draft")

        self.settings_layout.addWidget(tab_widget)
        self.settings_layout.addStretch()

    def show_mmd_details(self, computer_name, is_mmd_device):
        import platform
        import psutil
        import socket

        while self.detail_layout.count():
            item = self.detail_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        details = []
        details.append(f"Computer Name: {computer_name}")
        details.append("Status: MMD Access Granted" if is_mmd_device else "Status: No MMD Access (TMD Device)")

        architecture = "64-bit" if platform.machine().endswith('64') else "32-bit"
        details.append(f"Architecture: {architecture}")

        try:
            cpu_name = platform.processor() or "Unknown Processor"
            details.append(f"Processor: {cpu_name}")
        except Exception as e:
            details.append(f"Processor: Error retrieving data ({str(e)})")

        try:
            cpu_freq = psutil.cpu_freq()
            if cpu_freq:
                details.append(
                    f"CPU Frequency: Current: {cpu_freq.current:.2f} MHz (Min: {cpu_freq.min:.2f} MHz, Max: {cpu_freq.max:.2f} MHz)")
            else:
                details.append("CPU Frequency: Not available")
        except Exception as e:
            details.append(f"CPU Frequency: Error retrieving data ({str(e)})")

        try:
            ram = psutil.virtual_memory()
            details.append(
                f"RAM: {ram.percent}% used ({ram.used / (1024 ** 3):.2f} GB used of {ram.total / (1024 ** 3):.2f} GB)")
        except Exception as e:
            details.append(f"RAM: Error retrieving data ({str(e)})")

        try:
            disk_usage = psutil.disk_usage('C:\\')
            percent_used = disk_usage.percent
            details.append(
                f"C Disk: {percent_used}% ({disk_usage.used / (1024 ** 3):.2f} GB used of {disk_usage.total / (1024 ** 3):.2f} GB)")
        except Exception as e:
            details.append(f"C Disk: Error retrieving data ({str(e)})")

        try:
            hostname = socket.gethostname()
            ip_addresses = socket.getaddrinfo(hostname, None)
            ipv4 = None
            ipv6 = None
            for addr in ip_addresses:
                if addr[0] == socket.AF_INET:
                    ipv4 = addr[4][0]
                elif addr[0] == socket.AF_INET6:
                    ipv6 = addr[4][0]
            details.append(f"IPv4: {ipv4 if ipv4 else 'Not available'}")
            details.append(f"IPv6: {ipv6 if ipv6 else 'Not available'}")
        except Exception as e:
            details.append(f"IP Address: Error retrieving data ({str(e)})")

        for detail in details:
            label = QTextEdit(detail)
            label.setFont(QFont("Maersk Text", 10))
            label.setStyleSheet("color: #FFFFFF;")
            self.detail_layout.addWidget(label)

        self.detail_frame.setVisible(True)

    def _populate_folders_from_outlook(self, email, parent_item):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            account = None
            for acc in namespace.Accounts:
                if acc.SmtpAddress.lower() == email.lower() or acc.DisplayName.lower() == email.lower():
                    account = acc
                    break

            if not account:
                self.logger.warning(f"Account {email} not found in Outlook.")
                QTreeWidgetItem(parent_item, ["Account not accessible"])
                return

            folders = namespace.Folders.Item(account.DisplayName).Folders
            self._populate_folders(folders, parent_item)

            del folders
            del namespace
            del outlook

        except Exception as e:
            self.logger.error(f"Error accessing Outlook for {email}: {str(e)}")
            QTreeWidgetItem(parent_item, ["Error loading folders"])

    def _populate_folders(self, folders, parent_item):
        for folder in folders:
            folder_item = QTreeWidgetItem(parent_item, [folder.Name])
            if folder.Folders.Count > 0:
                self._populate_folders(folder.Folders, folder_item)

    def load_general_settings(self):
        setting_file = os.path.join(PathResolvingService.get_instance().get_input_dir(), 'general.properties')
        if not os.path.exists(setting_file):
            return {}
        return load_key_value_from_file_properties(setting_file)

    def save_general_settings(self):
        general_settings = {}
        default_path_input = self.findChild(QLineEdit, "default_path_input")
        if default_path_input:
            general_settings['default_path'] = default_path_input.text()
        log_level_combo = self.findChild(QComboBox, "log_level_combo")
        if log_level_combo:
            general_settings['log_level'] = log_level_combo.currentText()
        theme_combo = self.findChild(QComboBox, "theme_combo")
        if theme_combo:
            general_settings['theme'] = theme_combo.currentText()

        setting_file = os.path.join(PathResolvingService.get_instance().get_input_dir(), 'general.properties')
        with open(setting_file, 'w') as f:
            for key, value in general_settings.items():
                f.write(f"{key}={value}\n")
        QMessageBox.information(self, "Settings Saved", "General settings have been saved.")

    def update_button_states(self):
        has_task = self.current_task_name is not None and self.automated_task is not None
        running_task = has_task and self.automated_task.is_alive() if has_task else False

        self.perform_button.setHidden(not has_task)
        self.reset_button.setHidden(not has_task)
        self.pause_button.setHidden(not has_task)
        self.logging_textbox.setHidden(not has_task)
        self.progress_bar.setHidden(not has_task)

        if has_task:
            self.perform_button.setDisabled(False)
            self.reset_button.setDisabled(False)
            self.pause_button.setDisabled(not running_task)
        else:
            self.perform_button.setDisabled(True)
            self.reset_button.setDisabled(True)
            self.pause_button.setDisabled(True)

    def clear_settings_layout(self):
        # self.logger.info("Clearing settings_layout")
        while self.settings_layout.count():
            item = self.settings_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
                self.logger.debug(f"Deleted widget from settings_layout")

        if hasattr(self, 'account_tree_widget'):
            self.account_tree_widget.deleteLater()
            del self.account_tree_widget


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GUIApp()
    window.show()
    sys.exit(app.exec_())

from logging import Logger
from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QWidget, QLineEdit, QPushButton, QCheckBox, QHBoxLayout, QFileDialog, QLabel

from src.common.FileUtil import persist_settings_to_file
from src.common.ThreadLocalLogger import get_current_logger
from src.gui.UITaskPerformingStates import UITaskPerformingStates


class UIComponentFactory:
    _instance = None

    def __init__(self, app: UITaskPerformingStates):
        self.app = app

    @staticmethod
    def get_instance(app: UITaskPerformingStates) -> 'UIComponentFactory':
        if app is None:
            raise Exception('Must provide the GUI app instance')

        if UIComponentFactory._instance is None:
            UIComponentFactory._instance = UIComponentFactory(app)
            return UIComponentFactory._instance

        if UIComponentFactory._instance.app is not app:
            UIComponentFactory._instance = UIComponentFactory(app)

        return UIComponentFactory._instance

    def create_component(self, setting_key: str, setting_value: str, parent_widget: QWidget) -> Optional[QWidget]:
        setting_key_in_lowercase: str = setting_key.lower()

        if setting_key_in_lowercase.endswith('invoked_class'):
            return None

        if setting_key_in_lowercase.startswith('use.'):
            return self.create_checkbox_input(setting_key, setting_value, parent_widget)

        if setting_key_in_lowercase.endswith('.folder'):
            return self.create_folder_path_input(setting_key, setting_value, parent_widget)

        if setting_key_in_lowercase.endswith('.path'):
            return self.create_file_path_input(setting_key, setting_value, parent_widget)

        return self.create_textbox_input(setting_key, setting_value, parent_widget)

    def create_textbox_input(self, setting_key: str, setting_value: str, parent_widget: QWidget) -> QLineEdit:
        def update_field_data(text: str):
            try:
                field_name = input_field.property("setting_key")
                self.app.get_ui_settings()[field_name] = text
                self.app.get_task_instance().settings = self.app.get_ui_settings()
                persist_settings_to_file(self.app.get_task_name(), self.app.get_ui_settings())

                logger: Logger = get_current_logger()
                logger.debug(f"Change data on field {field_name} to {text}")
            except Exception as e:
                logger.error(f"Error updating field {field_name}: {str(e)}")

        layout = QHBoxLayout()
        layout.setSpacing(10)  # Add spacing between elements
        parent_widget.setLayout(layout)

        label = QLabel(setting_key)
        label.setFont(QFont("Maersk Headline", 9))
        label.setStyleSheet("color: #363636; background-color: #00243D; padding: 2px 5px;")
        layout.addWidget(label)

        input_field = QLineEdit(parent_widget)
        input_field.setFont(QFont("Maersk Headline", 9))
        input_field.setStyleSheet("""
            QLineEdit {
                background-color: #F0F0F0;
                border: 1px solid #D4D4D4;
                border-radius: 5px;
                padding: 5px;
                color: #363636;
            }
            QLineEdit:hover {
                border: 1px solid #1686BD;
            }
        """)
        input_field.setText('' if setting_value is None else setting_value)
        input_field.setProperty("setting_key", setting_key)
        input_field.textChanged.connect(update_field_data)
        layout.addWidget(input_field)
        layout.addStretch(1)  # Add stretch to push button to the right

        return input_field

    def create_folder_path_input(self, setting_key: str, setting_value: str, parent_widget: QWidget) -> QLineEdit:
        def choose_dir_callback():
            try:
                dir_path = QFileDialog.getExistingDirectory(parent_widget, "Choose Folder", "")
                if dir_path:
                    input_field.setText(dir_path)
                    update_field_data(dir_path)
            except Exception as e:
                logger = get_current_logger()
                logger.error(f"Error choosing directory: {str(e)}")

        def update_field_data(text: str):
            try:
                field_name = input_field.property("setting_key")
                self.app.get_ui_settings()[field_name] = text
                self.app.get_task_instance().settings = self.app.get_ui_settings()
                persist_settings_to_file(self.app.get_task_name(), self.app.get_ui_settings())
                logger: Logger = get_current_logger()
                logger.debug(f"Change data on field {field_name} to {text}")
            except Exception as e:
                logger.error(f"Error updating field {field_name}: {str(e)}")

        input_field = self.create_textbox_input(setting_key, setting_value, parent_widget)

        choose_btn = QPushButton("Choose Folder")
        choose_btn.setFont(QFont("Maersk Headline", 9))
        choose_btn.setStyleSheet("""
            QPushButton {
                background-color: #2FACE8;
                color: #FFFFFF;
                padding: 5px 15px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #42B0D5;
            }
        """)
        choose_btn.clicked.connect(choose_dir_callback)

        # Add button after the stretch
        layout = input_field.parent().layout()
        layout.addWidget(choose_btn)

        return input_field

    def create_file_path_input(self, setting_key: str, setting_value: str, parent_widget: QWidget) -> QLineEdit:
        def choose_file_callback():
            try:
                file_path, _ = QFileDialog.getOpenFileName(parent_widget, "Choose File", "", "All Files (*)")
                if file_path:
                    input_field.setText(file_path)
                    update_field_data(file_path)
            except Exception as e:
                logger = get_current_logger()
                logger.error(f"Error choosing file: {str(e)}")

        def update_field_data(text: str):
            try:
                field_name = input_field.property("setting_key")
                self.app.get_ui_settings()[field_name] = text
                self.app.get_task_instance().settings = self.app.get_ui_settings()
                persist_settings_to_file(self.app.get_task_name(), self.app.get_ui_settings())
                logger: Logger = get_current_logger()
                logger.debug(f"Change data on field {field_name} to {text}")
            except Exception as e:
                logger.error(f"Error updating field {field_name}: {str(e)}")

        input_field = self.create_textbox_input(setting_key, setting_value, parent_widget)

        choose_btn = QPushButton("Choose File")
        choose_btn.setFont(QFont("Maersk Headline", 9))
        choose_btn.setStyleSheet("""
            QPushButton {
                background-color: #2FACE8;
                color: #FFFFFF;
                padding: 5px 15px;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #42B0D5;
            }
        """)
        choose_btn.clicked.connect(choose_file_callback)

        # Add button after the stretch
        layout = input_field.parent().layout()
        layout.addWidget(choose_btn)

        return input_field

    def create_checkbox_input(self, setting_key: str, setting_value: str, parent_widget: QWidget) -> QCheckBox:
        def update_checkbox_callback(checked: bool):
            try:
                self.app.get_ui_settings()[setting_key] = str(checked)
                self.app.get_task_instance().settings = self.app.get_ui_settings()
                self.app.get_task_instance().use_gui = checked
                persist_settings_to_file(self.app.get_task_name(), self.app.get_ui_settings())
            except Exception as e:
                logger = get_current_logger()
                logger.error(f"Error updating checkbox {setting_key}: {str(e)}")

        is_checked = setting_value.lower() == 'true'

        checkbox = QCheckBox(setting_key, parent_widget)
        checkbox.setFont(QFont("Maersk Headline", 9))
        checkbox.setStyleSheet("""
            QCheckBox {
                color: #363636;
                background-color: #F0F0F0;
                padding: 5px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
            }
            QCheckBox::indicator:unchecked {
                background-color: #FFFFFF;
                border: 1px solid #D4D4D4;
                border-radius: 5px;
            }
            QCheckBox::indicator:checked {
                background-color: #2FACE8;
                border: 1px solid #1686BD;
                border-radius: 5px;
            }
        """)
        checkbox.setChecked(is_checked)
        checkbox.stateChanged.connect(lambda state: update_checkbox_callback(state == Qt.Checked))
        parent_widget.layout().addWidget(checkbox)

        return checkbox

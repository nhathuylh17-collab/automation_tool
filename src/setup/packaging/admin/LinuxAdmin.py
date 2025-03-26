import os
import subprocess
import sys
from logging import Logger

from src.common.ThreadLocalLogger import get_current_logger
from src.setup.packaging.admin.OSAdmin import OSAdmin
from src.setup.packaging.path.PathResolvingService import PathResolvingService


class LinuxAdmin(OSAdmin):

    def __init__(self, is_use_gui: bool):
        super().__init__()
        self.is_use_gui = is_use_gui

    def is_currently_running_with_admin(self) -> bool:
        return os.geteuid() == 0

    def provide_admin_privilege(self) -> None:

        logger: Logger = get_current_logger()

        if not self.is_currently_running_with_admin():
            logger.info("Executable instance is not running with root privileges. Attempting to relaunch with sudo...")
            try:
                sys.path.append(PathResolvingService.get_instance().resolve('src'))
                executable_path: str = self._get_python_3_10_in_venv(sys_executable=sys.executable)

                if self.is_use_gui:
                    python_file_script_path: str = os.path.abspath(sys.argv[0])
                    params: list[str] = [executable_path, python_file_script_path]
                    subprocess.call(['pkexec'] + params)
                    sys.exit(0)

                python_file_script_path = os.path.abspath(sys.argv[0])
                params = [executable_path, python_file_script_path]
                subprocess.call(['sudo'] + params)
                sys.exit(0)
            except Exception as e:
                logger.error(f"Failed to relaunch python_file_script_path with sudo: {e}")
                sys.exit(1)
        else:
            logger.warning("Script is already running with root privileges.")

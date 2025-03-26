import os
from abc import ABC, abstractmethod


class OSAdmin(ABC):

    @abstractmethod
    def is_currently_running_with_admin(self) -> bool:
        pass

    @abstractmethod
    def provide_admin_privilege(self) -> None:
        pass

    def _get_python_3_10_in_venv(self, sys_executable: str) -> str:

        if sys_executable.endswith('3.10') or sys_executable.endswith('pythonw.exe'):
            return sys_executable

        last_index_of_slash = sys_executable.rfind(os.sep)
        venv_python_dir = sys_executable[:last_index_of_slash]
        new_python_version = 'python3.10'
        new_path = os.path.join(venv_python_dir, new_python_version)

        if os.path.exists(new_path):
            return new_path

        raise Exception('Please use the Python 3.10 to create the virtual environment')

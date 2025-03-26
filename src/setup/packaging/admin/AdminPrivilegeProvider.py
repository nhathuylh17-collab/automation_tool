import inspect
import platform
from types import FrameType

from src.gui import GUIApp
from src.setup.packaging.admin.OSAdmin import OSAdmin


class AdminPrivilegeProvider:

    @staticmethod
    def validate_and_provide():
        platform_name = platform.system()

        provider: OSAdmin = None
        if platform_name == 'Windows':
            from src.setup.packaging.admin.WindowsAdmin import WindowsAdmin
            provider = WindowsAdmin()

        if platform_name == 'Linux':
            from src.setup.packaging.admin.LinuxAdmin import LinuxAdmin

            if AdminPrivilegeProvider.is_caller_from(GUIApp):
                provider = LinuxAdmin(is_use_gui=True)
            else:
                provider = LinuxAdmin(is_use_gui=False)

        if provider is None:
            raise Exception('Unsupported os')

        if not provider.is_currently_running_with_admin():
            provider.provide_admin_privilege()

    @staticmethod
    def is_caller_from(accepted_caller: type) -> bool:
        previous_frame: FrameType = inspect.currentframe().f_back.f_back.f_back
        caller_locals = previous_frame.f_locals

        # Check if 'self' or 'cls' is in locals and is an instance of an accepted caller
        if 'self' in caller_locals and isinstance(caller_locals['self'], accepted_caller):
            return True

        if 'cls' in caller_locals and isinstance(caller_locals['cls'], accepted_caller):
            return True

        # Script or main function, static function
        # Generate accepted filenames from the accepted caller classes
        accepted_filename = inspect.getfile(accepted_caller)

        # Check if the filename of the caller is in the accepted filenames
        caller_filename = previous_frame.f_code.co_filename
        if caller_filename == accepted_filename:
            return True

        return False

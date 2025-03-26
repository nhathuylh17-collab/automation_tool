import sys
from logging import Logger

import pyuac

from src.common.ThreadLocalLogger import get_current_logger
from src.setup.packaging.admin.OSAdmin import OSAdmin


class WindowsAdmin(OSAdmin):

    def is_currently_running_with_admin(self) -> bool:
        try:
            return pyuac.isUserAdmin()
        except Exception:
            return False

    def provide_admin_privilege(self) -> None:

        logger: Logger = get_current_logger()
        logger.info("Hung")

        if not self.is_currently_running_with_admin():
            print("Re-launching as admin!")
            pyuac.runAsAdmin()
            sys.exit(0)
        else:
            logger.warning("Script is already running with root privileges.")

import os
from logging import Logger

import win32api

from src.common.ThreadLocalLogger import get_current_logger
from src.setup.driver.query.chrome.ChromeDriverInfoQuery import ChromeDriverInfoQuery


class WindowsChromeDriverInfoQuery(ChromeDriverInfoQuery):

    def get_base_version_from_local(self) -> str:
        logger: Logger = get_current_logger()
        base_number_version: str = ''

        # Define possible Chrome executable paths
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ]

        try:
            # Find the Chrome executable
            chrome_path = None
            for path in chrome_paths:
                if os.path.exists(path):
                    chrome_path = path
                    break

            if not chrome_path:
                message = "Chrome executable not found in standard locations"
                logger.error(message)
                raise Exception(message)

            # Get version from executable metadata
            info = win32api.GetFileVersionInfo(chrome_path, "\\")
            version = f"{info['FileVersionMS'] >> 16}.{info['FileVersionMS'] & 0xFFFF}.{info['FileVersionLS'] >> 16}.{info['FileVersionLS'] & 0xFFFF}"
            base_number_version = version.split('.')[0]

            if not base_number_version:
                message = "Error retrieving Chrome version from executable"
                logger.error(message)
                raise Exception(message)

            logger.info(f"Local machine Chrome version used is {base_number_version}")
            return base_number_version

        except Exception as e:
            message = f"Error retrieving Chrome version: {str(e)}"
            logger.error(message)
            raise Exception(message)

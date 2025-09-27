from logging import Logger
from typing import Callable

from selenium.webdriver.common.by import By

from src.common.ThreadLocalLogger import get_current_logger
from src.task.WebTask import WebTask


class SCP_upload(WebTask):

    def __init__(self, settings: dict[str, str], callback_before_run_task: Callable[[], None]):
        super().__init__(settings, callback_before_run_task)
        self._document_folder = self._download_folder

    def mandatory_settings(self) -> list[str]:
        mandatory_keys: list[str] = ['username', 'password', 'download.folder', 'excel.path', 'excel.sheet',
                                     'excel.column.bill']
        return mandatory_keys

    def automate(self):
        logger: Logger = get_current_logger()
        logger.info(
            "---------------------------------------------------------------------------------------------------------")
        logger.info("Start processing")
        self._driver.get('https://scm-preprod.maersk.com/scp/cshub/home/')

        logger.info('Try to login')
        self.__login()

        key = 'msal.idtoken'
        value = self._driver.execute_script(f"return window.localStorage.getItem('{key}');")
        logger.info('Completed automation')

    def __login(self) -> None:
        username: str = self._settings['username']
        password: str = self._settings['password']

        self._type_when_element_present(by=By.ID, value='i0116', content=username)
        self._click_when_element_present(by=By.ID, value='idSIButton9')

        self.sleep()

        self._type_when_element_present(by=By.ID, value='i0118', content=password)
        self._click_when_element_present(by=By.ID, value='idSIButton9')

        while True:
            url_expect = self._driver.current_url
            if 'scm-preprod' in url_expect:
                break

            if 'login' in url_expect:
                self.sleep()
                print('sleep')

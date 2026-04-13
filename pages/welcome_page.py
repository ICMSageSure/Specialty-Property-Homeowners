from playwright.sync_api import Page


class WelcomePage:
    def __init__(self, page: Page):
        self.page = page

    def setup_wholesaler(self, wholesaler: str = "APPUND", producer: str = "APPAGT"):
        self.page.wait_for_selector("#ctl00_WelcomeUserControl_ddlWholesaler", timeout=30000)
        self.page.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option(wholesaler)
        self.page.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill(producer)
        self.page.locator("#ctl00_WelcomeUserControl_btnUpdateQuotingAs").click()
        self.page.wait_for_selector(
            "#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId", timeout=30000
        )
        self.page.locator(
            "#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId"
        ).click()
        self.page.wait_for_timeout(3000)

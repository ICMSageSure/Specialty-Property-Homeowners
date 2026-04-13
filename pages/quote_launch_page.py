from playwright.sync_api import Page


class QuoteLaunchPage:
    def __init__(self, page: Page):
        self.page = page

    def start_ho_quote(self) -> Page:
        self.page.get_by_role("paragraph").filter(has_text="Start New Quote").get_by_role("link").click()
        self.page.locator("#optionHomeowners").get_by_text("GeoVera Specialty Insurance").click()
        with self.page.expect_popup() as popup_info:
            self.page.get_by_role("button", name="OK").click()
        return popup_info.value

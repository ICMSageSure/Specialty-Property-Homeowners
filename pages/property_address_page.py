from playwright.sync_api import Page


class PropertyAddressPage:
    _AUTOCOMPLETE   = "#RadAutoCompleteBoxGeo_Input"
    _STREET         = "#txtStreet"
    _CITY           = "#txtCity"
    _STATE          = "#ddlState"
    _ZIP            = "#txtZip"
    _ZIP4           = "#txtZip4"
    _EDIT_LINK      = "#hrefEdit"
    _TARGET_LIBRARY = "#ctl00_ContentPlaceHolder1_ctl05_txtTargetLibrary"
    _SYSTEM_DATE    = "#ctl00_ContentPlaceHolder1_ctl05_txtSystemDate"

    def __init__(self, page: Page):
        self.page = page

    def wait_for_page(self):
        self.page.wait_for_selector(self._AUTOCOMPLETE, timeout=30000)

    def enter_address(self, address: str):
        self.page.locator(self._AUTOCOMPLETE).click()
        self.page.locator(self._AUTOCOMPLETE).press_sequentially(address)
        self.page.wait_for_timeout(1000)
        self.page.locator(self._AUTOCOMPLETE).press("Enter")
        self.submit()

    def edit_address_manually(self, street: str, city: str, state: str, zip_code: str, zip4: str = ""):
        """Use when autocomplete fails — click the edit link to unlock the address fields."""
        self.page.locator(self._EDIT_LINK).click()
        self.page.locator(self._STREET).fill(street)
        self.page.locator(self._CITY).fill(city)
        self.page.locator(self._STATE).select_option(state)
        self.page.locator(self._ZIP).fill(zip_code)
        if zip4:
            self.page.locator(self._ZIP4).fill(zip4)

    def set_test_env(self, target_library: str, system_date: str):
        self.page.locator(self._TARGET_LIBRARY).fill(target_library)
        self.page.locator(self._SYSTEM_DATE).fill(system_date)

    def get_street(self) -> str:
        return self.page.locator(self._STREET).input_value()

    def get_city(self) -> str:
        return self.page.locator(self._CITY).input_value()

    def get_state(self) -> str:
        return self.page.locator(self._STATE).input_value()

    def get_zip(self) -> str:
        return self.page.locator(self._ZIP).input_value()

    def submit(self):
        self.page.get_by_role("button", name="Next -->").click()

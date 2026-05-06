from playwright.sync_api import Page, expect

from pages.fg_policy_base_page import FGPolicyBasePage


class FGPolicyDetailPanelPage(FGPolicyBasePage):
    _PREFIX = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_"

    # -- Register panel --
    _PNL_REGISTER   = f"{_PREFIX}pnlRegister"
    _BTN_REGISTER   = f"{_PREFIX}btnRegister"

    # -- Paperless status panel --
    _PNL_PAPERLESS_STATUS  = f"{_PREFIX}pnlPaperlessStatus"
    _LBL_PAPERLESS_STATUS  = f"{_PREFIX}lblPaperlessUnSubsribed"

    # -- Policy details panel (left pane) --
    _LBL_PROPERTY_ADDRESS   = f"{_PREFIX}lblPropertyAddress"
    _LBL_CITY_STATE_ZIP     = f"{_PREFIX}lblCityStateZip"
    _LBL_MAIL_ADDRESS       = f"{_PREFIX}lblMailAddress"
    _LBL_MAIL_CITY_STATE_ZIP = f"{_PREFIX}lblMailCityStateZip"
    _LBL_EXPIRATION_DATE    = f"{_PREFIX}lblExpirationDate"

    # -- Policy status panel (center pane) --
    _LBL_LAST_ACTIVITY      = f"{_PREFIX}lblLastActivity"
    _LBL_LAST_ACTIVITY_DATE = f"{_PREFIX}lblLastActivityDate"
    _LBL_PRODUCT            = f"{_PREFIX}lblProduct"
    _LBL_OPTION             = f"{_PREFIX}lblOption"
    _LBL_POLICY_STATUS      = f"{_PREFIX}lblPolicyStatus"

    # -- Payment summary (right pane) --
    _LBL_OUTSTANDING_BAL  = f"{_PREFIX}lblOutStandingBalanceValue"
    _LBL_AMOUNT_DUE       = f"{_PREFIX}lblAmountDueValue"
    _LBL_PAYMENT_DUE_DATE = f"{_PREFIX}lblPaymentDueDate"

    def __init__(self, page: Page):
        super().__init__(page)

    # ------------------------------------------------------------------ #
    # Page load                                                            #
    # ------------------------------------------------------------------ #

    def wait_for_page(self) -> None:
        self.page.wait_for_selector(self._POLICY_NUMBER, timeout=30000)

    # ------------------------------------------------------------------ #
    # Register panel                                                       #
    # ------------------------------------------------------------------ #

    def is_register_panel_visible(self) -> bool:
        return self.page.locator(self._PNL_REGISTER).is_visible()

    def click_register_now(self) -> None:
        self.page.locator(self._BTN_REGISTER).click()

    # ------------------------------------------------------------------ #
    # Paperless status panel                                               #
    # ------------------------------------------------------------------ #

    def get_paperless_status(self) -> str:
        return self.page.locator(self._LBL_PAPERLESS_STATUS).inner_text().strip()

    def verify_paperless_status(self, status: str) -> None:
        expect(self.page.locator(self._LBL_PAPERLESS_STATUS)).to_have_text(status)

    # ------------------------------------------------------------------ #
    # Policy details — left pane                                           #
    # ------------------------------------------------------------------ #

    def get_property_address(self) -> str:
        return self.page.locator(self._LBL_PROPERTY_ADDRESS).inner_text().strip()

    def get_property_city_state_zip(self) -> str:
        return self.page.locator(self._LBL_CITY_STATE_ZIP).inner_text().strip()

    def get_mailing_address(self) -> str:
        return self.page.locator(self._LBL_MAIL_ADDRESS).inner_text().strip()

    def get_mailing_city_state_zip(self) -> str:
        return self.page.locator(self._LBL_MAIL_CITY_STATE_ZIP).inner_text().strip()

    def get_expiration_date(self) -> str:
        return self.page.locator(self._LBL_EXPIRATION_DATE).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Policy status — center pane                                          #
    # ------------------------------------------------------------------ #

    def get_last_activity(self) -> str:
        return self.page.locator(self._LBL_LAST_ACTIVITY).inner_text().strip()

    def get_last_activity_date(self) -> str:
        return self.page.locator(self._LBL_LAST_ACTIVITY_DATE).inner_text().strip()

    def get_product(self) -> str:
        return self.page.locator(self._LBL_PRODUCT).inner_text().strip()

    def get_option(self) -> str:
        return self.page.locator(self._LBL_OPTION).inner_text().strip()

    def get_policy_status(self) -> str:
        return self.page.locator(self._LBL_POLICY_STATUS).inner_text().strip()

    def verify_policy_status(self, status: str) -> None:
        expect(self.page.locator(self._LBL_POLICY_STATUS)).to_have_text(status)

    # ------------------------------------------------------------------ #
    # Payment summary — right pane                                         #
    # ------------------------------------------------------------------ #

    def get_outstanding_balance(self) -> str:
        return self.page.locator(self._LBL_OUTSTANDING_BAL).inner_text().strip()

    def get_amount_due(self) -> str:
        return self.page.locator(self._LBL_AMOUNT_DUE).inner_text().strip()

    def get_payment_due_date(self) -> str:
        return self.page.locator(self._LBL_PAYMENT_DUE_DATE).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Diagnostic                                                           #
    # ------------------------------------------------------------------ #

    def print_detail_panel_summary(self) -> None:
        print(f"Policy Number:     {self.get_policy_number()}")
        print(f"Insured:           {self.get_insured_name()}")
        print(f"Product:           {self.get_product()}")
        print(f"Status:            {self.get_policy_status()}")
        print(f"Expires:           {self.get_expiration_date()}")
        print(f"Property Address:  {self.get_property_address()}, {self.get_property_city_state_zip()}")
        print(f"Mailing Address:   {self.get_mailing_address()}, {self.get_mailing_city_state_zip()}")
        print(f"Last Activity:     {self.get_last_activity()} ({self.get_last_activity_date()})")
        print(f"Outstanding Bal:   {self.get_outstanding_balance()}")
        print(f"Amount Due:        {self.get_amount_due()}")
        print(f"Payment Due Date:  {self.get_payment_due_date()}")
        print(f"Paperless:         {self.get_paperless_status()}")

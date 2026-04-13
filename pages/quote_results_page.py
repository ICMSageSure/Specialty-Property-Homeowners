import re
from datetime import datetime
from playwright.sync_api import Page


class QuoteResultsPage:
    _HEADER      = "#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lblQuote1Header"
    _QUOTE_NUM   = "#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lblQuote1Number"
    _FULL_PAY    = "#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lstPaymentOptions_ctl00_lblQuote1PaymentOption"
    _THREE_PAY   = "#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lstPaymentOptions_ctl01_lblQuote1PaymentOption"
    _DOLLAR_RE   = re.compile(r"\$\d{1,3}(?:,\d{3})*(?:\.\d{2})")

    def __init__(self, page: Page):
        self.page = page

    def wait_for_results(self):
        self.page.wait_for_selector(self._HEADER, timeout=30000)

    def get_product_name(self) -> str:
        return self.page.locator(self._HEADER).inner_text()

    def get_quote_number(self) -> str:
        return self.page.locator(self._QUOTE_NUM).inner_text()

    def get_premium(self) -> str:
        full_pay = self.page.locator(self._FULL_PAY).inner_text()
        match = self._DOLLAR_RE.search(full_pay)
        return match.group() if match else full_pay

    def print_quote_summary(self):
        print(f"Product:       {self.get_product_name()}")
        print(f"Quote #:       {self.get_quote_number()}")
        print(f"Full Pay:      {self.page.locator(self._FULL_PAY).inner_text()}")
        three_pay = self.page.locator(self._THREE_PAY).text_content().strip()
        match = self._DOLLAR_RE.search(three_pay)
        print(f"3-Pay Install: {match.group() if match else three_pay}")

    def continue_to_bind(self):
        self.page.get_by_role("button", name="Continue to Bind Coverage").click()
        self.page.wait_for_timeout(2000)
        if self.page.locator("#chkPopInConfirmation").is_visible():
            self.page.locator("#chkPopInConfirmation").check()
            self.page.get_by_role("button", name="Continue", exact=True).click()


class BindRequestPage:
    _EFF_DATE   = "#ctl00_ContentPlaceHolder1_ucBinderRequest_txtBinderEffectiveDate1"
    _BILL_TO    = "#ctl00_ContentPlaceHolder1_ucBinderRequest_ddlBillTo1"
    _NON_CANCEL = "#ctl00_ContentPlaceHolder1_ucBinderRequest_chkNonCancellation"

    def __init__(self, page: Page):
        self.page = page

    def fill_and_submit(self, eff_date: str = None, bill_to: str = "0", payment_plan: str = "0"):
        if eff_date is None:
            eff_date = datetime.today().strftime("%m/%d/%Y")
        self.page.wait_for_selector(self._EFF_DATE, timeout=30000)
        self.page.locator(self._EFF_DATE).fill(eff_date)
        self.page.locator(self._BILL_TO).select_option(bill_to)
        self.page.locator("#ddlPaymentPlan1").select_option(payment_plan)
        if self.page.locator(self._NON_CANCEL).is_visible():
            self.page.locator(self._NON_CANCEL).check()
        self.page.get_by_role("button", name="Request to Bind").click()


class BinderDetailPage:
    _AMOUNT_DUE = "#ctl00_ContentPlaceHolder1_ucBinderDetail_lblAmountDue"
    _FIELDS = {
        "Binder Number":  "#ctl00_ContentPlaceHolder1_ucBinderDetail_lblBinderNumber",
        "Bill To":        "#ctl00_ContentPlaceHolder1_ucBinderDetail_lblBillTo",
        "Payment Plan":   "#ctl00_ContentPlaceHolder1_ucBinderDetail_lblPaymentPlan",
        "Effective Date": "#ctl00_ContentPlaceHolder1_ucBinderDetail_lblEffectiveDate",
        "Amount Due Now": "#ctl00_ContentPlaceHolder1_ucBinderDetail_lblAmountDue",
    }

    def __init__(self, page: Page):
        self.page = page

    def wait_and_print(self):
        self.page.wait_for_selector(self._AMOUNT_DUE, timeout=30000)
        for label, selector in self._FIELDS.items():
            print(f"{label}: {self.page.locator(selector).inner_text()}")
        self.page.wait_for_timeout(5000)

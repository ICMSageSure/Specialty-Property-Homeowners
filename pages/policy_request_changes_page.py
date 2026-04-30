from playwright.sync_api import Page, expect


class PolicyRequestChangesPage:
    _CHANGE_TYPE = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_ddlChangeType"
    _EFF_DATE    = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_txtEffectiveDate"
    _COMMENTS    = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_txtComments"
    _FILE_UPLOAD = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_fuSCRAttachFile"

    def __init__(self, page: Page):
        self.page = page

    def select_change_type(self, value: str) -> None:
        self.page.locator(self._CHANGE_TYPE).select_option(value)
        self.page.wait_for_timeout(3000)

    def fill_effective_date(self, date: str) -> None:
        self.page.locator(self._EFF_DATE).fill(date)
        self.page.locator(self._EFF_DATE).press("Tab")

    def verify_change_request_header(self, label: str) -> None:
        expect(self.page.locator("h3.policyHeader")).to_have_text(f"Request Changes to {label}")

    def fill_comment(self, text: str) -> None:
        self.page.locator(self._COMMENTS).fill(text)


    def upload_file(self, path: str) -> None:
        self.page.locator(self._FILE_UPLOAD).set_input_files(path)

    def select_cancel_reason(self, value: str) -> None:
        self.page.locator("#ddlCancelReason").select_option(value)
        self.page.wait_for_timeout(1500)

    def select_mailing_address_check(self, value: str) -> None:
        """value: 'true' = Yes, 'false' = No"""
        self.page.locator(
            "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_ddlMailingAddress"
        ).select_option(value)
        self.page.wait_for_timeout(2000)

    def click_continue(self) -> None:
        self.page.get_by_role("button", name="Continue").click()


class PolicyChangesConfirmationPage:
    _PREFIX = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_ucSCRChangeRequestConfirmation_"

    def __init__(self, page: Page):
        self.page = page

    def verify_header(self, label: str, comment: str) -> None:
        expect(self.page.get_by_text("Change requests not yet")).to_be_visible()
        expect(self.page.get_by_role("cell", name=f"{label}:", exact=True)).to_be_visible()
        expect(self.page.get_by_text(comment)).to_be_visible()

    def fill_submitter_contact(
        self,
        first: str,
        last: str,
        email: str,
        area_code: str = "707",
        lat_code: str = "555",
        local_number: str = "1111",
    ) -> None:
        self.page.locator(f"{self._PREFIX}txtFirstName").fill(first)
        self.page.locator(f"{self._PREFIX}txtFirstName").press("Tab")
        self.page.locator(f"{self._PREFIX}txtLastName").fill(last)
        self.page.locator(f"{self._PREFIX}txtEmailAddress").fill(email)
        self.page.locator(f"{self._PREFIX}txtPhoneAreaCode").fill(area_code)
        self.page.locator(f"{self._PREFIX}txtPhoneLATCode").fill(lat_code)
        self.page.locator(f"{self._PREFIX}txtPhoneLocalNumber").fill(local_number)

    def submit(self) -> None:
        self.page.get_by_role("button", name="Submit All Changes").click()

    def verify_success(self) -> None:
        expect(self.page.get_by_text("Your change requests have")).to_be_visible()

    def click_done(self) -> None:
        self.page.get_by_role("button", name="Done").click()

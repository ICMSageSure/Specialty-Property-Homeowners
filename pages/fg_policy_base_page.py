from playwright.sync_api import Page, expect


class FGPolicyBasePage:
    _POLICY_NUMBER  = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyNumber"
    _POLICY_INSURED = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyInsured"

    def __init__(self, page: Page):
        self.page = page

    def get_policy_number(self) -> str:
        return self.page.locator(self._POLICY_NUMBER).inner_text().strip()

    def get_insured_name(self) -> str:
        return self.page.locator(self._POLICY_INSURED).inner_text().strip()

    def verify_policy_header(self, policy_number: str, insured_name: str) -> None:
        expect(self.page.locator(self._POLICY_NUMBER)).to_have_text(policy_number)
        expect(self.page.locator(self._POLICY_INSURED)).to_have_text(insured_name)

from playwright.sync_api import Page, expect


class FGPolicyOverviewPage:
    # -- Policy header --
    _POLICY_NUMBER  = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyNumber"
    _POLICY_INSURED = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyInsured"

    # -- Sidebar: top-level links --
    _LNK_BILLING         = "#ctl00_ctl00_ContentPlaceHolder1_lnkBillingPayment"
    _LNK_INSURED         = "#ctl00_ctl00_ContentPlaceHolder1_lnkInsured"
    _LNK_POLICY_INFO     = "#ctl00_ctl00_ContentPlaceHolder1_lnkPolicyInfo"
    _LNK_PROPERTY        = "#ctl00_ctl00_ContentPlaceHolder1_lnkPropertyDetails"
    _LNK_CHANGE_REQUEST  = "#ctl00_ctl00_ContentPlaceHolder1_lnkChangeRequest"
    _LNK_MESSAGE_CENTER  = "#ctl00_ctl00_ContentPlaceHolder1_lnkAlerts"

    # -- Sidebar: Billing/Payment sub-links --
    _LNK_BILLING_DIRECT  = "#ctl00_ctl00_ContentPlaceHolder1_lnkBilling"
    _LNK_PAY_PLAN_OPT    = "#ctl00_ctl00_ContentPlaceHolder1_lnkPayPlanOptions"

    # -- Sidebar: Insured sub-links --
    _LNK_INSURED_INFO    = "#ctl00_ctl00_ContentPlaceHolder1_lnkInsuredInfo"
    _LNK_REGISTER        = "#ctl00_ctl00_ContentPlaceHolder1_lnkRegister"
    _LNK_COMM_PREF      = "#ctl00_ctl00_ContentPlaceHolder1_lnkCommPref"

    # -- Sidebar: Policy Information sub-links --
    _LNK_POLICY_DETAILS  = "#ctl00_ctl00_ContentPlaceHolder1_lnkPolicyDetails"
    _LNK_DOCUMENTS       = "#ctl00_ctl00_ContentPlaceHolder1_lnkDocuments"
    _LNK_REF_DOCS        = "#ctl00_ctl00_ContentPlaceHolder1_lnkReference"
    _LNK_INSPECTIONS     = "#ctl00_ctl00_ContentPlaceHolder1_lnkInspections"
    _LNK_TRANSACTIONS    = "#ctl00_ctl00_ContentPlaceHolder1_lnkTransactions"
    _LNK_COMMENTS        = "#ctl00_ctl00_ContentPlaceHolder1_lnkComents"
    _LNK_CLAIMS          = "#ctl00_ctl00_ContentPlaceHolder1_lnkClaims"
    _LNK_MORTGAGEES      = "#ctl00_ctl00_ContentPlaceHolder1_lnkMortgagees"
    _LNK_ALT_PAYOR       = "#ctl00_ctl00_ContentPlaceHolder1_lnkAlternatePayor"
    _LNK_ADD_INSURED     = "#ctl00_ctl00_ContentPlaceHolder1_lnkAdditionalInsured"
    _LNK_ADD_INTEREST    = "#ctl00_ctl00_ContentPlaceHolder1_lnkAdditionalInterest"
    _LNK_PRODUCER        = "#ctl00_ctl00_ContentPlaceHolder1_lnkProducer"

    # -- Sidebar: Property Details sub-links --
    _LNK_PROPERTY_INFO   = "#ctl00_ctl00_ContentPlaceHolder1_lnkPropertyInfo"
    _LNK_ENDORSEMENTS    = "#ctl00_ctl00_ContentPlaceHolder1_lnkOptionalEndorsements"
    _LNK_GENERAL_INFO    = "#ctl00_ctl00_ContentPlaceHolder1_lnkGeneralInfo"

    # -- Sidebar: Message Center sub-links --
    _LNK_ALERTS_INBOX    = "#ctl00_ctl00_ContentPlaceHolder1_lnkAlertInbox"
    _LNK_SEND_CORR       = "#ctl00_ctl00_ContentPlaceHolder1_lnkAlertSendCorrespondence"

    # -- Payment overview panel --
    _PAY_DUE_DATE        = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_lblPaymentDueDate"
    _LNK_ENROLL_AUTOPAY  = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_lnkEnrollAutoPay"
    _LNK_VIEW_INSTALL    = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_lnkInstallments"

    def __init__(self, page: Page):
        self.page = page

    # ------------------------------------------------------------------ #
    # Header                                                               #
    # ------------------------------------------------------------------ #

    def wait_for_page(self) -> None:
        self.page.wait_for_selector(self._POLICY_NUMBER, timeout=30000)

    def get_policy_number(self) -> str:
        return self.page.locator(self._POLICY_NUMBER).inner_text().strip()

    def get_insured_name(self) -> str:
        return self.page.locator(self._POLICY_INSURED).inner_text().strip()

    def verify_policy_header(self, policy_number: str, insured_name: str) -> None:
        expect(self.page.locator(self._POLICY_NUMBER)).to_have_text(policy_number)
        expect(self.page.locator(self._POLICY_INSURED)).to_have_text(insured_name)

    # ------------------------------------------------------------------ #
    # Sidebar navigation — top-level sections                             #
    # ------------------------------------------------------------------ #

    def go_to_billing_payment(self) -> None:
        self.page.locator(self._LNK_BILLING).click()

    def go_to_insured(self) -> None:
        self.page.locator(self._LNK_INSURED).click()

    def go_to_policy_information(self) -> None:
        self.page.locator(self._LNK_POLICY_INFO).click()

    def go_to_property_details(self) -> None:
        self.page.locator(self._LNK_PROPERTY).click()

    def go_to_change_request(self) -> None:
        self.page.locator(self._LNK_CHANGE_REQUEST).click()

    def go_to_message_center(self) -> None:
        self.page.locator(self._LNK_MESSAGE_CENTER).click()

    # ------------------------------------------------------------------ #
    # Sidebar navigation — Billing/Payment sub-links                      #
    # ------------------------------------------------------------------ #

    def go_to_payment_plan_options(self) -> None:
        self.page.locator(self._LNK_PAY_PLAN_OPT).click()

    # ------------------------------------------------------------------ #
    # Sidebar navigation — Insured sub-links                              #
    # ------------------------------------------------------------------ #

    def go_to_insured_information(self) -> None:
        self.page.locator(self._LNK_INSURED_INFO).click()

    def go_to_register_online_access(self) -> None:
        self.page.locator(self._LNK_REGISTER).click()

    def go_to_communication_preferences(self) -> None:
        self.page.locator(self._LNK_COMM_PREF).click()

    # ------------------------------------------------------------------ #
    # Sidebar navigation — Policy Information sub-links                   #
    # ------------------------------------------------------------------ #

    def go_to_policy_details(self) -> None:
        self.page.locator(self._LNK_POLICY_DETAILS).click()

    def go_to_policy_documents(self) -> None:
        self.page.locator(self._LNK_DOCUMENTS).click()

    def go_to_reference_documents(self) -> None:
        self.page.locator(self._LNK_REF_DOCS).click()

    def go_to_inspections(self) -> None:
        self.page.locator(self._LNK_INSPECTIONS).click()

    def go_to_transactions(self) -> None:
        self.page.locator(self._LNK_TRANSACTIONS).click()

    def go_to_comments(self) -> None:
        self.page.locator(self._LNK_COMMENTS).click()

    def go_to_claims(self) -> None:
        self.page.locator(self._LNK_CLAIMS).click()

    def go_to_mortgagee_details(self) -> None:
        self.page.locator(self._LNK_MORTGAGEES).click()

    def go_to_alternate_payor(self) -> None:
        self.page.locator(self._LNK_ALT_PAYOR).click()

    def go_to_additional_insured(self) -> None:
        self.page.locator(self._LNK_ADD_INSURED).click()

    def go_to_additional_interest(self) -> None:
        self.page.locator(self._LNK_ADD_INTEREST).click()

    def go_to_insurance_producer(self) -> None:
        self.page.locator(self._LNK_PRODUCER).click()

    # ------------------------------------------------------------------ #
    # Sidebar navigation — Property Details sub-links                     #
    # ------------------------------------------------------------------ #

    def go_to_property_information(self) -> None:
        self.page.locator(self._LNK_PROPERTY_INFO).click()

    def go_to_endorsements(self) -> None:
        self.page.locator(self._LNK_ENDORSEMENTS).click()

    def go_to_general_information(self) -> None:
        self.page.locator(self._LNK_GENERAL_INFO).click()

    # ------------------------------------------------------------------ #
    # Sidebar navigation — Message Center sub-links                       #
    # ------------------------------------------------------------------ #

    def go_to_alerts_inbox(self) -> None:
        self.page.locator(self._LNK_ALERTS_INBOX).click()

    def go_to_send_correspondence(self) -> None:
        self.page.locator(self._LNK_SEND_CORR).click()

    # ------------------------------------------------------------------ #
    # Payment overview panel                                              #
    # ------------------------------------------------------------------ #

    def get_payment_due_date(self) -> str:
        return self.page.locator(self._PAY_DUE_DATE).inner_text().strip()

    def click_enroll_auto_pay(self) -> None:
        self.page.locator(self._LNK_ENROLL_AUTOPAY).click()

    def click_view_remaining_installments(self) -> None:
        self.page.locator(self._LNK_VIEW_INSTALL).click()

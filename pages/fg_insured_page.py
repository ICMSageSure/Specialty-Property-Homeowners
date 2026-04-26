from playwright.sync_api import Page, expect


class FGInsuredPage:
    _PREFIX = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_"

    # -- Page header --
    _LBL_HEADER = f"{_PREFIX}lblHeader"

    # -- Policy header (shared banner) --
    _POLICY_NUMBER  = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyNumber"
    _POLICY_INSURED = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyInsured"

    # -- Term dates --
    _LBL_EFF_DATE  = f"{_PREFIX}lblCurrentTermEffectiveDate"
    _LBL_EXP_DATE  = f"{_PREFIX}lblCurrentTermExpirationDate"

    # -- Insured details: current term --
    _LBL_INSURED_TYPE_CUR  = f"{_PREFIX}lblInsuredTypeCurrentTerm"
    _LBL_INSURED_NAME_CUR  = f"{_PREFIX}lblInsured1NameCurrentTerm"
    _LBL_INSURED_SSN_CUR   = f"{_PREFIX}lblInsured1SSNCurrentTerm"
    _LBL_INSURED_DOB_CUR   = f"{_PREFIX}lblInsured1DOBCurrentTerm"

    # -- Insured details: next term --
    _LBL_INSURED_TYPE_NEXT = f"{_PREFIX}lblInsuredTypeNextTerm"
    _LBL_INSURED_NAME_NEXT = f"{_PREFIX}lblInsured1NameNextTerm"
    _LBL_INSURED_SSN_NEXT  = f"{_PREFIX}lblInsured1SSNNextTerm"
    _LBL_INSURED_DOB_NEXT  = f"{_PREFIX}lblInsured1DOBNextTerm"

    # -- Mailing address: current term --
    _LBL_STREET_CUR        = f"{_PREFIX}lblStreetCurrentTerm"
    _LBL_CITY_STATE_ZIP_CUR = f"{_PREFIX}lblCityStateZipCurrentTerm"

    # -- Mailing address: next term --
    _LBL_STREET_NEXT           = f"{_PREFIX}lblStreetNextTerm"
    _LBL_CITY_STATE_ZIP_NEXT   = f"{_PREFIX}lblCityStateZipNextTerm"

    # -- Contact: current term --
    _LBL_EMAIL_CUR         = f"{_PREFIX}lblEmailAddressCurrentTerm"
    _LBL_PRIMARY_PHONE_CUR = f"{_PREFIX}lblPrimaryPhoneCurrentTerm"
    _LBL_SECONDARY_PHONE_CUR = f"{_PREFIX}lblSecondaryPhoneCurrentTerm"

    # -- Contact: next term --
    _LBL_EMAIL_NEXT          = f"{_PREFIX}lblEmailAddressNextTerm"
    _LBL_PRIMARY_PHONE_NEXT  = f"{_PREFIX}lblPrimaryPhoneNextTerm"
    _LBL_SECONDARY_PHONE_NEXT = f"{_PREFIX}lblSecondaryPhoneNextTerm"

    # -- Prior homeowners: current term --
    _LBL_PRIOR_HOME_CUR    = f"{_PREFIX}lblPriorHomeCurrentTerm"
    _LBL_PRIOR_CARRIER_CUR = f"{_PREFIX}lblPriorCarrierCurrentTerm"
    _LBL_PRIOR_COV_CUR     = f"{_PREFIX}lblPriorCovAmtCurrentTerm"
    _LBL_PRIOR_EXP_CUR     = f"{_PREFIX}lblPriorExpirationCurrentTerm"

    # -- Prior homeowners: next term --
    _LBL_PRIOR_HOME_NEXT    = f"{_PREFIX}lblPriorHomeNextTerm"
    _LBL_PRIOR_CARRIER_NEXT = f"{_PREFIX}lblPriorCarrierNextTerm"
    _LBL_PRIOR_COV_NEXT     = f"{_PREFIX}lblPriorCovAmtNextTerm"
    _LBL_PRIOR_EXP_NEXT     = f"{_PREFIX}lblPriorExpirationNextTerm"

    # -- Edit buttons --
    _BTN_EDIT_TOP    = f"{_PREFIX}btnEdit1"
    _BTN_EDIT_BOTTOM = f"{_PREFIX}btnEdit2"

    # -- Popup: Invite to Paperless --
    _PNL_PAPERLESS          = f"{_PREFIX}pnlInviteToPaperlessPopup"
    _BTN_PAPERLESS_CANCEL   = f"{_PREFIX}btnPaperlessCancel"
    _BTN_PAPERLESS_CONFIRM  = f"{_PREFIX}btnConfirmAndSend"
    _LBL_PAPERLESS_NO_EMAIL = f"{_PREFIX}ucInviteToPaperlessPopUp_lblInviteToPaperlessNoEmail"

    # -- Popup: Paperless Confirmation --
    _PNL_PAPERLESS_CONF     = f"{_PREFIX}pnlInviteToPaperlessConfPopup"
    _LBL_PAPERLESS_CONF_MSG = f"{_PREFIX}InviteToPaperlessConfPopup_lblInviteToPaperlessConfirmation"
    _BTN_PAPERLESS_CONF_OK  = f"{_PREFIX}BtnConfOK"

    # -- Popup: Policy/Registered Name Warning --
    _PNL_NAME_WARNING    = f"{_PREFIX}pnlPolicyNameWarningPopup"
    _BTN_WARNING_CANCEL  = f"{_PREFIX}btnCancel"
    _BTN_WARNING_CONT    = f"{_PREFIX}btnContinue"

    def __init__(self, page: Page):
        self.page = page

    # ------------------------------------------------------------------ #
    # Page load                                                            #
    # ------------------------------------------------------------------ #

    def wait_for_page(self) -> None:
        self.page.wait_for_selector(self._LBL_HEADER, timeout=30000)

    # ------------------------------------------------------------------ #
    # Policy banner                                                        #
    # ------------------------------------------------------------------ #

    def get_policy_number(self) -> str:
        return self.page.locator(self._POLICY_NUMBER).inner_text().strip()

    def get_insured_name(self) -> str:
        return self.page.locator(self._POLICY_INSURED).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Term dates                                                           #
    # ------------------------------------------------------------------ #

    def get_effective_date(self) -> str:
        return self.page.locator(self._LBL_EFF_DATE).inner_text().strip()

    def get_expiration_date(self) -> str:
        return self.page.locator(self._LBL_EXP_DATE).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Insured details — read                                               #
    # ------------------------------------------------------------------ #

    def get_insured_name_current(self) -> str:
        return self.page.locator(self._LBL_INSURED_NAME_CUR).inner_text().strip()

    def get_insured_name_next(self) -> str:
        return self.page.locator(self._LBL_INSURED_NAME_NEXT).inner_text().strip()

    def get_insured_type_current(self) -> str:
        return self.page.locator(self._LBL_INSURED_TYPE_CUR).inner_text().strip()

    def get_dob_current(self) -> str:
        return self.page.locator(self._LBL_INSURED_DOB_CUR).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Mailing address — read                                               #
    # ------------------------------------------------------------------ #

    def get_street_current(self) -> str:
        return self.page.locator(self._LBL_STREET_CUR).inner_text().strip()

    def get_city_state_zip_current(self) -> str:
        return self.page.locator(self._LBL_CITY_STATE_ZIP_CUR).inner_text().strip()

    def get_street_next(self) -> str:
        return self.page.locator(self._LBL_STREET_NEXT).inner_text().strip()

    def get_city_state_zip_next(self) -> str:
        return self.page.locator(self._LBL_CITY_STATE_ZIP_NEXT).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Contact info — read                                                  #
    # ------------------------------------------------------------------ #

    def get_email_current(self) -> str:
        return self.page.locator(self._LBL_EMAIL_CUR).inner_text().strip()

    def get_primary_phone_current(self) -> str:
        return self.page.locator(self._LBL_PRIMARY_PHONE_CUR).inner_text().strip()

    def get_secondary_phone_current(self) -> str:
        return self.page.locator(self._LBL_SECONDARY_PHONE_CUR).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Prior homeowners — read                                              #
    # ------------------------------------------------------------------ #

    def get_prior_home_carrier_current(self) -> str:
        return self.page.locator(self._LBL_PRIOR_CARRIER_CUR).inner_text().strip()

    def get_prior_home_coverage_current(self) -> str:
        return self.page.locator(self._LBL_PRIOR_COV_CUR).inner_text().strip()

    def get_prior_home_expiration_current(self) -> str:
        return self.page.locator(self._LBL_PRIOR_EXP_CUR).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Verification helpers                                                 #
    # ------------------------------------------------------------------ #

    def verify_insured_name(self, name: str) -> None:
        expect(self.page.locator(self._LBL_INSURED_NAME_CUR)).to_have_text(name)

    def verify_mailing_address(self, street: str, city_state_zip: str) -> None:
        expect(self.page.locator(self._LBL_STREET_CUR)).to_have_text(street)
        expect(self.page.locator(self._LBL_CITY_STATE_ZIP_CUR)).to_have_text(city_state_zip)

    def print_insured_summary(self) -> None:
        print(f"Insured:       {self.get_insured_name_current()}")
        print(f"Type:          {self.get_insured_type_current()}")
        print(f"DOB:           {self.get_dob_current()}")
        print(f"Address:       {self.get_street_current()}, {self.get_city_state_zip_current()}")
        print(f"Email:         {self.get_email_current()}")
        print(f"Primary Phone: {self.get_primary_phone_current()}")

    # ------------------------------------------------------------------ #
    # Edit                                                                 #
    # ------------------------------------------------------------------ #

    def click_edit(self) -> None:
        self.page.locator(self._BTN_EDIT_TOP).click()

    # ------------------------------------------------------------------ #
    # Popup: Invite to Paperless                                           #
    # ------------------------------------------------------------------ #

    def verify_paperless_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_PAPERLESS)).to_be_visible()

    def cancel_paperless_invite(self) -> None:
        self.page.locator(self._BTN_PAPERLESS_CANCEL).click()

    def confirm_and_send_paperless(self) -> None:
        self.page.locator(self._BTN_PAPERLESS_CONFIRM).click()

    # ------------------------------------------------------------------ #
    # Popup: Paperless Confirmation                                        #
    # ------------------------------------------------------------------ #

    def verify_paperless_confirmation_visible(self) -> None:
        expect(self.page.locator(self._PNL_PAPERLESS_CONF)).to_be_visible()

    def get_paperless_confirmation_message(self) -> str:
        return self.page.locator(self._LBL_PAPERLESS_CONF_MSG).inner_text().strip()

    def close_paperless_confirmation(self) -> None:
        self.page.locator(self._BTN_PAPERLESS_CONF_OK).click()

    # ------------------------------------------------------------------ #
    # Popup: Policy/Registered Name Warning                                #
    # ------------------------------------------------------------------ #

    def verify_name_warning_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_NAME_WARNING)).to_be_visible()

    def cancel_name_warning(self) -> None:
        self.page.locator(self._BTN_WARNING_CANCEL).click()

    def continue_past_name_warning(self) -> None:
        self.page.locator(self._BTN_WARNING_CONT).click()


class FGInsuredRegistrationForm:
    _PREFIX = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_"

    _PNL_REGISTRATION    = f"{_PREFIX}pnlRegistration"
    _TXT_EMAIL           = f"{_PREFIX}txtEmailAddress"
    _TXT_EMAIL_CONFIRM   = f"{_PREFIX}txtEmailAddressConfirm"
    _TXT_PRIMARY_AREA    = f"{_PREFIX}txtPrimaryPhoneArea"
    _TXT_PRIMARY_EXCH    = f"{_PREFIX}txtPrimaryPhoneExchange"
    _TXT_PRIMARY_NUM     = f"{_PREFIX}txtPrimaryPhoneNumber"
    _CHK_PRIMARY_MOBILE  = f"{_PREFIX}chkPrimaryIsMobile"
    _TXT_SEC_AREA        = f"{_PREFIX}txtSecondaryPhoneArea"
    _TXT_SEC_EXCH        = f"{_PREFIX}txtSecondaryPhoneExchange"
    _TXT_SEC_NUM         = f"{_PREFIX}txtSecondaryPhoneNumber"
    _CHK_SEC_MOBILE      = f"{_PREFIX}chkSecondaryIsMobile"
    _BTN_SUBMIT          = f"{_PREFIX}btnSubmit"

    def __init__(self, page: Page):
        self.page = page

    def wait_for_form(self) -> None:
        self.page.wait_for_selector(self._PNL_REGISTRATION, timeout=30000)

    def fill_email(self, email: str) -> None:
        self.page.locator(self._TXT_EMAIL).fill(email)
        self.page.locator(self._TXT_EMAIL_CONFIRM).fill(email)

    def fill_primary_phone(self, area: str, exchange: str, number: str, is_mobile: bool = False) -> None:
        self.page.locator(self._TXT_PRIMARY_AREA).fill(area)
        self.page.locator(self._TXT_PRIMARY_EXCH).fill(exchange)
        self.page.locator(self._TXT_PRIMARY_NUM).fill(number)
        if is_mobile:
            self.page.locator(self._CHK_PRIMARY_MOBILE).check()

    def fill_secondary_phone(self, area: str, exchange: str, number: str, is_mobile: bool = False) -> None:
        self.page.locator(self._TXT_SEC_AREA).fill(area)
        self.page.locator(self._TXT_SEC_EXCH).fill(exchange)
        self.page.locator(self._TXT_SEC_NUM).fill(number)
        if is_mobile:
            self.page.locator(self._CHK_SEC_MOBILE).check()

    def submit(self) -> None:
        self.page.locator(self._BTN_SUBMIT).click()

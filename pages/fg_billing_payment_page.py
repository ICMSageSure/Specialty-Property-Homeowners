from playwright.sync_api import Page, expect


class FGBillingPaymentPage:
    _PREFIX = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_"

    # -- Policy header --
    _POLICY_NUMBER  = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyNumber"
    _POLICY_INSURED = "#ctl00_ctl00_ContentPlaceHolder1_lblPolicyInsured"
    _LBL_POLICY_NUM = f"{_PREFIX}lblPolicyNumber"
    _LBL_BILL_TO    = f"{_PREFIX}lblBillTo"

    # -- Billing summary --
    _LBL_CURRENT_DUE     = f"{_PREFIX}lblCurrentDue"
    _LBL_AMOUNT_DUE_ON   = f"{_PREFIX}lblAmountDueOn"
    _LBL_LAST_PAYMENT    = f"{_PREFIX}lblLastPayment"
    _LBL_PAY_PLAN        = f"{_PREFIX}lblPayPlan"
    _LBL_REMAINING_BAL   = f"{_PREFIX}lblRemainingBalance"
    _LNK_VIEW_REMAINING  = f"{_PREFIX}lnkViewRemainingBalance"

    # -- Auto Pay status --
    _LBL_AUTOPAY_STATUS  = f"{_PREFIX}lblIsAutoPay"
    _LBL_AUTOPAY_ENROLL  = f"{_PREFIX}lblAutoPayEnrollMessage"
    _LBL_AUTOPAY_OFF     = f"{_PREFIX}lblAutoPayOffMessage"

    # -- Primary action buttons --
    _BTN_PAY_PLAN_OPT    = f"{_PREFIX}btnPayPlanOptions1"
    _BTN_ENROLL_AUTOPAY  = f"{_PREFIX}btnEnrollAutoPay"

    # -- History tabs --
    _BTN_BILLING_TAB      = "#btnShowBillingHistory"
    _BTN_TRANSACTION_TAB  = "#btnShowTransactionHistory"

    # -- Billing history controls --
    _DDL_BILLING_SORT    = f"{_PREFIX}ddlBillingSort"
    _DDL_BILLING_FILTER  = f"{_PREFIX}ddlBillingFilter"
    _BTN_RECENT_BILLING  = f"{_PREFIX}btnRecentBilling"
    _BTN_ALL_BILLING     = f"{_PREFIX}btnAllBilling"
    _PNL_BILLING_HISTORY = "#pnlBillingHistory"
    _TBL_BILLING_HISTORY = "#tblHistory"
    _LBL_NO_RECORDS      = f"{_PREFIX}noBillingRecordsLabel"

    # -- Transaction history controls --
    _DDL_TXN_SORT        = f"{_PREFIX}ddlTransactionSort"
    _DDL_TXN_FILTER      = f"{_PREFIX}ddlTransactionFilter"
    _BTN_RECENT_TXN      = f"{_PREFIX}btnRecentTransactions"
    _BTN_ALL_TXN         = f"{_PREFIX}btnAllTransactions"

    # -- Make Payment popup --
    _PNL_MAKE_PAYMENT    = f"{_PREFIX}pnlMakePaymentPopup"
    _BTN_PAY_CLOSE       = "#btnClose"
    _BTN_PAY_CONTINUE    = "#btnContinue"

    # -- Remaining Installments popup --
    # ID appears multiple times on the page (hidden copies); scope all to the visible popup panel
    _PNL_INSTALLMENTS       = f"{_PREFIX}pnlRemainingInstallmentsPopup"
    _BTN_INSTALL_CLOSE      = "input[name='ctl00$ctl00$ContentPlaceHolder1$PolicyContentPlaceHolder2$ucPayPlanOptionsPopup$ucPayPlanListGrid$btnCloseInstallments']"
    _TBL_INSTALLMENTS       = f"#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_pnlRemainingInstallmentsPopup #dgAvailableInstallments"
    _LBL_INSTALL_AMOUNT     = "#InstallAmountLable"
    _LBL_INSTALL_DISCLAIMER = "#lblSubject"
    _LBL_CUSTOMER_PHONE     = "#lblCustomerPhoneNumber2"

    # -- Edit/Enroll Auto Pay popup --
    _PNL_EDIT_AUTOPAY    = f"{_PREFIX}pnlEditAutoPayPopup"
    _BTN_AUTOPAY_CLOSE   = "#btnCloseEditAutoPay"
    _BTN_AUTOPAY_CONT    = "#btnContinueEditAutoPay"

    # -- Cancel Auto Pay popup --
    _PNL_CANCEL_AUTOPAY  = f"{_PREFIX}pnlCancelAutoPayPopup"
    _BTN_CANCEL_AP_CLOSE = f"{_PREFIX}btnCloseCancelAutoPay"
    _BTN_CANCEL_AP_CONT  = f"{_PREFIX}btnContinueCancelAutoPay"

    # -- Manage Payments popup --
    _PNL_MANAGE_PAY      = f"{_PREFIX}pnlManagePaymentsPopup"
    _BTN_MANAGE_CLOSE    = "#btnCloseManagePayments"
    _BTN_MANAGE_CONT     = "#btnContinueManagePayments"

    # -- Edit Pay Plan popup --
    _PNL_EDIT_PAY_PLAN   = f"{_PREFIX}pnlEditPayPlanPopup"
    _LNK_PAY_PLAN_CANCEL = "#lnkCancel"

    def __init__(self, page: Page):
        self.page = page

    # ------------------------------------------------------------------ #
    # Page load                                                            #
    # ------------------------------------------------------------------ #

    def wait_for_page(self) -> None:
        self.page.wait_for_selector(self._LBL_CURRENT_DUE, timeout=30000)

    # ------------------------------------------------------------------ #
    # Policy header                                                        #
    # ------------------------------------------------------------------ #

    def get_policy_number(self) -> str:
        return self.page.locator(self._POLICY_NUMBER).inner_text().strip()

    def get_insured_name(self) -> str:
        return self.page.locator(self._POLICY_INSURED).inner_text().strip()

    def get_bill_to(self) -> str:
        return self.page.locator(self._LBL_BILL_TO).inner_text().strip()

    # ------------------------------------------------------------------ #
    # Billing summary                                                      #
    # ------------------------------------------------------------------ #

    def get_current_due(self) -> str:
        return self.page.locator(self._LBL_CURRENT_DUE).inner_text().strip()

    def get_amount_due_status(self) -> str:
        return self.page.locator(self._LBL_AMOUNT_DUE_ON).inner_text().strip()

    def get_last_payment(self) -> str:
        return self.page.locator(self._LBL_LAST_PAYMENT).inner_text().strip()

    def get_pay_plan(self) -> str:
        return self.page.locator(self._LBL_PAY_PLAN).inner_text().strip()

    def get_remaining_balance(self) -> str:
        return self.page.locator(self._LBL_REMAINING_BAL).inner_text().strip()

    def click_view_remaining_balance(self) -> None:
        self.page.locator(self._LNK_VIEW_REMAINING).click()

    def verify_current_due(self, amount: str) -> None:
        expect(self.page.locator(self._LBL_CURRENT_DUE)).to_have_text(amount)

    def verify_pay_plan(self, plan: str) -> None:
        expect(self.page.locator(self._LBL_PAY_PLAN)).to_have_text(plan)

    def print_billing_summary(self) -> None:
        print(f"Current Due:       {self.get_current_due()}")
        print(f"Status:            {self.get_amount_due_status()}")
        print(f"Last Payment:      {self.get_last_payment()}")
        print(f"Pay Plan:          {self.get_pay_plan()}")
        print(f"Remaining Balance: {self.get_remaining_balance()}")

    # ------------------------------------------------------------------ #
    # Auto Pay                                                             #
    # ------------------------------------------------------------------ #

    def get_autopay_status(self) -> str:
        return self.page.locator(self._LBL_AUTOPAY_STATUS).inner_text().strip()

    def is_autopay_on(self) -> bool:
        return self.get_autopay_status().upper() == "ON"

    def click_enroll_autopay(self) -> None:
        self.page.locator(self._BTN_ENROLL_AUTOPAY).click()

    # ------------------------------------------------------------------ #
    # Pay Plan Options                                                     #
    # ------------------------------------------------------------------ #

    def click_pay_plan_options(self) -> None:
        self.page.locator(self._BTN_PAY_PLAN_OPT).click()

    # ------------------------------------------------------------------ #
    # History tabs                                                         #
    # ------------------------------------------------------------------ #

    def show_billing_history(self) -> None:
        self.page.locator(self._BTN_BILLING_TAB).click()

    def show_transaction_history(self) -> None:
        self.page.locator(self._BTN_TRANSACTION_TAB).click()

    # ------------------------------------------------------------------ #
    # Billing history controls                                             #
    # ------------------------------------------------------------------ #

    def select_billing_sort(self, value: str) -> None:
        self.page.locator(self._DDL_BILLING_SORT).select_option(value)

    def select_billing_filter(self, value: str) -> None:
        self.page.locator(self._DDL_BILLING_FILTER).select_option(value)

    def click_recent_billing(self) -> None:
        self.page.locator(self._BTN_RECENT_BILLING).click()

    def click_see_all_billing(self) -> None:
        self.page.locator(self._BTN_ALL_BILLING).click()

    def verify_no_billing_records(self) -> None:
        expect(self.page.locator(self._LBL_NO_RECORDS)).to_be_visible()

    # ------------------------------------------------------------------ #
    # Transaction history controls                                         #
    # ------------------------------------------------------------------ #

    def select_transaction_sort(self, value: str) -> None:
        self.page.locator(self._DDL_TXN_SORT).select_option(value)

    def select_transaction_filter(self, value: str) -> None:
        self.page.locator(self._DDL_TXN_FILTER).select_option(value)

    def click_recent_transactions(self) -> None:
        self.page.locator(self._BTN_RECENT_TXN).click()

    def click_see_all_transactions(self) -> None:
        self.page.locator(self._BTN_ALL_TXN).click()

    # ------------------------------------------------------------------ #
    # Make Payment popup                                                   #
    # ------------------------------------------------------------------ #

    def verify_make_payment_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_MAKE_PAYMENT)).to_be_visible()

    def close_make_payment_popup(self) -> None:
        self.page.locator(self._BTN_PAY_CLOSE).click()

    def agree_and_continue_payment(self) -> None:
        self.page.locator(self._BTN_PAY_CONTINUE).click()

    # ------------------------------------------------------------------ #
    # Remaining Installments popup                                         #
    # ------------------------------------------------------------------ #

    def verify_installments_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_INSTALLMENTS)).to_be_visible()

    def close_installments_popup(self) -> None:
        self.page.locator(self._BTN_INSTALL_CLOSE).click()

    def get_install_amount(self) -> str:
        return self.page.locator(self._LBL_INSTALL_AMOUNT).inner_text().strip()

    def get_customer_service_phone(self) -> str:
        return self.page.locator(self._LBL_CUSTOMER_PHONE).inner_text().strip()

    def get_installments_row_count(self) -> int:
        # counts data rows only (skips the header row)
        return self.page.locator(f"{self._TBL_INSTALLMENTS} tbody tr").count() - 1

    def verify_pay_plan_label(self, label: str) -> None:
        expect(self.page.locator(f"{self._TBL_INSTALLMENTS} #lblTotal")).to_have_text(label)

    # ------------------------------------------------------------------ #
    # Enroll / Edit Auto Pay popup                                         #
    # ------------------------------------------------------------------ #

    def verify_autopay_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_EDIT_AUTOPAY)).to_be_visible()

    def close_autopay_popup(self) -> None:
        self.page.locator(self._BTN_AUTOPAY_CLOSE).click()

    def agree_and_continue_autopay(self) -> None:
        self.page.locator(self._BTN_AUTOPAY_CONT).click()

    # ------------------------------------------------------------------ #
    # Cancel Auto Pay popup                                                #
    # ------------------------------------------------------------------ #

    def verify_cancel_autopay_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_CANCEL_AUTOPAY)).to_be_visible()

    def close_cancel_autopay_popup(self) -> None:
        self.page.locator(self._BTN_CANCEL_AP_CLOSE).click()

    def confirm_cancel_autopay(self) -> None:
        self.page.locator(self._BTN_CANCEL_AP_CONT).click()

    # ------------------------------------------------------------------ #
    # Manage Payments popup                                                #
    # ------------------------------------------------------------------ #

    def verify_manage_payments_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_MANAGE_PAY)).to_be_visible()

    def close_manage_payments_popup(self) -> None:
        self.page.locator(self._BTN_MANAGE_CLOSE).click()

    def agree_and_continue_manage_payments(self) -> None:
        self.page.locator(self._BTN_MANAGE_CONT).click()

    # ------------------------------------------------------------------ #
    # Edit Pay Plan popup                                                  #
    # ------------------------------------------------------------------ #

    def verify_edit_pay_plan_popup_visible(self) -> None:
        expect(self.page.locator(self._PNL_EDIT_PAY_PLAN)).to_be_visible()

    def cancel_edit_pay_plan(self) -> None:
        self.page.locator(self._LNK_PAY_PLAN_CANCEL).click()

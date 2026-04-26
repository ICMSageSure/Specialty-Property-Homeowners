import os
from playwright.sync_api import Page, expect
from pages.login_page import LoginPage
from pages.fg_policy_overview_page import FGPolicyOverviewPage
from pages.fg_billing_payment_page import FGBillingPaymentPage

USERNAME     = "icortesmarquez@geovera"
POLICY_NUM   = "GC50018511"


def _open_billing_page(page: Page) -> tuple[Page, FGBillingPaymentPage]:
    """Log in, search for the test policy, open in popup, navigate to Billing/Payment."""
    LoginPage(page).login(USERNAME, os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Policies").click()
    page.get_by_text("Select Product").click()
    page.get_by_role("tree").get_by_text("Homeowners").click()
    page.wait_for_timeout(2000)
    page.get_by_text("Policy #").first.click()
    page.get_by_role("option", name="Policy #").click()
    page.get_by_role("textbox", name="Search").fill(POLICY_NUM)
    page.get_by_role("button", name="View").click()
    with page.expect_popup() as popup_info:
        page.get_by_role("gridcell", name=POLICY_NUM).click()
    policy_page = popup_info.value
    FGPolicyOverviewPage(policy_page).go_to_billing_payment()
    billing = FGBillingPaymentPage(policy_page)
    billing.wait_for_page()
    return policy_page, billing


# ------------------------------------------------------------------ #
# Page load & summary                                                  #
# ------------------------------------------------------------------ #

def test_billing_page_loads(page: Page, snap) -> None:
    """Billing/Payment page opens and policy header is correct."""
    policy_page, billing = _open_billing_page(page)
    snap(policy_page, "billing_loaded")

    assert billing.get_policy_number() == POLICY_NUM
    assert billing.get_insured_name() != ""


def test_billing_summary_fields_present(page: Page, snap) -> None:
    """All key summary labels are present and non-empty."""
    policy_page, billing = _open_billing_page(page)

    expect(policy_page.locator(billing._LBL_CURRENT_DUE)).to_be_visible()
    expect(policy_page.locator(billing._LBL_AMOUNT_DUE_ON)).to_be_visible()
    expect(policy_page.locator(billing._LBL_LAST_PAYMENT)).to_be_visible()
    expect(policy_page.locator(billing._LBL_PAY_PLAN)).to_be_visible()
    snap(policy_page, "billing_summary")


def test_billing_summary_print(page: Page) -> None:
    """Smoke test: print billing summary to console (no assertion failures = pass)."""
    _, billing = _open_billing_page(page)
    billing.print_billing_summary()


def test_bill_to_displayed(page: Page) -> None:
    """Bill-To label is visible and non-empty."""
    _, billing = _open_billing_page(page)
    assert billing.get_bill_to() != ""


# ------------------------------------------------------------------ #
# Auto Pay                                                             #
# ------------------------------------------------------------------ #

def test_autopay_status_displayed(page: Page, snap) -> None:
    """Auto Pay status (ON or OFF) is displayed."""
    policy_page, billing = _open_billing_page(page)
    status = billing.get_autopay_status()
    snap(policy_page, "autopay_status")
    assert status.upper() in ("ON", "OFF"), f"Unexpected Auto Pay status: {status!r}"


# ------------------------------------------------------------------ #
# History tabs                                                         #
# ------------------------------------------------------------------ #

def test_billing_history_tab(page: Page, snap) -> None:
    """Clicking Billing History tab shows the billing history panel."""
    policy_page, billing = _open_billing_page(page)
    billing.show_billing_history()
    policy_page.wait_for_timeout(1000)
    snap(policy_page, "billing_history_tab")
    expect(policy_page.locator(billing._PNL_BILLING_HISTORY)).to_be_visible()


def test_transaction_history_tab(page: Page, snap) -> None:
    """Clicking Transaction History tab switches the view."""
    policy_page, billing = _open_billing_page(page)
    billing.show_billing_history()
    billing.show_transaction_history()
    policy_page.wait_for_timeout(1000)
    snap(policy_page, "transaction_history_tab")
    expect(policy_page.locator("#pnlTransactionHistory")).to_be_visible()


def test_billing_history_recent_and_all(page: Page, snap) -> None:
    """Recent and See All billing buttons are clickable without error."""
    policy_page, billing = _open_billing_page(page)
    billing.show_billing_history()
    policy_page.wait_for_timeout(500)
    billing.click_see_all_billing()
    policy_page.wait_for_timeout(1000)
    snap(policy_page, "billing_see_all")
    billing.click_recent_billing()
    policy_page.wait_for_timeout(1000)
    snap(policy_page, "billing_recent")


# ------------------------------------------------------------------ #
# Remaining Installments popup (Payment Plan Options sidebar link)    #
# ------------------------------------------------------------------ #

def test_remaining_installments_popup_opens(page: Page, snap) -> None:
    """Hovering Billing/Payment nav and clicking Payment Plan Options opens the popup."""
    policy_page, billing = _open_billing_page(page)
    overview = FGPolicyOverviewPage(policy_page)
    policy_page.locator("#ctl00_ctl00_ContentPlaceHolder1_lnkBillingPayment").hover()
    policy_page.wait_for_timeout(500)
    overview.go_to_payment_plan_options()
    policy_page.wait_for_timeout(1000)
    snap(policy_page, "installments_popup_open")
    billing.verify_installments_popup_visible()


def test_remaining_installments_table_visible(page: Page, snap) -> None:
    """Installments popup shows the available payment plans table."""
    policy_page, billing = _open_billing_page(page)
    overview = FGPolicyOverviewPage(policy_page)
    policy_page.locator("#ctl00_ctl00_ContentPlaceHolder1_lnkBillingPayment").hover()
    policy_page.wait_for_timeout(500)
    overview.go_to_payment_plan_options()
    policy_page.wait_for_timeout(1000)
    billing.verify_installments_popup_visible()
    expect(policy_page.locator(billing._TBL_INSTALLMENTS)).to_be_visible()
    snap(policy_page, "installments_table")


def test_remaining_installments_popup_closes(page: Page, snap) -> None:
    """Cancel button closes the Remaining Installments popup."""
    policy_page, billing = _open_billing_page(page)
    overview = FGPolicyOverviewPage(policy_page)
    policy_page.locator("#ctl00_ctl00_ContentPlaceHolder1_lnkBillingPayment").hover()
    policy_page.wait_for_timeout(500)
    overview.go_to_payment_plan_options()
    policy_page.wait_for_timeout(1000)
    billing.verify_installments_popup_visible()
    billing.close_installments_popup()
    policy_page.wait_for_timeout(500)
    snap(policy_page, "installments_popup_closed")
    expect(policy_page.locator(billing._PNL_INSTALLMENTS)).to_be_hidden()


# ------------------------------------------------------------------ #
# Pay Plan Options button (on page body)                              #
# ------------------------------------------------------------------ #

def test_pay_plan_options_button(page: Page, snap) -> None:
    """Pay Plan Options button opens the Remaining Installments popup (Full Pay policy)."""
    policy_page, billing = _open_billing_page(page)
    billing.click_pay_plan_options()
    policy_page.wait_for_timeout(1000)
    snap(policy_page, "pay_plan_options_popup")
    billing.verify_installments_popup_visible()

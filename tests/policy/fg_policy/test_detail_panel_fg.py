import os
from playwright.sync_api import Page, expect
from pages.login_page import LoginPage
from pages.fg_policy_detail_panel_page import FGPolicyDetailPanelPage

USERNAME   = "icortesmarquez@geovera"
POLICY_NUM = "GC50018511"
STATE      = "FL"


def _open_detail_panel(page: Page) -> tuple[Page, FGPolicyDetailPanelPage]:
    """Log in, search for the test policy, open in popup, return the detail panel page object."""
    LoginPage(page).login(USERNAME, os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Policies").click()
    page.get_by_text("Select Product").click()
    page.get_by_role("tree").get_by_text("Homeowners").click()
    page.wait_for_timeout(2000)
    page.get_by_text("Policy #").first.click()
    page.get_by_role("option", name="Policy #").click()
    page.get_by_role("textbox", name="Search").fill(POLICY_NUM)
    page.get_by_text("Select State...").first.click()
    page.wait_for_timeout(500)
    page.get_by_role("option", name=STATE).click()
    page.get_by_role("button", name="View").click()
    with page.expect_popup() as popup_info:
        page.get_by_role("gridcell", name=POLICY_NUM).click()
    policy_page = popup_info.value
    panel = FGPolicyDetailPanelPage(policy_page)
    panel.wait_for_page()
    return policy_page, panel


# ------------------------------------------------------------------ #
# Page load & policy header                                            #
# ------------------------------------------------------------------ #

def test_detail_panel_loads(page: Page, snap) -> None:
    """Overview page opens and policy header shows the correct policy number."""
    policy_page, panel = _open_detail_panel(page)
    snap(policy_page, "detail_panel_loaded")
    assert panel.get_policy_number() == POLICY_NUM
    assert panel.get_insured_name() != ""


# ------------------------------------------------------------------ #
# Policy details — left pane                                           #
# ------------------------------------------------------------------ #

def test_property_address_displayed(page: Page, snap) -> None:
    """Property address and city/state/zip are present and non-empty."""
    policy_page, panel = _open_detail_panel(page)
    snap(policy_page, "property_address")
    assert panel.get_property_address() != ""
    assert panel.get_property_city_state_zip() != ""


def test_mailing_address_displayed(page: Page, snap) -> None:
    """Mailing address and city/state/zip are present and non-empty."""
    policy_page, panel = _open_detail_panel(page)
    snap(policy_page, "mailing_address")
    assert panel.get_mailing_address() != ""
    assert panel.get_mailing_city_state_zip() != ""


def test_expiration_date_displayed(page: Page) -> None:
    """Expiration date label is present and non-empty."""
    _, panel = _open_detail_panel(page)
    assert panel.get_expiration_date() != ""


# ------------------------------------------------------------------ #
# Policy status — center pane                                          #
# ------------------------------------------------------------------ #

def test_product_displayed(page: Page, snap) -> None:
    """Product label is present and non-empty."""
    policy_page, panel = _open_detail_panel(page)
    snap(policy_page, "policy_status_pane")
    assert panel.get_product() != ""


def test_policy_status_is_valid(page: Page) -> None:
    """Policy status is one of the known valid values."""
    _, panel = _open_detail_panel(page)
    status = panel.get_policy_status()
    assert status in ("Active", "Cancelled", "Expired", "Pending"), (
        f"Unexpected policy status: {status!r}"
    )


def test_last_activity_displayed(page: Page) -> None:
    """Last activity description and date are present and non-empty."""
    _, panel = _open_detail_panel(page)
    assert panel.get_last_activity() != ""
    assert panel.get_last_activity_date() != ""


# ------------------------------------------------------------------ #
# Payment summary — right pane                                         #
# ------------------------------------------------------------------ #

def test_outstanding_balance_displayed(page: Page, snap) -> None:
    """Outstanding balance label is present and non-empty."""
    policy_page, panel = _open_detail_panel(page)
    snap(policy_page, "payment_summary")
    assert panel.get_outstanding_balance() != ""


def test_amount_due_displayed(page: Page) -> None:
    """Amount due label is present and non-empty."""
    _, panel = _open_detail_panel(page)
    assert panel.get_amount_due() != ""


# ------------------------------------------------------------------ #
# Paperless status panel                                               #
# ------------------------------------------------------------------ #

def test_paperless_status_displayed(page: Page, snap) -> None:
    """Paperless status label is present and non-empty."""
    policy_page, panel = _open_detail_panel(page)
    snap(policy_page, "paperless_status")
    assert panel.get_paperless_status() != ""


# ------------------------------------------------------------------ #
# Diagnostic smoke test                                                #
# ------------------------------------------------------------------ #

def test_print_detail_panel_summary(page: Page) -> None:
    """Smoke test: print all detail panel fields to console (no assertion failures = pass)."""
    _, panel = _open_detail_panel(page)
    panel.print_detail_panel_summary()

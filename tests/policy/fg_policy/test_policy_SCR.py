import os
import pytest
from playwright.sync_api import Page
from datetime import datetime, timedelta
from pages.login_page import LoginPage
from pages.policy_request_changes_page import PolicyRequestChangesPage, PolicyChangesConfirmationPage

USERNAME = "icortesmarquez@geovera"
POLICY_NUMBER = "GC60015733"
FILE_PATH = "C:/Users/icortesmarquez/Desktop/Playwright/HO_PW/test_data/testfile.pdf"

# (policy_number, change_type_value, label, first_name, last_name, email)
# Change types cycle through all 8 submittable types; Mortgagee (1) and Coverage (9) are info-only and excluded.
POLICY_CHANGE_MAP = [
    ("GC60015739", "0",  "Mailing Address",      "mailaddress",   "wbc",    "icmgeotester+mailingaddress@gmail.com"),
    ("GC60015722", "3",  "Named Insured",         "namedinsured",  "tester", "icmgeotester+namedinsured@gmail.com"),
    ("GC60015723", "4",  "Cancellation",          "cancellation",  "tester", "icmgeotester+cancellation@gmail.com"),
    ("GC60015724", "5",  "Property Information",  "property",      "tester", "icmgeotester+property@gmail.com"),
    ("GC60015725", "6",  "Mitigation/Renovation", "mitigation",    "tester", "icmgeotester+mitigation@gmail.com"),
    ("GC60015726", "7",  "Payment and Billing",   "payment",       "tester", "icmgeotester+payment@gmail.com"),
    ("GC60015727", "8",  "Other",                 "other",         "tester", "icmgeotester+other@gmail.com"),
    ("GC60015728", "11", "Reinstatement",         "reinstatement", "tester", "icmgeotester+reinstatement@gmail.com"),
    ("GC60015729", "0",  "Mailing Address",       "mailaddress",   "wbc",    "icmgeotester+mailingaddress@gmail.com"),
    ("GC60015730", "3",  "Named Insured",         "namedinsured",  "tester", "icmgeotester+namedinsured@gmail.com"),
    ("GC60015731", "4",  "Cancellation",          "cancellation",  "tester", "icmgeotester+cancellation@gmail.com"),
    ("GC60015732", "5",  "Property Information",  "property",      "tester", "icmgeotester+property@gmail.com"),
    ("GC60015737", "6",  "Mitigation/Renovation", "mitigation",    "tester", "icmgeotester+mitigation@gmail.com"),
    ("GC60015733", "7",  "Payment and Billing",   "payment",       "tester", "icmgeotester+payment@gmail.com"),
    ("GC60015734", "8",  "Other",                 "other",         "tester", "icmgeotester+other@gmail.com"),
    ("GC60015738", "11", "Reinstatement",         "reinstatement", "tester", "icmgeotester+reinstatement@gmail.com"),
    ("GC60015722", "0",  "Mailing Address",       "mailaddress",   "wbc",    "icmgeotester+mailingaddress@gmail.com"),
]


def _get_tomorrow() -> str:
    return (datetime.now() + timedelta(days=1)).strftime("%m/%d/%Y")


def _login_and_open_change_request(page: Page) -> Page:
    LoginPage(page).login(USERNAME, os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Policies").click()
    page.get_by_text("Select Product").click()
    page.get_by_role("tree").get_by_text("Homeowners").click()
    page.wait_for_timeout(2000)
    page.get_by_text("Select State...").first.click()
    page.wait_for_timeout(500)
    page.get_by_role("option", name="FL").click()
    page.get_by_text("Policy #").first.click()
    page.wait_for_timeout(500)
    page.get_by_role("option", name="Policy #").click()
    page.get_by_role("textbox", name="Search").fill(POLICY_NUMBER)
    page.get_by_role("button", name="View").click()
    with page.expect_popup(timeout=60000) as popup_info:
        page.get_by_role("gridcell", name=POLICY_NUMBER).click()
    page1 = popup_info.value
    page1.get_by_role("link", name="Request Changes").click()
    page1.wait_for_timeout(2000)
    return page1


def test_fg_policy_mailing_address(page: Page, snap) -> None:
    label = "Mailing Address"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("0")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("mailaddress", "wbc", "icmgeotester+mailingaddress@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


def test_fg_policy_mortgagee(page: Page, snap) -> None:
    """Mortgagee changes must be made from the Mortgagee Details page — no form here."""
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("1")
    snap(page1, "mortgagee_info_page")

    from playwright.sync_api import expect
    expect(page1.get_by_text("Mortgagee Details").last).to_be_visible()


def test_fg_policy_coverage_limits_and_deductibles(page: Page, snap) -> None:
    """Coverage changes must be made via the Premium Calculator tool — no form here."""
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("9")
    snap(page1, "coverage_info_page")

    from playwright.sync_api import expect
    expect(page1.get_by_text("Premium Calculator")).to_be_visible()


def test_fg_policy_named_insured(page: Page, snap) -> None:
    label = "Named Insured"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("3")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("namedinsured", "tester", "icmgeotester+namedinsured@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


def test_fg_policy_cancellation(page: Page, snap) -> None:
    label = "Cancellation"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("4")
    reqchg.fill_effective_date(_get_tomorrow())
    page1.wait_for_timeout(2000)
    reqchg.select_mailing_address_check("false")
    reqchg.select_cancel_reason("99")
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("cancellation", "tester", "icmgeotester+cancellation@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


def test_fg_policy_property_information(page: Page, snap) -> None:
    label = "Property Information"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("5")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("property", "tester", "icmgeotester+property@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


def test_fg_policy_mitigation_renovation(page: Page, snap) -> None:
    label = "Mitigation/Renovation"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("6")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("mitigation", "tester", "icmgeotester+mitigation@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


def test_fg_policy_payment_and_billing(page: Page, snap) -> None:
    label = "Payment and Billing"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("7")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("payment", "tester", "icmgeotester+payment@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


def test_fg_policy_other(page: Page, snap) -> None:
    label = "Other"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("8")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("other", "tester", "icmgeotester+other@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


def test_fg_policy_reinstatement(page: Page, snap) -> None:
    label = "Reinstatement"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("11")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact("reinstatement", "tester", "icmgeotester+reinstatement@gmail.com")
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()


# ---------------------------------------------------------------------------
# Batch run — each policy gets a different change type per POLICY_CHANGE_MAP
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("policy_num,change_type,label,first,last,email", POLICY_CHANGE_MAP)
def test_fg_batch_change(page: Page, snap, policy_num: str, change_type: str, label: str, first: str, last: str, email: str) -> None:
    comment = f"Test. Change Type = {label}."

    LoginPage(page).login(USERNAME, os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Policies").click()
    page.get_by_text("Select Product").click()
    page.get_by_role("tree").get_by_text("Homeowners").click()
    page.wait_for_timeout(2000)
    page.get_by_text("Select State...").first.click()
    page.wait_for_timeout(500)
    page.get_by_role("option", name="FL").click()
    page.get_by_text("Policy #").first.click()
    page.wait_for_timeout(500)
    page.get_by_role("option", name="Policy #").click()
    page.get_by_role("textbox", name="Search").fill(policy_num)
    page.get_by_role("button", name="View").click()
    with page.expect_popup(timeout=60000) as popup_info:
        page.get_by_role("gridcell", name=policy_num).click()
    page1 = popup_info.value
    page1.get_by_role("link", name="Request Changes").click()
    page1.wait_for_timeout(2000)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type(change_type)
    reqchg.fill_effective_date(_get_tomorrow())

    if change_type == "4":
        page1.wait_for_timeout(2000)
        reqchg.select_mailing_address_check("false")
        reqchg.select_cancel_reason("99")

    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    snap(page1, "fields_filled")
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    conf.fill_submitter_contact(first, last, email)
    conf.submit()
    conf.verify_success()
    snap(page1, "after_submit")
    conf.click_done()

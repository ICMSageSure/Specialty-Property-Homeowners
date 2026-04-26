import os
from playwright.sync_api import Page
from datetime import datetime, timedelta
from pages.login_page import LoginPage
from pages.policy_request_changes_page import PolicyRequestChangesPage, PolicyChangesConfirmationPage

USERNAME = "icortesmarquez@geovera"
POLICY_NUMBER = "GC50018717"
FILE_PATH = "C:/Users/icortesmarquez/Desktop/Playwright/HO_PW/test_data/testfile.pdf"


def _get_tomorrow() -> str:
    return (datetime.now() + timedelta(days=1)).strftime("%m/%d/%Y")


def _login_and_open_change_request(page: Page) -> Page:
    LoginPage(page).login(USERNAME, os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Policies").click()
    page.get_by_text("Select Product").click()
    page.get_by_role("tree").get_by_text("Homeowners").click()
    page.wait_for_timeout(2000)
    page.get_by_text("Select State...").first.click()
    page.get_by_role("option", name="FL").click()
    page.get_by_text("Policy #").first.click()
    page.get_by_role("option", name="Policy #").click()
    page.get_by_role("textbox", name="Search").fill(POLICY_NUMBER)
    page.get_by_role("button", name="View").click()
    with page.expect_popup() as popup_info:
        page.get_by_role("gridcell", name=POLICY_NUMBER).click()
    page1 = popup_info.value
    page1.get_by_role("link", name="Request Changes").click()
    return page1


def test_fg_policy_mailing_address(page: Page) -> None:
    label = "Mailing Address"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("0")
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill mailing address-specific confirmation fields
    conf.fill_submitter_contact("mailaddress", "wbc", "icmgeotester+mailingaddress@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_mortgagee(page: Page) -> None:
    label = "Mortgagee"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("1")
    # TODO: fill mortgagee fields (e.g. name, loan number, address)
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill mortgagee-specific confirmation fields
    conf.fill_submitter_contact("mortgagee", "tester", "icmgeotester+mortgagee@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_coverage_limits_and_deductibles(page: Page) -> None:
    label = "Coverage Limits and Deductibles"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("9")
    # TODO: fill coverage/deductible fields that appear after type selection
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill coverage-specific confirmation fields
    conf.fill_submitter_contact("coverage", "tester", "icmgeotester+coverage@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_named_insured(page: Page) -> None:
    label = "Named Insured"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("3")
    # TODO: fill named insured fields (e.g. add/remove/change, first/last name)
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill named insured confirmation fields
    conf.fill_submitter_contact("namedinsured", "tester", "icmgeotester+namedinsured@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_cancellation(page: Page) -> None:
    label = "Cancellation"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("4")
    # TODO: fill cancellation fields (e.g. reason)
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill cancellation-specific confirmation fields
    conf.fill_submitter_contact("cancellation", "tester", "icmgeotester+cancellation@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_property_information(page: Page) -> None:
    label = "Property Information"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("5")
    # TODO: fill property information fields that appear after type selection
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill property information confirmation fields
    conf.fill_submitter_contact("property", "tester", "icmgeotester+property@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_mitigation_renovation(page: Page) -> None:
    label = "Mitigation/Renovation"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("6")
    # TODO: fill mitigation/renovation fields that appear after type selection
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill mitigation/renovation confirmation fields
    conf.fill_submitter_contact("mitigation", "tester", "icmgeotester+mitigation@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_payment_and_billing(page: Page) -> None:
    label = "Payment and Billing"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("7")
    # TODO: fill payment/billing fields that appear after type selection
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill payment and billing confirmation fields
    conf.fill_submitter_contact("payment", "tester", "icmgeotester+payment@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_other(page: Page) -> None:
    label = "Other"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("8")
    # TODO: fill "other" fields that appear after type selection
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill other-specific confirmation fields
    conf.fill_submitter_contact("other", "tester", "icmgeotester+other@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()


def test_fg_policy_reinstatement(page: Page) -> None:
    label = "Reinstatement"
    comment = f"Test. Change Type = {label}."
    page1 = _login_and_open_change_request(page)

    reqchg = PolicyRequestChangesPage(page1)
    reqchg.select_change_type("11")
    # TODO: fill reinstatement fields that appear after type selection
    reqchg.fill_effective_date(_get_tomorrow())
    reqchg.fill_comment(comment)
    reqchg.upload_file(FILE_PATH)
    reqchg.click_continue()

    conf = PolicyChangesConfirmationPage(page1)
    conf.verify_header(label, comment)
    # TODO: fill reinstatement-specific confirmation fields
    conf.fill_submitter_contact("reinstatement", "tester", "icmgeotester+reinstatement@gmail.com")
    conf.submit()
    conf.verify_success()
    conf.click_done()

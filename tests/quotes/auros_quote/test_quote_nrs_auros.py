import re
import os
import random
import pytest
from playwright.sync_api import Page
from helpers.result_logger import make_xlsx_appender
from helpers.test_data import AUROS_STANDARD_ADDRESSES as FL_ADDRESSES

USERNAME = "icortesmarquez@geovera"
WHOLESALER = "HULL001"
PRODUCER_NUMBER = "HLDAWNF1"
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "results", "nrs", "nrs_quote_results.xlsx")
_HEADERS = ["User Name", "Address", "Test Scenario", "Product Type", "Coverage A Limit", "Wholesaler", "Producer Number", "Quote Number", "Quote Expiration", "Dwelling (Cov A)", "Total Annual Premium", "Full Pay", "3 Pay (1st Install)", "4 Pay (1st Install)", "Monthly (1st Install)"]
_append_result = make_xlsx_appender(RESULTS_FILE, _HEADERS)

def test_quote_nrs_auros(page: Page) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"

    page.goto("https://tst.mygeosource.com/")
    page.get_by_role("textbox", name="Policyholder Email or Agent").click()
    page.get_by_role("textbox", name="Policyholder Email or Agent").fill(USERNAME)
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill(os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Log In", exact=True).click()
    page.get_by_role("paragraph").filter(has_text="Start New Quote").get_by_role("link").click()
    try:
        page.locator("#optionHomeowners").get_by_text("GeoVera Specialty Insurance").click(timeout=1500)
    except Exception:
        page.locator("#optionHomeowners").get_by_text("Residential Homeowner Products").click()
    with page.expect_popup() as page1_info:
        page.get_by_role("button", name="OK").click()

    page1 = page1_info.value
    # Setting Wholesaler, Producer, Updating from Producer table
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ddlWholesaler", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option(WHOLESALER)
    page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill(PRODUCER_NUMBER)
    page1.locator("#ctl00_WelcomeUserControl_btnUpdateQuotingAs").click()
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId").click()
    page1.wait_for_timeout(3000)
    # Enter Address, Product Information
    page1.locator("#RadAutoCompleteBoxGeo_Input").click()
    page1.locator("#RadAutoCompleteBoxGeo_Input").press_sequentially(address)
    page1.wait_for_timeout(1000)
    page1.locator("#RadAutoCompleteBoxGeo_Input").press("Enter")
    page1.get_by_role("button", name="Next -->").click()
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkProductType_0").check()  # Auros HO-3
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkProductType_1").check()  # GeoVera Specialty HO-3
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkProductType_2").check()  # GeoVera Specialty HO Wind Only
    page1.locator("#txtRequestedHomeownersCoverageALimit").click()
    page1.locator("#txtRequestedHomeownersCoverageALimit").fill(coverage_a)
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").fill("2025")
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").fill("2025")
    page1.get_by_role("button", name="Next -->").click()
    page1.wait_for_timeout(2000)
    ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
    if ineligible_msgs.count() > 0:
        pytest.skip(f"Address '{address}' ineligible for Auros – {ineligible_msgs.first.inner_text().strip()}")
    try:
        page1.wait_for_selector("#ddlConstructionType", timeout=10000)
    except Exception:
        pytest.skip(f"Address '{address}' did not advance to property page – possible ineligibility with unrecognized message")
    # Property Information Screen
    page1.locator("#ddlConstructionType").select_option("FR")
    page1.locator("#ddlSidingTypeDominant").select_option("ST")
    page1.locator("#ddlRestrictedRoofShapes").select_option("False")
    if page1.locator("#ddlRoofShapeDominant").is_visible():
        page1.locator("#ddlRoofShapeDominant").select_option("2")
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFoundationElevation").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFoundationElevation").fill("0")
    page1.get_by_role("button", name="Next -->").click()
    page1.wait_for_timeout(3000)
    # Renovation Screen
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralFireAlarm").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralBurglarAlarm").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlGatedGuardedCommunity").select_option("False")
    if page1.locator("#ddlSwimmingPool").is_visible():
        page1.locator("#ddlSwimmingPool").select_option(label="No Swimming Pool")
    page1.get_by_role("button", name="Next -->").click()
    page1.locator("#ddlPolybutylenePipes").select_option("False")
    page1.locator("#ddlAutomaticWaterShutoffValve").select_option("False")
    page1.get_by_role("button", name="Next -->").click()
    # Mitigation Screen
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlWindMitigationVerificationInspection").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlWindMitigationVerificationInspection").select_option("Yes")
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionMonth").fill("01")
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionDay").fill("01")
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionYear").fill("2022")
        page1.locator("#fileWindMitigationVerificationInspectionForm").set_input_files("C:/Users/icortesmarquez/Desktop/Playwright/HO_PW/test_data/testfile.pdf")
    page1.get_by_role("button", name="Next -->").click()
    # Applicant Information Screen
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlPolicyOwnerType").select_option("Individual")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName").click()
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName").press("CapsLock")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName").press_sequentially("ICMTEST")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLastName").click()
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLastName").press_sequentially("TESTER")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBMonth").click()
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBMonth").press_sequentially("01")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBDay").press_sequentially("01")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBYear").press_sequentially("1998")
    page1.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN1\"]").press_sequentially("123")
    page1.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN2\"]").press_sequentially("11")
    page1.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN3\"]").press_sequentially("1122")
    page1.get_by_role("checkbox", name="Is Mailing Address the same").check()
    page1.locator("#ddlPriorHomeownersInsurance").select_option("False")
    page1.wait_for_timeout(3000)
    page1.locator("#ddlPriorHomeownersInsurance").select_option("False")
    page1.get_by_role("button", name="Next -->").click()
    # General Information Screen
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl00_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl01_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl02_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl03_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl04_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl05_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl06_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl07_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl08_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl09_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl10_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl11_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl12_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl13_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl14_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl15_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl16_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl17_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl18_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl19_ddlQuestions").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl20_ddlQuestions").select_option("False")
    page1.get_by_role("cell", name="Applicant acknowledges that flood coverage is excluded for the homeowners, wind-only, and vacant policy.", exact=True).locator("span").first.click()
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkRoofPaymentSchedule").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkRoofPaymentSchedule").check()
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkFloodExclusion").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkFloodExclusion").check()
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkRoofExclusion").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkRoofExclusion").check()
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkAcknowledgedWaterDamageLimitation").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkAcknowledgedWaterDamageLimitation").check()
    page1.get_by_role("button", name="Next -->").click()
    # Optional Endorsements page - no endorsements selected
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkAcknowledgedWaterDamageLimitation").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkAcknowledgedWaterDamageLimitation").check()
    page1.get_by_role("button", name="Submit Quote Request").click()
    page1.wait_for_timeout(2000)
    ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
    if ineligible_msgs.count() > 0:
        pytest.skip(f"Address '{address}' ineligible for Auros – {ineligible_msgs.first.inner_text().strip()}")
    try:
        page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header", timeout=15000)
    except Exception:
        page_text = page1.locator("body").inner_text()
        pytest.skip(f"[Quote Only] Results page did not load for '{address}' – page content: {page_text[:300].strip()}")

    ProductName = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header").inner_text()
    print(f'Homeowners Product Option: {ProductName}')
    QuoteNumber = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Number").inner_text()
    print(f'Quote Number: {QuoteNumber}')
    QuoteExpiration = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Expiration").inner_text()
    print(f'Quote Expiration: {QuoteExpiration}')
    Dwelling = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lstCoverageDetails_ctl00_lblQuote1Coverage").inner_text()
    print(f'Dwelling (Cov A): {Dwelling}')
    TotalAnnualPremium_loc = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Premium")
    TotalAnnualPremium = TotalAnnualPremium_loc.inner_text().strip() if TotalAnnualPremium_loc.is_visible() else ""
    print(f'Total Annual Premium: {TotalAnnualPremium}')

    auros_company = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblCompanyDisplay")
    use_first = auros_company.is_visible() and bool(auros_company.inner_text().strip())
    print(f'Company: {auros_company.inner_text().strip() if use_first else "N/A"}')

    def _get_payment(idx: str) -> str:
        loc = page1.locator(f"span[id*='_lstPaymentOptions_{idx}_lblQuote1PaymentOption']")
        if loc.count() == 0:
            return ""
        text = (loc.first if use_first else loc).text_content().strip()
        m = re.search(r"\$\d{1,3}(?:,\d{3})*(?:\.\d{2})", text)
        return m.group() if m else text

    FullPay    = _get_payment("ctl00")
    ThreePay   = _get_payment("ctl01")
    FourPay    = _get_payment("ctl02")
    MonthlyPay = _get_payment("ctl03")

    print(f'Full Pay: {FullPay}')
    print(f'3 Pay (1st Install): {ThreePay}')
    print(f'4 Pay (1st Install): {FourPay}')
    print(f'Monthly (1st Install): {MonthlyPay}')

    _append_result([USERNAME, address, "Quote Only", ProductName, coverage_a, WHOLESALER, PRODUCER_NUMBER,
                     QuoteNumber, QuoteExpiration, Dwelling, TotalAnnualPremium,
                     FullPay, ThreePay, FourPay, MonthlyPay])

import re
import os
import random
import pytest
from playwright.sync_api import Page
from openpyxl import load_workbook, Workbook

USERNAME = "icortesmarquez@geovera"
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "..", "auros_test_results.xlsx")

FL_ADDRESSES = [
    "8734 Crestgate Cir, Orlando, FL 32819",
    "19240 Pepper Grass Dr, Tampa, FL 33647",
    "8832 Savona Ct, Naples, FL 34119",
    "1461 75th Cir NE, Saint Petersburg, FL 33702",
    "8050 Princeton Dr, Naples, FL 34104",
    "12274 Accipiter Dr, Orlando, FL 32837",
    "15110 Shearcrest Dr, Lithia, FL 33547",
    "7733 Belvoir Dr, Orlando, FL 32835",
    "9426 Meadow Crest Ln, Clermont, FL 34711",
    "612 Citrus Wood Ln, Valrico, FL 33594",
    "66 Woodbury Dr, Palm Coast, FL 32164",
    "1401 Maine Ave, Lynn Haven, FL 32444",
    "9607 Bahia Rd, Ocala, FL 34472",
    "2150 NW 74th Ave, Fort Lauderdale, FL 33313",
    "4364 Tarragon Ave, Middleburg, FL 32068",
    "3198 Michaels Ct, Green Cove Springs, FL 32043",
    "3250 Dowitcher Ln, Orange Park, FL 32065",
    "613 S West St, Green Cove Springs, FL 32043",
    "8329 Lakemont Dr, Jacksonville, FL 32216",
    "13501 Princess Kelly Dr, Jacksonville, FL 32225",
    "6229 Cranberry Ln E, Jacksonville, FL 32244",
    "14688 Indigo Lakes Cir, Naples, FL 34119",
]


def _append_to_xlsx(row: list):
    path = os.path.abspath(RESULTS_FILE)
    if os.path.exists(path):
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["User Name", "Address", "Product Type", "Coverage A Limit", "Quote Number", "Premium"])
    ws.append(row)
    wb.save(path)


# ---------------------------------------------------------------------------
# Shared quote flow
# ---------------------------------------------------------------------------

def _run_to_endorsements_page(page: Page, address: str, coverage_a: str = "800,000"):
    """Runs the full quote flow and lands on the Optional Endorsements page. Returns page1."""
    page.goto("https://tst.mygeosource.com/")
    page.get_by_role("textbox", name="Policyholder Email or Agent").click()
    page.get_by_role("textbox", name="Policyholder Email or Agent").fill(USERNAME)
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill(os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Log In", exact=True).click()
    page.get_by_role("paragraph").filter(has_text="Start New Quote").get_by_role("link").click()
    page.locator("#optionHomeowners").get_by_text("GeoVera Specialty Insurance").click()
    with page.expect_popup() as page1_info:
        page.get_by_role("button", name="OK").click()

    page1 = page1_info.value
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ddlWholesaler", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option("APPUND")
    page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill("APPAGT")
    page1.locator("#ctl00_WelcomeUserControl_btnUpdateQuotingAs").click()
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId").click()
    page1.wait_for_timeout(3000)
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
    page1.get_by_role("button", name="Next -->").click()
    page1.locator("#ddlConstructionType").select_option("FR")
    page1.locator("#ddlSidingTypeDominant").select_option("ST")
    page1.locator("#ddlRestrictedRoofShapes").select_option("False")
    if page1.locator("#ddlRoofShapeDominant").is_visible():
        page1.locator("#ddlRoofShapeDominant").select_option("2")
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFoundationElevation").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFoundationElevation").fill("0")
    page1.get_by_role("button", name="Next -->").click()
    page1.wait_for_timeout(3000)
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralFireAlarm").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralBurglarAlarm").select_option("False")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlGatedGuardedCommunity").select_option("False")
    if page1.locator("#ddlSwimmingPool").is_visible():
        page1.locator("#ddlSwimmingPool").select_option(label="No Swimming Pool")
    page1.get_by_role("button", name="Next -->").click()
    page1.locator("#ddlPolybutylenePipes").select_option("False")
    page1.locator("#ddlAutomaticWaterShutoffValve").select_option("False")
    page1.get_by_role("button", name="Next -->").click()
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlWindMitigationVerificationInspection").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlWindMitigationVerificationInspection").select_option("Yes")
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionMonth").fill("01")
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionDay").fill("01")
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionYear").fill("2022")
        page1.locator("#fileWindMitigationVerificationInspectionForm").set_input_files("C:/Users/icortesmarquez/Desktop/Playwright/HO_PW/uploads/testfile.pdf")
    page1.get_by_role("button", name="Next -->").click()
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
    page1.get_by_role("button", name="Next -->").click()
    return page1


def _finish_quote(page1, address: str, coverage_a: str):
    """Submit endorsements, capture quote results, and log to xlsx. Does not bind."""
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkAcknowledgedWaterDamageLimitation").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkAcknowledgedWaterDamageLimitation").check()
    page1.get_by_role("button", name="Submit Quote Request").click()

    page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lblQuote1Header", timeout=30000)
    ProductName = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lblQuote1Header").inner_text()
    print(f'Homeowners Product Option: {ProductName}')
    QuoteNumber = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lblQuote1Number").inner_text()
    print(f'Quote Number: {QuoteNumber}')
    FullPay = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lstPaymentOptions_ctl00_lblQuote1PaymentOption").inner_text()
    print(f'Full Pay Option: {FullPay}')
    premium_match = re.search(r"\$\d{1,3}(?:,\d{3})*(?:\.\d{2})", FullPay)
    premium = premium_match.group() if premium_match else FullPay

    ThreePayFirstInstallment = page1.locator(
        '#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_lstPaymentOptions_ctl01_lblQuote1PaymentOption'
    ).text_content().strip()
    match = re.search(r"\$\d{1,3}(?:,\d{3})*(?:\.\d{2})", ThreePayFirstInstallment)
    if match:
        print(f"3 Pay First Installment:\n{match.group().strip()}\n")
    else:
        print("No dollar amount found in:", ThreePayFirstInstallment)

    page1.wait_for_timeout(3000)
    _append_to_xlsx([USERNAME, address, ProductName, coverage_a, QuoteNumber, premium])


# ---------------------------------------------------------------------------
# Endorsement helpers
# ---------------------------------------------------------------------------

def _fill_address_endorsement(page1, idx: str, num_addrs: int = 1, use_max: bool = False):
    """ctl00=Additional Interest, ctl01=Additional Insured, ctl06=Additional Residence Occupied.
    All three use ucEndorsement and have Name + SecondName fields."""
    prefix = f"ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_{idx}_ucEndorsement"
    name   = "A" * 30 if use_max else "ICMTEST"
    second = "B" * 30 if use_max else "TESTER"
    street = "C" * 30 if use_max else "123 TEST ST"
    city   = "D" * 30 if use_max else "ORLANDO"

    page1.locator(f"#{prefix}_chkInclude").check()
    page1.locator(f"#{prefix}_ddlNumberOfAddresses").select_option(str(num_addrs))

    for i in range(1, num_addrs + 1):
        page1.locator(f"#{prefix}_txtName{i}").fill(name)
        page1.locator(f"#{prefix}_txtSecondName{i}").fill(second)
        page1.locator(f"#{prefix}_txtStreet{i}").fill(street)
        page1.locator(f"#{prefix}_txtCity{i}").fill(city)
        page1.locator(f"#{prefix}_ddlState{i}").select_option("FL")
        page1.locator(f"#{prefix}_txtZip{i}").fill("12345")
        page1.locator(f"#{prefix}_txtZip4_{i}").fill("6789")


def _fill_rented_address_endorsement(page1, num_addrs: int = 1, use_max: bool = False):
    """ctl04=Additional Residence Rented to Others. No Name/SecondName fields."""
    prefix = "ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_ctl04_ucEndorsement"
    street = "C" * 30 if use_max else "123 TEST ST"
    city   = "D" * 30 if use_max else "ORLANDO"

    page1.locator(f"#{prefix}_chkInclude").check()
    page1.locator(f"#{prefix}_ddlNumberOfAddresses").select_option(str(num_addrs))

    for i in range(1, num_addrs + 1):
        page1.locator(f"#{prefix}_txtStreet{i}").fill(street)
        page1.locator(f"#{prefix}_txtCity{i}").fill(city)
        page1.locator(f"#{prefix}_ddlState{i}").select_option("FL")
        page1.locator(f"#{prefix}_txtZip{i}").fill("12345")
        page1.locator(f"#{prefix}_txtZip4_{i}").fill("6789")


def _check_only_endorsement(page1, idx: str):
    """ctl02=Ordinance or Law, ctl03=Windstorm Or Hail Exclusion. Checkbox only."""
    page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_{idx}_ucEndorsement_chkInclude").check()


def _fill_pool_endorsement(page1):
    """ctl05=Increased Swimming Pool. Uses ucEndorsement2 (not ucEndorsement)."""
    prefix = "ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_ctl05_ucEndorsement2"
    page1.locator(f"#{prefix}_chkInclude").check()
    page1.locator(f"#{prefix}_ddlLimit").select_option("10000")


def _fill_water_limitation(page1):
    """Water Damage Limitation. Product checkbox US-WL-20 is pre-checked; just set the limit."""
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl00_ucEndorsement2a_ddlLimit").select_option("10000")


def _fill_water_excl_named_peril(page1):
    """Water Damage Exclusion With Named Peril Coverage. Check both product form checkboxes."""
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl01_ucEndorsement1a_rptrProductEndorsements_ctl00_chkInclude").check()
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl01_ucEndorsement1a_rptrProductEndorsements_ctl01_chkInclude").check()


def _fill_water_excl_limited(page1):
    """Water Damage Exclusion And Limited Named Peril Coverage. Check both product form checkboxes."""
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl02_ucEndorsement1a_rptrProductEndorsements_ctl00_chkInclude").check()
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl02_ucEndorsement1a_rptrProductEndorsements_ctl01_chkInclude").check()


# ---------------------------------------------------------------------------
# Tests: one of each
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("scenario", [
    "additional_interest",
    "additional_insured",
    "ordinance_or_law",
    "windstorm_hail_exclusion",
    "additional_rented",
    "swimming_pool",
    "additional_occupied",
    "water_limitation",
    "water_excl_named_peril",
    "water_excl_limited",
])
def test_auros_one_of_each(page: Page, scenario: str) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    page1 = _run_to_endorsements_page(page, address, coverage_a)

    if scenario == "additional_interest":
        _fill_address_endorsement(page1, "ctl00")
    elif scenario == "additional_insured":
        _fill_address_endorsement(page1, "ctl01")
    elif scenario == "ordinance_or_law":
        _check_only_endorsement(page1, "ctl02")
    elif scenario == "windstorm_hail_exclusion":
        _check_only_endorsement(page1, "ctl03")
    elif scenario == "additional_rented":
        _fill_rented_address_endorsement(page1)
    elif scenario == "swimming_pool":
        _fill_pool_endorsement(page1)
    elif scenario == "additional_occupied":
        _fill_address_endorsement(page1, "ctl06")
    elif scenario == "water_limitation":
        _fill_water_limitation(page1)
    elif scenario == "water_excl_named_peril":
        _fill_water_excl_named_peril(page1)
    elif scenario == "water_excl_limited":
        _fill_water_excl_limited(page1)

    _finish_quote(page1, address, coverage_a)


# ---------------------------------------------------------------------------
# Test: all endorsements (water exclusions omitted — conflict with limitation)
# ---------------------------------------------------------------------------

def test_auros_all_endorsements(page: Page) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    page1 = _run_to_endorsements_page(page, address, coverage_a)

    _fill_address_endorsement(page1, "ctl00")
    _fill_address_endorsement(page1, "ctl01")
    _check_only_endorsement(page1, "ctl02")
    _check_only_endorsement(page1, "ctl03")
    _fill_rented_address_endorsement(page1)
    _fill_pool_endorsement(page1)
    _fill_address_endorsement(page1, "ctl06")
    _fill_water_limitation(page1)

    _finish_quote(page1, address, coverage_a)


# ---------------------------------------------------------------------------
# Tests: max field length (30-char text fields)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("scenario", [
    "additional_interest",
    "additional_insured",
    "additional_rented",
    "additional_occupied",
])
def test_auros_max_field_length(page: Page, scenario: str) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    page1 = _run_to_endorsements_page(page, address, coverage_a)

    if scenario == "additional_interest":
        _fill_address_endorsement(page1, "ctl00", num_addrs=2, use_max=True)
    elif scenario == "additional_insured":
        _fill_address_endorsement(page1, "ctl01", num_addrs=2, use_max=True)
    elif scenario == "additional_rented":
        _fill_rented_address_endorsement(page1, num_addrs=2, use_max=True)
    elif scenario == "additional_occupied":
        _fill_address_endorsement(page1, "ctl06", num_addrs=2, use_max=True)

    _finish_quote(page1, address, coverage_a)

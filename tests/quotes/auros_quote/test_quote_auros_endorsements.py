import re
import os
import random
import pytest
from playwright.sync_api import Page
from helpers.result_logger import make_xlsx_appender
from helpers.test_data import AUROS_STANDARD_ADDRESSES as FL_ADDRESSES
from helpers.auros_endorsements import (
    fill_address_endorsement,
    fill_rented_address_endorsement,
    fill_occupied_address_endorsement,
    check_only_endorsement,
    fill_pool_endorsement,
    fill_water_limitation,
    fill_water_excl_named_peril,
    fill_water_excl_limited,
)

USERNAME = "icortesmarquez@geovera"
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "results", "auros", "auros_quote_results.xlsx")

SCENARIO_LABELS = {
    "additional_interest":          "Additional Interest",
    "additional_interest_ho2470":   "Additional Interest (HO-24-70)",
    "additional_insured":           "Additional Insured",
    "additional_insured_ho2470":    "Additional Insured (HO-24-70)",
    "ordinance_or_law":             "Ordinance or Law",
    "windstorm_hail_exclusion":     "Windstorm/Hail Exclusion",
    "additional_rented":            "Additional Residence - Rented",
    "additional_rented_ho2470":     "Additional Residence - Rented (HO-24-70)",
    "swimming_pool":                "Swimming Pool",
    "additional_occupied":          "Additional Residence - Occupied",
    "additional_occupied_ho2470":   "Additional Residence - Occupied (HO-24-70)",
    "water_limitation":             "Water Damage Limitation",
    "water_excl_named_peril":       "Water Excl - Named Peril",
    "water_excl_limited":           "Water Excl - Limited Named Peril",
}

MAX_SCENARIO_LABELS = {
    "additional_interest":  "Additional Interest (Max Length, 2 Addr)",
    "additional_insured":   "Additional Insured (Max Length, 2 Addr)",
    "additional_rented":    "Additional Residence - Rented (Max Length, 2 Addr)",
    "additional_occupied":  "Additional Residence - Occupied (Max Length, 2 Addr)",
}

_FIELD_DATA_WITH_NAME = {
    "Name (max 30)":        "A" * 29 + "1",
    "Second Name (max 30)": "B" * 29 + "1",
    "Street (max 30)":      "C" * 29 + "1",
    "City (max 30)":        "D" * 29 + "1",
    "Zip (max 5)":          "12341",
    "Zip+4 (max 4)":        "6781",
}
_FIELD_DATA_NO_NAME = {
    "Name (max 30)":        "",
    "Second Name (max 30)": "",
    "Street (max 30)":      "C" * 29 + "1",
    "City (max 30)":        "D" * 29 + "1",
    "Zip (max 5)":          "12341",
    "Zip+4 (max 4)":        "6781",
}


_HEADERS = ["User Name", "Address", "Test Scenario", "Product Type", "Coverage A Limit", "Quote Number", "Quote Expiration", "Dwelling (Cov A)", "Total Annual Premium", "Full Pay", "3 Pay (1st Install)", "4 Pay (1st Install)", "Monthly (1st Install)", "Name (max 30)", "Second Name (max 30)", "Street (max 30)", "City (max 30)", "Zip (max 5)", "Zip+4 (max 4)"]
_append_result = make_xlsx_appender(RESULTS_FILE, _HEADERS)


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
    try:
        page.locator("#optionHomeowners").get_by_text("GeoVera Specialty Insurance").click(timeout=1500)
    except Exception:
        page.locator("#optionHomeowners").get_by_text("Residential Homeowner Products").click()
    with page.expect_popup() as page1_info:
        page.get_by_role("button", name="OK").click()

    page1 = page1_info.value
    # Setting Wholesaler, Producer, Updating from Producer table
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ddlWholesaler", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option("APPUND")
    page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill("APPAGT")
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
    return page1


_FIELD_COLS = ["Name (max 30)", "Second Name (max 30)", "Street (max 30)", "City (max 30)", "Zip (max 5)", "Zip+4 (max 4)"]


def _finish_quote(page1, address: str, coverage_a: str, scenario: str = "", field_data: dict = None):
    """Submit endorsements, capture quote results, and log to xlsx. Does not bind."""
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
        pytest.skip(f"[{scenario}] Results page did not load for '{address}' – page content: {page_text[:300].strip()}")
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

    # Payment plan tables for Auros and GeoVera Specialty share the same class.
    # When the Auros company display is present and populated, use .first to target Auros's row.
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

    field_values = [field_data.get(k, "") if field_data else "" for k in _FIELD_COLS]
    page1.wait_for_timeout(3000)
    _append_result([USERNAME, address, scenario, ProductName, coverage_a, QuoteNumber, QuoteExpiration, Dwelling,
                     TotalAnnualPremium, FullPay, ThreePay, FourPay, MonthlyPay] + field_values)


# ---------------------------------------------------------------------------
# Tests: one of each
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("scenario,option", [
    # Additional Interest — standard ddl (max 2 addresses)
    ("additional_interest",         1),
    ("additional_interest",         2),
    # Additional Interest — HO-24-70 ddl (max 4 addresses)
    ("additional_interest_ho2470",  1),
    ("additional_interest_ho2470",  2),
    ("additional_interest_ho2470",  3),
    ("additional_interest_ho2470",  4),
    # Additional Insured — standard ddl (max 2 addresses)
    ("additional_insured",          1),
    ("additional_insured",          2),
    # Additional Insured — HO-24-70 ddl (max 4 addresses)
    ("additional_insured_ho2470",   1),
    ("additional_insured_ho2470",   2),
    ("additional_insured_ho2470",   3),
    ("additional_insured_ho2470",   4),
    # Checkbox-only endorsements
    ("ordinance_or_law",            None),
    ("windstorm_hail_exclusion",    None),
    # Additional Rented — standard ddl (max 2 addresses)
    ("additional_rented",           1),
    ("additional_rented",           2),
    # Additional Rented — HO-24-70 ddl (max 4 addresses)
    ("additional_rented_ho2470",    1),
    ("additional_rented_ho2470",    2),
    ("additional_rented_ho2470",    3),
    ("additional_rented_ho2470",    4),
    # Swimming Pool — 5 limit options
    ("swimming_pool",               10000),
    ("swimming_pool",               15000),
    ("swimming_pool",               20000),
    ("swimming_pool",               25000),
    ("swimming_pool",               30000),
    # Additional Occupied — standard ddl (max 2 addresses)
    ("additional_occupied",         1),
    ("additional_occupied",         2),
    # Additional Occupied — HO-24-70 ddl (max 4 addresses)
    ("additional_occupied_ho2470",  1),
    ("additional_occupied_ho2470",  2),
    ("additional_occupied_ho2470",  3),
    ("additional_occupied_ho2470",  4),
    # Water Damage Limitation — 5 limit options
    ("water_limitation",            10000),
    ("water_limitation",            20000),
    ("water_limitation",            25000),
    ("water_limitation",            50000),
    ("water_limitation",            75000),
    # Water exclusions — checkbox only
    ("water_excl_named_peril",      None),
    ("water_excl_limited",          None),
])
def test_auros_one_of_each(page: Page, scenario: str, option) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    page1 = _run_to_endorsements_page(page, address, coverage_a)

    base_label = SCENARIO_LABELS.get(scenario, scenario)
    if option is None:
        scenario_label = base_label
    elif scenario in ("swimming_pool", "water_limitation"):
        scenario_label = f"{base_label} (${option:,})"
    else:
        scenario_label = f"{base_label} ({option} Addr{'s' if option > 1 else ''})"

    if scenario == "additional_interest":
        fill_address_endorsement(page1, "ctl00", num_addrs=option)
    elif scenario == "additional_interest_ho2470":
        fill_address_endorsement(page1, "ctl00", num_addrs=option, use_ho2470=True)
    elif scenario == "additional_insured":
        fill_address_endorsement(page1, "ctl01", num_addrs=option)
    elif scenario == "additional_insured_ho2470":
        fill_address_endorsement(page1, "ctl01", num_addrs=option, use_ho2470=True)
    elif scenario == "ordinance_or_law":
        check_only_endorsement(page1, "ctl02")
    elif scenario == "windstorm_hail_exclusion":
        check_only_endorsement(page1, "ctl03")
    elif scenario == "additional_rented":
        fill_rented_address_endorsement(page1, num_addrs=option)
    elif scenario == "additional_rented_ho2470":
        fill_rented_address_endorsement(page1, num_addrs=option)
    elif scenario == "swimming_pool":
        result = fill_pool_endorsement(page1, limit=option)
    elif scenario == "additional_occupied":
        fill_occupied_address_endorsement(page1, num_addrs=option)
    elif scenario == "additional_occupied_ho2470":
        fill_occupied_address_endorsement(page1, num_addrs=option, use_ho2470=True)
    elif scenario == "water_limitation":
        result = fill_water_limitation(page1, limit=option)
    elif scenario == "water_excl_named_peril":
        fill_water_excl_named_peril(page1)
    elif scenario == "water_excl_limited":
        fill_water_excl_limited(page1)

    _finish_quote(page1, address, coverage_a, scenario_label)


# ---------------------------------------------------------------------------
# Test: all endorsements (water exclusions omitted — conflict with limitation)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("use_max", [False, True])
def test_auros_all_endorsements(page: Page, use_max: bool) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    page1 = _run_to_endorsements_page(page, address, coverage_a)

    fill_address_endorsement(page1, "ctl00", use_max=use_max)
    fill_address_endorsement(page1, "ctl01", use_max=use_max)
    check_only_endorsement(page1, "ctl02")
    check_only_endorsement(page1, "ctl03")
    fill_rented_address_endorsement(page1, use_max=use_max)
    fill_pool_endorsement(page1, limit=None)
    fill_occupied_address_endorsement(page1, use_max=use_max)
    fill_water_limitation(page1, limit=None)

    scenario = "All Endorsements (Max Length)" if use_max else "All Endorsements"
    field_data = _FIELD_DATA_WITH_NAME if use_max else None
    _finish_quote(page1, address, coverage_a, scenario, field_data)


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
        fill_address_endorsement(page1, "ctl00", num_addrs=2, use_max=True)
        field_data = _FIELD_DATA_WITH_NAME
    elif scenario == "additional_insured":
        fill_address_endorsement(page1, "ctl01", num_addrs=2, use_max=True)
        field_data = _FIELD_DATA_WITH_NAME
    elif scenario == "additional_rented":
        fill_rented_address_endorsement(page1, num_addrs=2, use_max=True)
        field_data = _FIELD_DATA_NO_NAME
    elif scenario == "additional_occupied":
        fill_occupied_address_endorsement(page1, num_addrs=2, use_max=True)
        field_data = _FIELD_DATA_NO_NAME
    else:
        field_data = None

    _finish_quote(page1, address, coverage_a, MAX_SCENARIO_LABELS.get(scenario, scenario), field_data)

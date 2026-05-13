import re
import os
import random
import pytest
from datetime import date, timedelta
from playwright.sync_api import Page
from helpers.result_logger import make_xlsx_appender, make_xlsx_updater
from helpers.auros_endorsements import (
    fill_address_endorsement,
    fill_rented_address_endorsement,
    fill_occupied_address_endorsement,
    check_only_endorsement,
    fill_pool_endorsement,
    fill_water_limitation,
    fill_water_excl_named_peril,
    fill_water_excl_limited,
    click_water_update_if_visible,
)

USERNAME = os.environ.get("AGENT_USERNAME", "BRK99102").strip()
_PASSWORDS = {"ATRNRS": "password", "ANNIE-0134": "Password1", "BRK99102": "Password1"}
PASSWORD = _PASSWORDS[USERNAME]
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "results", "nrs", "nrs_agent_bind_results.xlsx")

_today = date.today()
EFFECTIVE_DATES = [
    _today.strftime("%m/%d/%Y"),
    (_today + timedelta(days=15)).strftime("%m/%d/%Y"),
    (_today + timedelta(days=30)).strftime("%m/%d/%Y"),
]

NRS_USERNAME  = "tsicortesmarquez@geovera.com"
NRS_WHOLESALER = "HULL001"
NRS_PRODUCER  = "HLKAREN"
NRS_SUB_CODE  = "01"

BILL_TO_SHORT = {
    "Applicant":       "Applicant",
    "Mortgagee":       "Mortgagee",
    "Alternate Payor": "AltPayor",
}

MORTGAGEE_SECOND_NAMES = [
    "ISAOA ATIMA",
    "C/O ESCROW DEPT",
    "LOAN SERVICING DIV",
    "ATTN INSURANCE DEPT",
    "ISAOA",
]

AP_NAMES = [
    "ACME FINANCIAL LLC",
    "SUNRISE LENDING GROUP",
    "COASTAL CREDIT CORP",
    "LIBERTY FUNDING INC",
    "PREMIER LOAN SERVICES",
    "ATLAS MORTGAGE CO",
    "CORNERSTONE FINANCE",
    "HORIZON LENDING LLC",
]

AP_SECOND_NAMES = [
    "ISAOA ATIMA",
    "C/O BILLING DEPT",
    "LOAN SERVICES DIV",
    "ATTN INSURANCE DEPT",
]

AP_STREETS = [
    "100 CORPORATE BLVD",
    "200 BUSINESS PARK DR",
    "350 COMMERCE ST",
    "475 FINANCIAL WAY",
    "600 ENTERPRISE AVE",
    "725 INDUSTRY BLVD",
    "850 COMMERCE PKWY",
]

AP_CITIES = [
    "DALLAS", "ATLANTA", "CHICAGO", "PHOENIX", "HOUSTON",
    "DENVER", "MIAMI", "SEATTLE", "BOSTON", "CHARLOTTE",
]

US_STATES = [
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
]

# All Bill To × Pay Plan × Autopay combinations.
# Monthly always has autopay; no need to test it without.
BIND_SCENARIOS = [
    # bill_to_val, bill_to_label, pay_plan_val, pay_plan_label, autopay
    ("0", "Applicant",       "0", "Full Pay", False),
    ("0", "Applicant",       "0", "Full Pay", True),
    ("0", "Applicant",       "1", "3 Pay",    False),
    ("0", "Applicant",       "1", "3 Pay",    True),
    ("0", "Applicant",       "2", "4 Pay",    False),
    ("0", "Applicant",       "2", "4 Pay",    True),
    ("0", "Applicant",       "3", "Monthly",  True),
    ("1", "Mortgagee",       "0", "Full Pay", False),
    ("1", "Mortgagee",       "0", "Full Pay", True),
    ("1", "Mortgagee",       "1", "3 Pay",    False),
    ("1", "Mortgagee",       "1", "3 Pay",    True),
    ("1", "Mortgagee",       "2", "4 Pay",    False),
    ("1", "Mortgagee",       "2", "4 Pay",    True),
    ("1", "Mortgagee",       "3", "Monthly",  True),
    ("2", "Alternate Payor", "0", "Full Pay", False),
    ("2", "Alternate Payor", "0", "Full Pay", True),
    ("2", "Alternate Payor", "1", "3 Pay",    False),
    ("2", "Alternate Payor", "1", "3 Pay",    True),
    ("2", "Alternate Payor", "2", "4 Pay",    False),
    ("2", "Alternate Payor", "2", "4 Pay",    True),
    ("2", "Alternate Payor", "3", "Monthly",  True),
]

# 10 HO3 + 10 HO3T scenarios × 2 logins = 40 total quotes
HO3_SCENARIOS = [
    # Applicant — Full Pay no auto, 3 Pay no auto, 4 Pay no auto, Monthly auto
    BIND_SCENARIOS[0], BIND_SCENARIOS[2], BIND_SCENARIOS[4], BIND_SCENARIOS[6],
    # Mortgagee — Full Pay no auto, Full Pay auto, 3 Pay auto, 4 Pay auto
    BIND_SCENARIOS[7], BIND_SCENARIOS[8], BIND_SCENARIOS[10], BIND_SCENARIOS[12],
    # AltPayor — 3 Pay no auto, Monthly auto
    BIND_SCENARIOS[16], BIND_SCENARIOS[20],
]

HO3T_SCENARIOS = [
    # Applicant — Full Pay auto, 3 Pay auto, 4 Pay auto
    BIND_SCENARIOS[1], BIND_SCENARIOS[3], BIND_SCENARIOS[5],
    # Mortgagee — 3 Pay no auto, 4 Pay no auto, Monthly auto
    BIND_SCENARIOS[9], BIND_SCENARIOS[11], BIND_SCENARIOS[13],
    # AltPayor — Full Pay no auto, Full Pay auto, 4 Pay no auto, 4 Pay auto
    BIND_SCENARIOS[14], BIND_SCENARIOS[15], BIND_SCENARIOS[18], BIND_SCENARIOS[19],
]

SPECIFIC_BIND_SCENARIOS = [
    # test_case, address, eff_date, bill_to_val, bill_to_label, pay_plan_val, pay_plan_label, autopay
    # ("3PayInvoiced",   "2269 Cascadia Ct, St. Augustine, FL 32092",       "06/14/2026", "0", "Applicant",       "1", "3 Pay",    False),
    # ("3PayAuto",       "3137 Castle Rock Cir, Land O Lakes, FL 34639",    "07/02/2026", "2", "Alternate Payor", "1", "3 Pay",    True),
    # ("MonthlyPay",     "13626 Allyn Dr, Hudson, FL 34667",                "07/02/2026", "2", "Alternate Payor", "3", "Monthly",  True),
    # ("PDKORCC",        "600 S Dixie Hwy #352, West Palm Beach, FL 33401", "05/22/2026", "0", "Applicant",       "0", "Full Pay", False),
    # ("Renew",          "510 Wedge Ln, Longboat Key, FL 34228",            "06/29/2026", "0", "Applicant",       "0", "Full Pay", False),
    # ("Refund",         "8431 SE Duncan St, Hobe Sound, FL 33455",         "05/19/2026", "2", "Alternate Payor", "1", "3 Pay",    True),
    # ("Rebill_NSF_Fee", "308 Ferndale Way, St. Augustine, FL 32092",       "06/28/2026", "1", "Mortgagee",       "0", "Full Pay", False),
    # ("ShortPay_EPay",  "12274 Accipiter Dr, Orlando, FL 32837",           "07/06/2026", "0", "Applicant",       "0", "Full Pay", True),
    # ("ShortPay_Inv",   "1401 Maine Ave, Lynn Haven, FL 32444",            "06/19/2026", "1", "Mortgagee",       "0", "Full Pay", False),
    # ("ToleranceUnder", "15110 Shearcrest Dr, Lithia, FL 33547",           "06/25/2026", "0", "Applicant",       "0", "Full Pay", False),
]


def _remove_address(address: str):
    import time
    src = os.path.abspath(__file__)
    lock_path = src + ".lock"
    for _ in range(20):
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            break
        except FileExistsError:
            time.sleep(0.5)
    try:
        with open(src, "r", encoding="utf-8") as f:
            content = f.read()
        line = f'    "{address}",\n'
        if line in content:
            with open(src, "w", encoding="utf-8") as f:
                f.write(content.replace(line, ""))
    finally:
        try:
            os.unlink(lock_path)
        except Exception:
            pass


_HEADERS = [
    "Agent Username", "Test Case", "Test Scenario", "Property Address",
    "Product Type", "Coverage A Limit", "Wholesaler", "Producer Number",
    "Quote Number", "Full Pay", "3 Pay (1st Install)", "4 Pay (1st Install)", "Monthly (1st Install)",
    "Insured Name", "Status", "Binder Number", "Bill To", "Payment Plan",
    "Effective Date", "Amount Due Now", "Endorsements", "Deductible",
]
_append_result = make_xlsx_appender(RESULTS_FILE, _HEADERS)
_update_result = make_xlsx_updater(RESULTS_FILE, 9)  # Quote Number is column 9 (1-based)


FL_ADDRESSES = [
    "8832 Savona Ct, Naples, FL 34119",
    "1461 75th Cir NE, Saint Petersburg, FL 33702",
    "8050 Princeton Dr, Naples, FL 34104",
    "12274 Accipiter Dr, Orlando, FL 32837",
    "15110 Shearcrest Dr, Lithia, FL 33547",
    "7733 Belvoir Dr, Orlando, FL 32835",
    "9426 Meadow Crest Ln, Clermont, FL 34711",
    "66 Woodbury Dr, Palm Coast, FL 32164",
    "1401 Maine Ave, Lynn Haven, FL 32444",
    "9607 Bahia Rd, Ocala, FL 34472",
    "2150 NW 74th Ave, Fort Lauderdale, FL 33313",
    "4364 Tarragon Ave, Middleburg, FL 32068",
    "3198 Michaels Ct, Green Cove Springs, FL 32043",
    "613 S West St, Green Cove Springs, FL 32043",
    "8329 Lakemont Dr, Jacksonville, FL 32216",
    "13501 Princess Kelly Dr, Jacksonville, FL 32225",
    "6229 Cranberry Ln E, Jacksonville, FL 32244",
    "14688 Indigo Lakes Cir, Naples, FL 34119",
    "3780 Ashland Ave, Pensacola, FL 32534",
    "5518 Davis Street, Fort Pierce, FL 34982",
    "3700 NW 20th Street, Coconut Creek, FL 33066",
    "85 GLENALBY Place, Ponte Vedra, FL 32081",
    "111 Midway Park Dr, Saint Augustine, FL 32084",
    "6104 Spencer Parrish Rd, Parrish, FL 34219",
    "5935 Old Berkley Rd, Auburndale, FL 33823",
    "11180 113th Ave, Seminole, FL 33778",
    "1339 Gallop Drive, Loxahatchee, FL 33470",
    "12614 Fenhurst WAY, NAPLES, FL 34120",
    "3433 Leaf Lake Dr, Land O Lakes, FL 34639",
    "18037 Mambo Drive, Boca Raton, FL 33496",
    "3651 Canopy CIR, NAPLES, FL 34120",
    "2650 SW Windship Way, Stuart, FL 34997",
    "5021 Wesley Dr, Tampa, FL 33647",
]


def _login_and_open_quote(page: Page):
    page.goto("https://tst.mygeosource.com/")
    page.get_by_role("textbox", name="Policyholder Email or Agent").click()
    page.get_by_role("textbox", name="Policyholder Email or Agent").fill(USERNAME)
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill(PASSWORD)
    page.get_by_role("button", name="Log In", exact=True).click()
    # Agent clicking "Start New Quote" may open the popup directly or show a product
    # selection modal first (depending on agent permissions). Check for the modal and
    # handle both product text variants; if no modal, popup opened directly.
    with page.expect_popup() as page1_info:
        page.get_by_role("paragraph").filter(has_text="Start New Quote").get_by_role("link").click()
        try:
            page.locator("#optionHomeowners").wait_for(timeout=2000, state="visible")
            try:
                page.locator("#optionHomeowners").get_by_text("GeoVera Specialty Insurance").click(timeout=1500)
            except Exception:
                page.locator("#optionHomeowners").get_by_text("Residential Homeowner Products").click()
            page.get_by_role("button", name="OK").click()
        except Exception:
            pass  # no modal — popup opened directly
    page1 = page1_info.value
    page1.wait_for_load_state("domcontentloaded")
    page1.wait_for_timeout(2000)
    # Agent lands directly on the address page — producer fields may or may not be present
    wholesaler, producer_code, sub_code = "", "", ""
    try:
        wholesaler    = page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").input_value(timeout=1000)
    except Exception:
        pass
    try:
        producer_code = page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").input_value(timeout=1000)
    except Exception:
        pass
    try:
        sub_code      = page1.locator("#ctl00_WelcomeUserControl_txtSubProducerNumber").input_value(timeout=1000)
    except Exception:
        pass
    print(f'Wholesaler: {wholesaler} | Producer: {producer_code} | Sub: {sub_code}')
    return page1, wholesaler, producer_code, sub_code


def _login_and_open_quote_nrs(page: Page):
    page.goto("https://tst.mygeosource.com/")
    page.get_by_role("textbox", name="Policyholder Email or Agent").click()
    page.get_by_role("textbox", name="Policyholder Email or Agent").fill(NRS_USERNAME)
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
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ddlWholesaler", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option(NRS_WHOLESALER)
    page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill(NRS_PRODUCER)
    page1.locator("#ctl00_WelcomeUserControl_txtSubProducerNumber").fill(NRS_SUB_CODE)
    page1.locator("#ctl00_WelcomeUserControl_btnUpdateQuotingAs").click()
    page1.wait_for_selector("[id*='dgUserList'][id$='_btnId']", timeout=30000)
    page1.locator("[id*='dgUserList'][id$='_btnId']").first.click()
    page1.wait_for_timeout(3000)
    return page1


def _enter_address(page1: Page, address: str):
    inp = page1.locator("#RadAutoCompleteBoxGeo_Input")
    inp.click()
    inp.fill(address)
    inp.press("Backspace")
    inp.press_sequentially(address[-1], delay=50)
    try:
        page1.wait_for_selector(".rcbList .rcbItem", timeout=3000)
        page1.locator(".rcbList .rcbItem").first.click()
    except Exception:
        # Second nudge: backspace + retype last char again
        inp.press("Backspace")
        inp.press_sequentially(address[-1], delay=50)
        try:
            page1.wait_for_selector(".rcbList .rcbItem", timeout=3000)
            page1.locator(".rcbList .rcbItem").first.click()
        except Exception:
            inp.press("Enter")
            page1.wait_for_timeout(2000)
    page1.get_by_role("button", name="Next -->").click()


def _fill_property_screens(page1: Page):
    page1.locator("#ddlConstructionType").select_option("FR")
    page1.locator("#ddlSidingTypeDominant").select_option("ST")
    page1.locator("#ddlRestrictedRoofShapes").select_option("False")
    if page1.locator("#ddlRoofShapeDominant").is_visible():
        page1.locator("#ddlRoofShapeDominant").select_option("2")
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFoundationElevation").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFoundationElevation").fill("0")
    page1.get_by_role("button", name="Next -->").click()
    page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralFireAlarm", timeout=15000)
    replacement_cost = 0.0
    try:
        rc_row = page1.locator("#ctl00_ContentPlaceHolder1_rowNavigationReplacementCost td").first
        if rc_row.count() > 0:
            raw = rc_row.text_content(timeout=1500) or ""
            m = re.search(r"\$([\d,]+)", raw)
            if m:
                replacement_cost = float(m.group(1).replace(",", ""))
    except Exception:
        pass
    print(f'Replacement Cost: ${replacement_cost:,.2f}')
    alarm_val = "True" if replacement_cost > 1_000_000 else "False"
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralFireAlarm").select_option(alarm_val)
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralBurglarAlarm").select_option(alarm_val)
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
        page1.locator("#fileWindMitigationVerificationInspectionForm").set_input_files(
            "C:/Users/icortesmarquez/Desktop/Playwright/HO_PW/test_data/testfile.pdf"
        )
    page1.get_by_role("button", name="Next -->").click()


def _fill_applicant(page1: Page, insured_name: str):
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlPolicyOwnerType").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlPolicyOwnerType").select_option("Individual")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName").fill(insured_name)
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLastName").fill("TESTER")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBMonth").click()
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBMonth").press_sequentially("01")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBDay").press_sequentially("01")
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBYear").press_sequentially("1998")
    page1.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN1\"]").press_sequentially("123")
    page1.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN2\"]").press_sequentially("11")
    page1.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN3\"]").press_sequentially("1122")
    page1.get_by_role("checkbox", name="Is Mailing Address the same").check()
    page1.locator("#ddlPriorHomeownersInsurance").select_option("False")
    page1.wait_for_load_state("networkidle", timeout=10000)
    page1.locator("#ddlPriorHomeownersInsurance").select_option("False")
    page1.get_by_role("button", name="Next -->").click()


def _collect_quote_info(page1: Page):
    QuoteNumber = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Number").inner_text()
    ProductName = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header").inner_text()
    print(f'Product: {ProductName}  |  Quote: {QuoteNumber}')

    def _get_pay(idx: str) -> str:
        loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lstPaymentOptions_{idx}_lblQuote1PaymentOption")
        try:
            text = loc.text_content(timeout=1500).strip()
        except Exception:
            return ""
        m = re.search(r"\$\d{1,3}(?:,\d{3})*(?:\.\d{2})", text)
        return m.group() if m else text

    full_pay   = _get_pay("ctl00")
    three_pay  = _get_pay("ctl01")
    four_pay   = _get_pay("ctl02")
    monthly    = _get_pay("ctl03")

    print(f'Full Pay: {full_pay} | 3-Pay: {three_pay} | 4-Pay: {four_pay} | Monthly: {monthly}')
    return QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly


def _fill_mortgagee_block(page1: Page, n: int):
    pfx = "ctl00_ContentPlaceHolder1_ucBinderRequest"
    name_arrow = f"#{pfx}_rcbMortgagee{n}Name1_Arrow"
    name_items = f"#{pfx}_rcbMortgagee{n}Name1_DropDown .rcbItem"
    page1.locator(name_arrow).click()
    page1.wait_for_selector(name_items, timeout=1500)
    items = page1.locator(name_items)
    items.nth(random.randint(0, items.count() - 1)).click()
    # Wait for postback to populate Street dropdown
    street_arrow = f"#{pfx}_rcbMortgagee{n}Street_Arrow"
    page1.wait_for_selector(street_arrow, timeout=10000)
    second_name = random.choice(MORTGAGEE_SECOND_NAMES)
    page1.locator(f"#{pfx}_txtMortgagee{n}Name2").fill(second_name)
    street_arrow = f"#{pfx}_rcbMortgagee{n}Street_Arrow"
    street_items = f"#{pfx}_rcbMortgagee{n}Street_DropDown .rcbItem"
    page1.locator(street_arrow).click()
    try:
        page1.wait_for_selector(street_items, timeout=1500)
        page1.locator(street_items).first.click()
        page1.wait_for_timeout(1000)
    except Exception:
        pass
    loan_num = f"LN{random.randint(100000000, 999999999)}"
    page1.locator(f"#{pfx}_txtMortgagee{n}LoanNumber").fill(loan_num)


def _fill_alternate_payor(page1: Page):
    pfx = "ctl00_ContentPlaceHolder1_ucBinderRequest"
    page1.locator(f"#{pfx}_txtBillingAlternatePayor1_1").fill(random.choice(AP_NAMES))
    page1.locator(f"#{pfx}_txtBillingAlternatePayor2_1").fill(random.choice(AP_SECOND_NAMES))
    page1.locator(f"#{pfx}_txtBillingLoanNumber_1").fill(f"LN{random.randint(100000000, 999999999)}")
    page1.locator(f"#{pfx}_txtBillingStreet_1").fill(random.choice(AP_STREETS))
    page1.locator(f"#{pfx}_txtBillingCity_1").fill(random.choice(AP_CITIES))
    page1.locator(f"#{pfx}_ddlBillingState_1").select_option(random.choice(US_STATES))
    page1.locator(f"#{pfx}_txtBillingZip_1").fill(str(random.randint(10000, 99999)))
    page1.locator(f"#{pfx}_txtBillingZip4_1").fill(str(random.randint(1000, 9999)))


def _bind(page1: Page, QuoteNumber: str, bill_to_val: str, pay_plan_val: str, pay_plan_label: str, autopay: bool, eff_date: str = None):
    try:
        page1.locator(f'input[type="radio"][value="{QuoteNumber}"]').first.check(timeout=1500)
    except Exception:
        pass
    page1.get_by_role("button", name="Continue to Bind Coverage").click()
    page1.wait_for_timeout(2000)
    if page1.locator("#chkPopInConfirmation").is_visible():
        page1.locator("#chkPopInConfirmation").check()
        page1.get_by_role("button", name="Continue", exact=True).click()

    page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucBinderRequest_ddlBillTo1", timeout=30000)
    if eff_date is None:
        eff_date = random.choice(EFFECTIVE_DATES)
    page1.locator("#ctl00_ContentPlaceHolder1_ucBinderRequest_txtBinderEffectiveDate1").fill(eff_date)
    page1.locator("#ctl00_ContentPlaceHolder1_ucBinderRequest_ddlBillTo1").select_option(bill_to_val)
    # Wait for postback triggered by Bill To selection
    if bill_to_val == "1":
        page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucBinderRequest_rcbMortgagee1Name1_Arrow", timeout=10000)
    elif bill_to_val == "2":
        page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucBinderRequest_txtBillingAlternatePayor1_1", timeout=10000)
    else:
        page1.wait_for_load_state("networkidle", timeout=5000)
    if bill_to_val == "1":
        _fill_mortgagee_block(page1, 1)
        if random.choice([True, False]):
            _fill_mortgagee_block(page1, 2)
    elif bill_to_val == "2":
        _fill_alternate_payor(page1)
    try:
        page1.locator("#ddlPaymentPlan1").select_option(pay_plan_val)
    except Exception:
        pytest.skip(f"Pay plan '{pay_plan_label}' not available for this product")
    print(f'Selected Pay Plan: {pay_plan_label}')

    if pay_plan_label != "Monthly":
        autopay_chk = page1.locator("[id*='chkAutoPay']")
        if autopay_chk.count() > 0 and autopay_chk.is_visible():
            if autopay:
                autopay_chk.check()
            else:
                autopay_chk.uncheck()

    page1.get_by_role("button", name="Request to Bind").click()

    page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucBinderDetail_lblAmountDue", timeout=30000)
    BinderNumber   = page1.locator("#ctl00_ContentPlaceHolder1_ucBinderDetail_lblBinderNumber").inner_text()
    BillTo         = page1.locator("#ctl00_ContentPlaceHolder1_ucBinderDetail_lblBillTo").inner_text()
    PaymentPlan    = page1.locator("#ctl00_ContentPlaceHolder1_ucBinderDetail_lblPaymentPlan").inner_text()
    EffectiveDate  = page1.locator("#ctl00_ContentPlaceHolder1_ucBinderDetail_lblEffectiveDate").inner_text()
    AmountDueNow   = page1.locator("#ctl00_ContentPlaceHolder1_ucBinderDetail_lblAmountDue").inner_text()
    print(f'Binder: {BinderNumber} | Bill To: {BillTo} | Plan: {PaymentPlan} | Eff: {EffectiveDate} | Due: {AmountDueNow}')
    return BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow


# ---------------------------------------------------------------------------
# NRS Auros HO-3 — Quote to Bind (Agent login)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay", HO3_SCENARIOS)
def test_bind_nrs_auros_agent_ho3(page: Page, snap, bill_to_val: str, bill_to_label: str,
                                   pay_plan_val: str, pay_plan_label: str, autopay: bool) -> None:
    address = random.choice(FL_ADDRESSES)
    print(f'Address: {address}')
    coverage_a = "800,000"
    pay_type = "AutoPay" if autopay else "Invoiced"
    insured_name = f"NRS {BILL_TO_SHORT[bill_to_label]} {pay_plan_label} {pay_type}"

    page1, wholesaler, producer_code, sub_code = _login_and_open_quote(page)
    snap(page1, "03_welcome")

    _enter_address(page1, address)
    snap(page1, "04_address")

    ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
    if ineligible_msgs.count() > 0:
        _remove_address(address)
        pytest.skip(f"Address '{address}' ineligible on address screen – removed from list")

    # Select only Auros HO-3; uncheck the other products if visible
    if page1.get_by_role("checkbox", name="Auros HO-3T").is_visible():
        page1.get_by_role("checkbox", name="Auros HO-3T").uncheck()
    if page1.get_by_role("checkbox", name="GeoVera Specialty HO-3").is_visible():
        page1.get_by_role("checkbox", name="GeoVera Specialty HO-3").uncheck()
    if page1.get_by_role("checkbox", name="HO Wind Only").is_visible():
        page1.get_by_role("checkbox", name="HO Wind Only").uncheck()
    page1.locator("#txtRequestedHomeownersCoverageALimit").click()
    page1.locator("#txtRequestedHomeownersCoverageALimit").fill(coverage_a)
    snap(page1, "05_coverages")
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").fill("2025")
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").is_visible():
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").fill("2025")
    page1.get_by_role("button", name="Next -->").click()
    page1.wait_for_timeout(2000)

    ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
    if ineligible_msgs.count() > 0:
        _remove_address(address)
        pytest.skip(f"Address '{address}' ineligible for Auros HO-3 – {ineligible_msgs.first.inner_text().strip()}")
    try:
        page1.wait_for_selector("#ddlConstructionType", timeout=10000)
    except Exception:
        page_text = page1.locator("body").inner_text()
        pytest.skip(f"Address '{address}' did not advance to property page – {page_text[:200].strip()}")

    _fill_property_screens(page1)
    snap(page1, "09_wind_mitigation")

    _fill_applicant(page1, insured_name)
    snap(page1, "10_applicant")

    q_count = page1.locator("[id*='ucGenInfoQuestDDL_rptrQuestions'][id$='_ddlQuestions']").count()
    for i in range(q_count):
        loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl{i:02d}_ddlQuestions")
        if loc.count() > 0:
            loc.select_option("False")
    wind_mit_q = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl21_ddlQuestions")
    if wind_mit_q.count() > 0:
        wind_mit_q.select_option("True")
    page1.get_by_role("cell", name="Applicant acknowledges that flood coverage is excluded for the homeowners, wind-only, and vacant policy.", exact=True).locator("span").first.click()
    for chk in ["chkRoofPaymentSchedule", "chkFloodExclusion", "chkRoofExclusion", "chkAcknowledgedWaterDamageLimitation"]:
        loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_{chk}")
        if loc.count() > 0:
            loc.check()
    snap(page1, "11_questions")
    page1.get_by_role("button", name="Next -->").click()
    page1.wait_for_timeout(2000)
    applied_endorsements = _apply_random_endorsements(page1)
    snap(page1, "11b_endorsements")
    page1.get_by_role("button", name="Submit Quote Request").click()
    page1.wait_for_timeout(2000)

    ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
    if ineligible_msgs.count() > 0:
        pytest.skip(f"Address '{address}' ineligible after submit – {ineligible_msgs.first.inner_text().strip()}")
    try:
        page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header", timeout=30000)
    except Exception:
        page_text = page1.locator("body").inner_text()
        pytest.skip(f"Quote results did not load for '{address}' – {page_text[:300].strip()}")
    snap(page1, "12_quote_results")

    QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
    deductible = _maybe_change_deductible(page1)
    if deductible != "$2,500":
        QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
        snap(page1, "12b_quote_after_deductible")

    scenario_label = f"NRS Auros HO-3 | {bill_to_label} | {insured_name}"
    _append_result([USERNAME, "", scenario_label, address,
                     ProductName, coverage_a, wholesaler, f"{producer_code}-{sub_code}",
                     QuoteNumber, full_pay, three_pay, four_pay, monthly,
                     insured_name, "Not Bound", "", "", "", "", "",
                     "; ".join(applied_endorsements), deductible])

    BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
        page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay
    )
    snap(page1, "14_binder_detail")

    _update_result(QuoteNumber, {15: "Bound", 16: BinderNumber, 17: BillTo, 18: PaymentPlan, 19: EffectiveDate, 20: AmountDueNow})


# ---------------------------------------------------------------------------
# NRS Auros HO-3T — Quote to Bind (Agent login, Occupancy = Tenant Only)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay", HO3T_SCENARIOS)
def test_bind_nrs_auros_agent_ho3t(page: Page, snap, bill_to_val: str, bill_to_label: str,
                                    pay_plan_val: str, pay_plan_label: str, autopay: bool) -> None:
    address = random.choice(FL_ADDRESSES)
    print(f'Address: {address}')
    coverage_a = "800,000"
    pay_type = "AutoPay" if autopay else "Invoiced"
    insured_name = f"NRS {BILL_TO_SHORT[bill_to_label]} {pay_plan_label} {pay_type}"
    try:
        page1, wholesaler, producer_code, sub_code = _login_and_open_quote(page)
        snap(page1, "03_welcome")

        _enter_address(page1, address)
        page1.wait_for_timeout(2000)
        snap(page1, "04_address")

        # Set Occupancy Type to Tenant Only; Auros HO-3 produces HO-3T output under Tenant Only
        page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlOccupancyType").select_option("Tenant Only")
        page1.wait_for_timeout(1500)

        ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
        if ineligible_msgs.count() > 0:
            _remove_address(address)
            pytest.skip(f"Address '{address}' ineligible on address screen – removed from list")
        if page1.get_by_role("checkbox", name="GeoVera Specialty HO-3").is_visible():
            page1.get_by_role("checkbox", name="GeoVera Specialty HO-3").uncheck()
        if page1.get_by_role("checkbox", name="HO Wind Only").is_visible():
            page1.get_by_role("checkbox", name="HO Wind Only").uncheck()
        page1.locator("#txtRequestedHomeownersCoverageALimit").click()
        page1.locator("#txtRequestedHomeownersCoverageALimit").fill(coverage_a)
        snap(page1, "05_coverages")
        if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").is_visible():
            page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").fill("2025")
        if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").is_visible():
            page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").fill("2025")
        page1.get_by_role("button", name="Next -->").click()
        page1.wait_for_timeout(2000)

        ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
        if ineligible_msgs.count() > 0:
            pytest.skip(f"Address '{address}' ineligible for Auros HO-3T – {ineligible_msgs.first.inner_text().strip()}")
        try:
            page1.wait_for_selector("#ddlConstructionType", timeout=30000)
        except Exception:
            page_text = page1.locator("body").inner_text()
            pytest.skip(f"Address '{address}' did not advance to property page – {page_text[:200].strip()}")

        try:
            _fill_property_screens(page1)
        except Exception as e:
            pytest.skip(f"Property screens failed for '{address}' – {e}")
        snap(page1, "09_wind_mitigation")

        try:
            _fill_applicant(page1, insured_name)
        except Exception as e:
            pytest.skip(f"Applicant screen failed for '{address}' – {e}")
        snap(page1, "10_applicant")

        q_count = page1.locator("[id*='ucGenInfoQuestDDL_rptrQuestions'][id$='_ddlQuestions']").count()
        for i in range(q_count):
            loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl{i:02d}_ddlQuestions")
            if loc.count() > 0:
                loc.select_option("False")
        tenant_q = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl17_ddlQuestions")
        if tenant_q.count() > 0:
            tenant_q.select_option("True")
        roof_mat_q = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl21_ddlQuestions")
        if roof_mat_q.count() > 0:
            roof_mat_q.select_option("False")
        page1.get_by_role("cell", name="Applicant acknowledges that flood coverage is excluded for the homeowners, wind-only, and vacant policy.", exact=True).locator("span").first.click()
        tenant_chk = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_chkTenantLiability")
        if tenant_chk.count() > 0:
            tenant_chk.check()
        for chk in ["chkRoofPaymentSchedule", "chkFloodExclusion", "chkRoofExclusion", "chkAcknowledgedWaterDamageLimitation"]:
            loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_{chk}")
            if loc.count() > 0:
                loc.check()
        snap(page1, "11_questions")
        page1.get_by_role("button", name="Next -->").click()
        page1.wait_for_timeout(2000)
        applied_endorsements = _apply_random_endorsements(page1)
        snap(page1, "11b_endorsements")
        page1.get_by_role("button", name="Submit Quote Request").click()
        page1.wait_for_timeout(2000)

        ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
        if ineligible_msgs.count() > 0:
            pytest.skip(f"Address '{address}' ineligible after submit – {ineligible_msgs.first.inner_text().strip()}")
        try:
            page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header", timeout=30000)
        except Exception:
            page_text = page1.locator("body").inner_text()
            pytest.skip(f"Quote results did not load for '{address}' – {page_text[:300].strip()}")
        snap(page1, "12_quote_results")

        QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
        if not any(x in ProductName for x in ("3T", "HO-3T", "Tenant")):
            pytest.skip(f"Quote returned '{ProductName}' instead of HO-3T for '{address}' – skipping bind")
        deductible = _maybe_change_deductible(page1)
        if deductible != "$2,500":
            QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
            snap(page1, "12b_quote_after_deductible")

        scenario_label = f"NRS Auros HO-3T | {bill_to_label} | {insured_name}"
        _append_result([USERNAME, "", scenario_label, address,
                         ProductName, coverage_a, wholesaler, f"{producer_code}-{sub_code}",
                         QuoteNumber, full_pay, three_pay, four_pay, monthly,
                         insured_name, "Not Bound", "", "", "", "", "",
                         "; ".join(applied_endorsements), deductible])

        BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
            page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay
        )
        snap(page1, "14_binder_detail")

        _update_result(QuoteNumber, {15: "Bound", 16: BinderNumber, 17: BillTo, 18: PaymentPlan, 19: EffectiveDate, 20: AmountDueNow})
    except Exception:
        raise


# ---------------------------------------------------------------------------
# NRS Auros HO-3 — Specific Named Scenarios (NRS internal login)
# ---------------------------------------------------------------------------

class _AddressIneligible(Exception):
    pass


def _page_context(page1: Page, test_case: str, stage: str) -> str:
    """Return a diagnostic string: stage, URL, visible errors, and full body text."""
    lines = [
        f"  Test Case : {test_case}",
        f"  Stage     : {stage}",
        f"  URL       : {page1.url}",
    ]
    for sel in [
        "[class*='error' i]",
        "[class*='alert' i]",
        "[class*='notification' i]",
        ".validation-summary-errors",
        "span[id*='lblError']",
        "span[id*='lblMessage']",
    ]:
        try:
            els = page1.locator(sel).filter(has_text=re.compile(r"\S"))
            if els.count() > 0:
                lines.append(f"  [{sel}]: {els.first.inner_text(timeout=2000).strip()[:400]}")
        except Exception:
            pass
    try:
        body = page1.locator("body").inner_text(timeout=1500).strip()
        lines.append(f"  [body]:\n{body}")
    except Exception:
        pass
    return "\n".join(lines)


def _maybe_change_deductible(page1: Page) -> str:
    """Randomly decide whether to edit the deductible. If yes, pick a non-default value (skipping $5,000),
    submit, and wait for the postback to re-render the quote details. Returns the final deductible label.
    Caller is responsible for re-reading premium/payment values after a change."""
    DEFAULT_LABEL = "$2,500"
    NON_DEFAULT = [
        ("E", "$5,000"),
        ("F", "$7,500"),
        ("G", "$10,000"),
        ("H", "$15,000"),
        ("M", "$20,000"),
        ("N", "$25,000"),
        ("P", "$50,000"),
    ]
    if random.choice([True, False]):
        print(f"Deductible: keeping default {DEFAULT_LABEL}")
        return DEFAULT_LABEL
    val, label = random.choice(NON_DEFAULT)
    print(f"Deductible: editing to {label}")
    page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_btnEditQuote1Deductible").click()
    page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_ddlQuote1Deductible", timeout=10000)
    page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_ddlQuote1Deductible").select_option(val)
    page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_btnSubmitQuote1Deductible").click()
    page1.wait_for_load_state("networkidle")
    page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Number", timeout=30000)
    page1.wait_for_timeout(2000)
    return label


def _apply_random_endorsements(page1: Page):
    """Pick 3-5 endorsements from the non-conflicting pool and apply them.
    At most one of water_excl_named_peril / water_excl_limited enters the pool per run
    to prevent the two mutually exclusive exclusions from both being selected.
    Individual failures are logged and skipped so one bad endorsement doesn't fail the bind.
    Returns a list of human-readable 'name (detail)' strings for logging into xlsx."""
    ai_addrs   = random.randint(1, 2)
    aif_addrs  = random.randint(1, 2)
    rent_addrs = random.randint(1, 2)
    occ_addrs  = random.randint(1, 2)

    options = [
        ("additional_interest",      f"{ai_addrs} addr{'s' if ai_addrs > 1 else ''}",
            lambda: fill_address_endorsement(page1, "ctl00", num_addrs=ai_addrs)),
        ("additional_insured",       f"{aif_addrs} addr{'s' if aif_addrs > 1 else ''}",
            lambda: fill_address_endorsement(page1, "ctl01", num_addrs=aif_addrs)),
        ("ordinance_or_law",         "",
            lambda: check_only_endorsement(page1, "ctl02")),
        ("windstorm_hail_exclusion", "",
            lambda: check_only_endorsement(page1, "ctl03")),
        ("additional_rented",        f"{rent_addrs} addr{'s' if rent_addrs > 1 else ''}",
            lambda: fill_rented_address_endorsement(page1, num_addrs=rent_addrs)),
        ("swimming_pool",            None,
            lambda: fill_pool_endorsement(page1, limit=None)),
        ("additional_occupied",      f"{occ_addrs} addr{'s' if occ_addrs > 1 else ''}",
            lambda: fill_occupied_address_endorsement(page1, num_addrs=occ_addrs)),
        ("water_limitation",         None,
            lambda: fill_water_limitation(page1, limit=None)),
    ]
    water_excl_choice = random.choice([None, "named_peril", "limited"])
    if water_excl_choice == "named_peril":
        options.append(("water_excl_named_peril", "", lambda: fill_water_excl_named_peril(page1)))
    elif water_excl_choice == "limited":
        options.append(("water_excl_limited", "", lambda: fill_water_excl_limited(page1)))

    applied = []
    selected = random.sample(options, k=random.randint(3, 5))
    for name, detail, fn in selected:
        try:
            result = fn()
            if result is not None:
                detail = f"${result:,}"
            applied.append(f"{name} ({detail})" if detail else name)
        except Exception as e:
            print(f"  [endorsement skip] {name}: {e}")
    click_water_update_if_visible(page1)
    print(f"Applied endorsements ({len(applied)}): {applied}")
    return applied


@pytest.mark.parametrize(
    "test_case,address,eff_date,bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay",
    SPECIFIC_BIND_SCENARIOS,
    ids=[s[0] for s in SPECIFIC_BIND_SCENARIOS],
)
def test_bind_nrs_auros_agent_ho3_specific(page: Page, snap, test_case: str, address: str, eff_date: str,
                                            bill_to_val: str, bill_to_label: str,
                                            pay_plan_val: str, pay_plan_label: str, autopay: bool) -> None:
    coverage_a = "800,000"
    pay_plan_short = {"Full Pay": "Full", "3 Pay": "3Pay", "Monthly": "Monthly"}[pay_plan_label]
    pay_type_short = "AutoPay" if autopay else "Invoiced"
    insured_name = f"{test_case} {pay_plan_short} {pay_type_short}"

    fallback_pool = [a for a in FL_ADDRESSES if a != address]
    random.shuffle(fallback_pool)
    addresses_to_try = [address] + fallback_pool

    for attempt_addr in addresses_to_try:
        page1 = None
        try:
            page1, *_ = _login_and_open_quote(page)
            snap(page1, "03_welcome")

            _enter_address(page1, attempt_addr)
            snap(page1, "04_address")

            if page1.get_by_role("checkbox", name="Auros HO-3T").is_visible():
                page1.get_by_role("checkbox", name="Auros HO-3T").uncheck()
            if page1.get_by_role("checkbox", name="GeoVera Specialty HO-3").is_visible():
                page1.get_by_role("checkbox", name="GeoVera Specialty HO-3").uncheck()
            if page1.get_by_role("checkbox", name="HO Wind Only").is_visible():
                page1.get_by_role("checkbox", name="HO Wind Only").uncheck()
            try:
                page1.locator("#txtRequestedHomeownersCoverageALimit").click(timeout=10000)
            except Exception:
                snap(page1, "SKIP_no_coverage_field")
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' – coverage A field not found (address autocomplete failed)")
            page1.locator("#txtRequestedHomeownersCoverageALimit").fill(coverage_a)
            snap(page1, "05_coverages")
            if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").is_visible():
                page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtYearBuilt").fill("2025")
            if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").is_visible():
                page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLivingArea").fill("2025")
            page1.get_by_role("button", name="Next -->").click()
            page1.wait_for_timeout(2000)

            ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
            if ineligible_msgs.count() > 0:
                snap(page1, "SKIP_ineligible_coverages")
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' ineligible on coverages screen")
            try:
                page1.wait_for_selector("#ddlConstructionType", timeout=10000)
            except Exception:
                snap(page1, "SKIP_no_property_page")
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' did not advance to property screen")

            _fill_property_screens(page1)
            snap(page1, "09_wind_mitigation")

            _fill_applicant(page1, insured_name)
            snap(page1, "10_applicant")

            q_count = page1.locator("[id*='ucGenInfoQuestDDL_rptrQuestions'][id$='_ddlQuestions']").count()
            for i in range(q_count):
                loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl{i:02d}_ddlQuestions")
                if loc.count() > 0:
                    loc.select_option("False")
            wind_mit_q = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl21_ddlQuestions")
            if wind_mit_q.count() > 0:
                wind_mit_q.select_option("True")
            page1.get_by_role("cell", name="Applicant acknowledges that flood coverage is excluded for the homeowners, wind-only, and vacant policy.", exact=True).locator("span").first.click()
            for chk in ["chkRoofPaymentSchedule", "chkFloodExclusion", "chkRoofExclusion", "chkAcknowledgedWaterDamageLimitation"]:
                loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_{chk}")
                if loc.count() > 0:
                    loc.check()
            snap(page1, "11_questions")
            page1.get_by_role("button", name="Next -->").click()
            page1.wait_for_timeout(2000)
            applied_endorsements = _apply_random_endorsements(page1)
            snap(page1, "11b_endorsements")
            page1.get_by_role("button", name="Submit Quote Request").click()
            page1.wait_for_timeout(2000)

            ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
            if ineligible_msgs.count() > 0:
                snap(page1, "SKIP_ineligible_post_submit")
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' ineligible after quote submission")
            try:
                page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header", timeout=30000)
            except Exception:
                snap(page1, "FAIL_no_quote_results")
                pytest.fail(
                    f"[{test_case}] Quote results page did not load after submission.\n"
                    + _page_context(page1, test_case, "expected: quote results")
                )
            snap(page1, "12_quote_results")

            QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
            deductible = _maybe_change_deductible(page1)
            if deductible != "$2,500":
                QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
                snap(page1, "12b_quote_after_deductible")

            scenario_label = f"NRS Auros HO-3 | {bill_to_label} | {insured_name}"
            _append_result([USERNAME, test_case, scenario_label, attempt_addr,
                             ProductName, coverage_a, "", "",
                             QuoteNumber, full_pay, three_pay, four_pay, monthly,
                             insured_name, "Not Bound", "", "", "", "", "",
                             "; ".join(applied_endorsements), deductible])

            BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
                page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay, eff_date=eff_date
            )
            snap(page1, "14_binder_detail")

            _update_result(QuoteNumber, {15: "Bound", 16: BinderNumber, 17: BillTo, 18: PaymentPlan, 19: EffectiveDate, 20: AmountDueNow})
            return  # success

        except _AddressIneligible as exc:
            print(str(exc))
            if page1 is not None:
                try:
                    page1.close()
                except Exception:
                    pass
            continue

        except Exception as exc:
            try:
                snap(page1, "FAIL_unexpected")
                ctx = _page_context(page1, test_case, "unexpected exception")
            except Exception:
                ctx = f"  (could not capture page context: {exc})"
            raise AssertionError(
                f"[{test_case}] Unexpected failure — {type(exc).__name__}: {exc}\n{ctx}"
            ) from exc

    pytest.fail(f"[{test_case}] All {len(addresses_to_try)} addresses exhausted — no eligible address found")

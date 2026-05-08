import re
import os
import random
import pytest
from playwright.sync_api import Page
from helpers.result_logger import make_xlsx_appender
from pages.login_page import LoginPage

USERNAME = os.environ.get("AUROS_USERNAME", "icortesmarquez@geovera.com")
PORTAL_USERNAME = "icortesmarquez@geovera"
TEMP_POLICY_NUM = "GC60015733"
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "results", "auros", "auros_bind_results.xlsx")
BAD_ADDRESSES_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "results", "auros", "bad_addresses.txt")
QUOTE_URL = "https://tst.mygeosource.com/RHQ/QuotingHO/Quote.aspx"

EFFECTIVE_DATES = ["05/04/2026", "05/15/2026"]

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


SPECIFIC_BIND_SCENARIOS = [
    # test_case, address, eff_date, bill_to_val, bill_to_label, pay_plan_val, pay_plan_label, autopay
    ("3PayInvoiced",   "1214 Fury St, Davenport, FL 33837",                       "06/14/2026", "0", "Applicant",       "1", "3 Pay",    False),
    ("3PayAuto",       "1840 Salem Ct, Dunedin, FL 34698",                        "07/02/2026", "2", "Alternate Payor", "1", "3 Pay",    True),
    ("MonthlyPay",     "2150 Massachusetts Ave, Lanark Village, FL 32323",        "07/02/2026", "2", "Alternate Payor", "3", "Monthly",  True),
    ("PDKORCC",        "9916 Lone Tree Ln, Orlando, FL 32836",                    "05/22/2026", "0", "Applicant",       "0", "Full Pay", False),
    ("Renew",          "6123 Blue Runner Ct, Lakewood Ranch, FL 34202",           "06/29/2026", "0", "Applicant",       "0", "Full Pay", False),
    ("Refund",         "7 Folcroft Ln, Palm Coast, FL 32137",                     "05/19/2026", "2", "Alternate Payor", "1", "3 Pay",    True),
    ("Rebill_NSF_Fee", "522 Country Meadows Way, Bradenton, FL 34212",            "06/28/2026", "1", "Mortgagee",       "0", "Full Pay", False),
    ("ShortPay_EPay",  "1910 NE 17th Ter, Fort Lauderdale, FL 33305",             "07/06/2026", "0", "Applicant",       "0", "Full Pay", True),
    ("ShortPay_Inv",   "4524 5th Ave, St. Augustine, FL 32095",                   "06/19/2026", "1", "Mortgagee",       "0", "Full Pay", False),
    ("ToleranceUnder", "19433 Whispering Brook Dr, Tampa, FL 33647",              "06/25/2026", "0", "Applicant",       "0", "Full Pay", False),
]


class _AddressIneligible(Exception):
    pass


def _log_bad_address(address: str, reason: str):
    import time
    path = os.path.abspath(BAD_ADDRESSES_FILE)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} | {address} | {reason}\n")
    print(f"[BAD ADDRESS] {address} — {reason}")


def _try_addresses(addresses, flow, fail_prefix: str = "", catch_all: bool = False):
    """Iterate addresses, run flow(address) until one succeeds.
    On caught exception, log via _log_bad_address and continue.
    flow is responsible for closing its own page on failure (re-raise after cleanup)."""
    catchable = (_AddressIneligible, Exception) if catch_all else (_AddressIneligible,)
    for address in addresses:
        try:
            flow(address)
            return
        except catchable as exc:
            _log_bad_address(address, str(exc))
    pytest.fail(f"{fail_prefix}All {len(addresses)} addresses exhausted — no eligible address found")


_HEADERS = [
    "User Name", "Address", "Test Scenario", "Test Case", "Product Type", "Coverage A Limit",
    "Quote Number", "Quote Expiration", "Dwelling (Cov A)", "Total Annual Premium",
    "Full Pay", "3 Pay (1st Install)", "4 Pay (1st Install)", "Monthly (1st Install)",
    "Name (max 30)", "Second Name (max 30)", "Street (max 30)", "City (max 30)",
    "Zip (max 5)", "Zip+4 (max 4)", "Binder Number", "Bill To", "Payment Plan",
    "Effective Date", "Amount Due Now",
]
_append_result = make_xlsx_appender(RESULTS_FILE, _HEADERS)


FL_ADDRESSES = [
    "8734 Crestgate Cir, Orlando, FL 32819",
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
    "37 N Waterview Dr, Palm Coast, FL 32137",
    "3700 NW 20th Street, Coconut Creek, FL 33066",
    "85 GLENALBY Place, Ponte Vedra, FL 32081",
    "111 Midway Park Dr, Saint Augustine, FL 32084",
    "6104 Spencer Parrish Rd, Parrish, FL 34219",
    "5935 Old Berkley Rd, Auburndale, FL 33823",
    "11180 113th Ave, Seminole, FL 33778",
    "1339 Gallop Drive, Loxahatchee, FL 33470",
    "23305 June Bug Trl, Brooksville, FL 34602",
    "12614 Fenhurst WAY, NAPLES, FL 34120",
    "3433 Leaf Lake Dr, Land O Lakes, FL 34639",
    "18037 Mambo Drive, Boca Raton, FL 33496",
    "3651 Canopy CIR, NAPLES, FL 34120",
    "2650 SW Windship Way, Stuart, FL 34997",
    "5021 Wesley Dr, Tampa, FL 33647",
    "6032 Elm Rd, Sebring, FL 33875",
]


def _lookup_policy(page: Page, policy_num: str) -> Page:
    """Log in, search for a policy by number, open it in a popup, and click New Quote."""
    LoginPage(page).login(PORTAL_USERNAME, os.environ["INTERNAL_PASSWORD"])
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
    page1.locator("#ctl00_ctl00_MyGeoSourceMenun1").hover()
    page1.wait_for_timeout(1000)
    page1.get_by_role("link", name="New Quote").click()
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ddlWholesaler", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option("APPUND")
    page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill("APPAGT")
    page1.locator("#ctl00_WelcomeUserControl_btnUpdateQuotingAs").click()
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId").click()
    page1.wait_for_timeout(3000)
    page1.goto(QUOTE_URL)
    page1.wait_for_load_state("domcontentloaded")
    page1.wait_for_timeout(2000)
    return page1


def _login_and_open_quote(page: Page):
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
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ddlWholesaler", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option("APPUND")
    page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill("APPAGT")
    page1.locator("#ctl00_WelcomeUserControl_btnUpdateQuotingAs").click()
    page1.wait_for_selector("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId", timeout=30000)
    page1.locator("#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId").click()
    page1.wait_for_timeout(3000)
    page1.goto(QUOTE_URL)
    page1.wait_for_load_state("domcontentloaded")
    page1.wait_for_timeout(2000)
    return page1


def _enter_address(page1: Page, address: str):
    inp = page1.locator("#RadAutoCompleteBoxGeo_Input")
    inp.click()
    inp.press_sequentially(address, delay=50)
    try:
        page1.wait_for_selector(".rcbList .rcbItem", timeout=8000)
        page1.locator(".rcbList .rcbItem").first.click()
        page1.wait_for_timeout(1000)
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
    page1.wait_for_timeout(3000)
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
    page1.wait_for_timeout(3000)
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
    """Fill mortgagee block n (1 or 2). Name is a required RadComboBox dropdown; selecting it
    triggers a postback that populates the Street dropdown. Second Name and Loan # get
    generated fake values. City/State/Zip are left alone (autofilled by the system)."""
    pfx = "ctl00_ContentPlaceHolder1_ucBinderRequest"

    # Open Name RadComboBox and pick a random item
    name_arrow = f"#{pfx}_rcbMortgagee{n}Name1_Arrow"
    name_items = f"#{pfx}_rcbMortgagee{n}Name1_DropDown .rcbItem"
    page1.locator(name_arrow).click()
    page1.wait_for_selector(name_items, timeout=1500)
    items = page1.locator(name_items)
    items.nth(random.randint(0, items.count() - 1)).click()
    # Wait for postback to populate Street dropdown
    page1.wait_for_timeout(2500)

    # Fill Second Name with fake data
    second_name = random.choice(MORTGAGEE_SECOND_NAMES)
    page1.locator(f"#{pfx}_txtMortgagee{n}Name2").fill(second_name)

    # Open Street RadComboBox (populated after Name postback) and pick first option
    street_arrow = f"#{pfx}_rcbMortgagee{n}Street_Arrow"
    street_items = f"#{pfx}_rcbMortgagee{n}Street_DropDown .rcbItem"
    page1.locator(street_arrow).click()
    try:
        page1.wait_for_selector(street_items, timeout=1500)
        page1.locator(street_items).first.click()
        page1.wait_for_timeout(1000)
    except Exception:
        pass

    # Fill Loan # with a generated value
    loan_num = f"LN{random.randint(100000000, 999999999)}"
    page1.locator(f"#{pfx}_txtMortgagee{n}LoanNumber").fill(loan_num)


def _fill_alternate_payor(page1: Page):
    """Fill all Alternate Payor fields with generated fake data. State is a dropdown."""
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
    page1.wait_for_timeout(2000)
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

    # Handle autopay toggle for non-Monthly plans (Monthly always has autopay pre-selected)
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
    page1.wait_for_timeout(5000)
    return BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow


# ---------------------------------------------------------------------------
# Auros HO-3 — Quote to Bind
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay", BIND_SCENARIOS)
def test_bind_auros_ho3(page: Page, snap, bill_to_val: str, bill_to_label: str,
                        pay_plan_val: str, pay_plan_label: str, autopay: bool) -> None:
    coverage_a = "800,000"
    insured_name = f"Auros {BILL_TO_SHORT[bill_to_label]} {pay_plan_label}" + (" AutoPay" if autopay else "")

    def _flow(address: str):
        page1 = None
        try:
            page1 = _login_and_open_quote(page)
            snap(page1, "03_welcome")

            _enter_address(page1, address)
            snap(page1, "04_address")

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
                raise _AddressIneligible(f"'{address}' ineligible on coverages screen")
            try:
                page1.wait_for_selector("#ddlConstructionType", timeout=10000)
            except Exception:
                raise _AddressIneligible(f"'{address}' did not advance to property page")

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
            page1.get_by_role("button", name="Submit Quote Request").click()
            page1.wait_for_timeout(2000)

            ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
            if ineligible_msgs.count() > 0:
                raise _AddressIneligible(f"'{address}' ineligible after quote submit")
            try:
                page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header", timeout=30000)
            except Exception:
                raise _AddressIneligible(f"'{address}' quote results did not load")
            snap(page1, "12_quote_results")

            QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)

            BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
                page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay
            )
            snap(page1, "14_binder_detail")

            scenario_label = f"Auros HO-3 | {bill_to_label} | {insured_name}"
            _append_result([USERNAME, address, scenario_label, "", ProductName, coverage_a,
                             QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                             insured_name, "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow])
        except _AddressIneligible:
            if page1 is not None:
                try:
                    page1.close()
                except Exception:
                    pass
            raise

    _try_addresses(random.sample(FL_ADDRESSES, len(FL_ADDRESSES)), _flow)


# ---------------------------------------------------------------------------
# Auros HO-3T — Quote to Bind  (Occupancy = Tenant Only)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay", BIND_SCENARIOS)
def test_bind_auros_ho3t(page: Page, snap, bill_to_val: str, bill_to_label: str,
                         pay_plan_val: str, pay_plan_label: str, autopay: bool) -> None:
    coverage_a = "800,000"
    insured_name = f"Auros {BILL_TO_SHORT[bill_to_label]} {pay_plan_label}" + (" AutoPay" if autopay else "")

    def _flow(address: str):
        page1 = None
        try:
            page1 = _login_and_open_quote(page)
            snap(page1, "03_welcome")

            _enter_address(page1, address)
            page1.wait_for_timeout(2000)
            snap(page1, "04_address")

            page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlOccupancyType").select_option("Tenant Only")
            page1.wait_for_timeout(1500)
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
                raise _AddressIneligible(f"'{address}' ineligible on coverages screen")
            try:
                page1.wait_for_selector("#ddlConstructionType", timeout=30000)
            except Exception:
                raise _AddressIneligible(f"'{address}' did not advance to property page")

            try:
                _fill_property_screens(page1)
            except Exception as e:
                raise _AddressIneligible(f"'{address}' property screens failed — {e}")
            snap(page1, "09_wind_mitigation")

            try:
                _fill_applicant(page1, insured_name)
            except Exception as e:
                raise _AddressIneligible(f"'{address}' applicant screen failed — {e}")
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
            page1.get_by_role("button", name="Submit Quote Request").click()
            page1.wait_for_timeout(2000)

            ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
            if ineligible_msgs.count() > 0:
                raise _AddressIneligible(f"'{address}' ineligible after quote submit")
            try:
                page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header", timeout=30000)
            except Exception:
                raise _AddressIneligible(f"'{address}' quote results did not load")
            snap(page1, "12_quote_results")

            QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
            if not any(x in ProductName for x in ("3T", "HO-3T", "Tenant")):
                raise _AddressIneligible(f"'{address}' returned '{ProductName}' instead of HO-3T")

            BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
                page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay
            )
            snap(page1, "14_binder_detail")

            scenario_label = f"Auros HO-3T | {bill_to_label} | {insured_name}"
            _append_result([USERNAME, address, scenario_label, "", ProductName, coverage_a,
                             QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                             insured_name, "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow])
        except _AddressIneligible:
            if page1 is not None:
                try:
                    page1.close()
                except Exception:
                    pass
            raise

    _try_addresses(random.sample(FL_ADDRESSES, len(FL_ADDRESSES)), _flow)


# ---------------------------------------------------------------------------
# Auros HO-3 — Specific Named Scenarios
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "test_case,address,eff_date,bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay",
    SPECIFIC_BIND_SCENARIOS,
    ids=[s[0] for s in SPECIFIC_BIND_SCENARIOS],
)
def test_bind_auros_ho3_specific(page: Page, snap, test_case: str, address: str, eff_date: str,
                                  bill_to_val: str, bill_to_label: str,
                                  pay_plan_val: str, pay_plan_label: str, autopay: bool) -> None:
    coverage_a = "800,000"
    pay_plan_short = {"Full Pay": "Full", "3 Pay": "3Pay", "Monthly": "Monthly"}[pay_plan_label]
    pay_type_short = "EPay" if autopay else "Inv"
    insured_name = f"{test_case} {pay_plan_short} {pay_type_short}"

    fallback_pool = [a for a in FL_ADDRESSES if a != address]
    random.shuffle(fallback_pool)
    addresses_to_try = [address] + fallback_pool

    def _flow(attempt_addr: str):
        page1 = None
        try:
            # page1 = _login_and_open_quote(page)
            page1 = _lookup_policy(page, TEMP_POLICY_NUM)
            snap(page1, "03_welcome")

            _enter_address(page1, attempt_addr)
            snap(page1, "04_address")

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
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' ineligible on coverages screen")
            try:
                page1.wait_for_selector("#ddlConstructionType", timeout=10000)
            except Exception:
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' did not advance to property page")

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
            page1.get_by_role("button", name="Submit Quote Request").click()
            page1.wait_for_timeout(2000)

            ineligible_msgs = page1.get_by_text(re.compile(r"ineligible", re.IGNORECASE))
            if ineligible_msgs.count() > 0:
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' ineligible after quote submit")
            try:
                page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header", timeout=30000)
            except Exception:
                raise _AddressIneligible(f"[{test_case}] '{attempt_addr}' quote results did not load")
            snap(page1, "12_quote_results")

            QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)

            BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
                page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay, eff_date=eff_date
            )
            snap(page1, "14_binder_detail")

            scenario_label = f"Auros HO-3 | {bill_to_label} | {insured_name}"
            _append_result([USERNAME, attempt_addr, scenario_label, test_case, ProductName, coverage_a,
                             QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                             insured_name, "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow])
        except (_AddressIneligible, Exception):
            if page1 is not None:
                try:
                    page1.close()
                except Exception:
                    pass
            raise

    _try_addresses(addresses_to_try, _flow, fail_prefix=f"[{test_case}] ", catch_all=True)

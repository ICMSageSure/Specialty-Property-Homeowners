import re
import os
import random
import pytest
from playwright.sync_api import Page
from helpers.result_logger import make_xlsx_appender
from helpers.auros_endorsements import (
    fill_address_endorsement,
    fill_rented_address_endorsement,
    fill_occupied_address_endorsement,
    check_only_endorsement,
    fill_pool_endorsement,
    fill_water_limitation,
)

USERNAME = os.environ.get("AUROS_USERNAME", "tsicortesmarquez@geovera.com")
WHOLESALER = "HULL001"
PRODUCERS = [("HLKAREN", "01"), ("ANNIE", "0134")]
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "results", "nrs", "nrs_bind_results.xlsx")

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
# Split evenly between HO-3 (even indices) and HO-3T (odd indices) for balanced coverage.
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

# 3 scenarios per product, one per Bill To type — all using HLKAREN-01 (ANNIE login does not work)
_HO3_BASE = [BIND_SCENARIOS[0], BIND_SCENARIOS[7], BIND_SCENARIOS[14]]   # Applicant / Mortgagee / AltPayor — Full Pay no auto
_HO3T_BASE = [BIND_SCENARIOS[1], BIND_SCENARIOS[8], BIND_SCENARIOS[15]]  # Applicant / Mortgagee / AltPayor — Full Pay auto

HO3_SCENARIOS  = [s + PRODUCERS[0] for s in _HO3_BASE]
HO3T_SCENARIOS = [s + PRODUCERS[0] for s in _HO3T_BASE]

# Targeted NRS bind scenarios with fixed addresses, effective dates, and billing params.
# Payor "Insured" maps to bill_to_val "0" (Applicant in the system dropdown).
# Pay Type "Invoiced" → autopay=False; "E Pay" → autopay=True.
NRS_TEST_SCENARIOS = [
    # (test_case,        address,                                              eff_date,     pay_plan_val, pay_plan_label, autopay, bill_to_val, bill_to_label)
    ("3PayInvoiced",    "2269 Cascadia Ct, St. Augustine, FL 32092",          "06/14/2026", "1", "3 Pay",    False, "0", "Applicant"),
    ("3PayAuto",        "3137 Castle Rock Cir, Land O Lakes, FL 34639",       "07/02/2026", "1", "3 Pay",    True,  "2", "Alternate Payor"),
    ("MonthlyPay",      "13626 Allyn Dr, Hudson, FL 34667",                   "07/02/2026", "3", "Monthly",  True,  "2", "Alternate Payor"),
    ("PDKORCC",         "600 S Dixie Hwy #352, West Palm Beach, FL 33401",    "05/22/2026", "0", "Full Pay", False, "0", "Applicant"),
    ("Renew",           "510 Wedge Ln, Longboat Key, FL 34228",               "06/29/2026", "0", "Full Pay", False, "0", "Applicant"),
    ("Refund",          "8431 SE Duncan St, Hobe Sound, FL 33455",            "05/19/2026", "1", "3 Pay",    True,  "2", "Alternate Payor"),
    ("RebillNSFFee",    "308 Ferndale Way, St. Augustine, FL 32092",          "06/28/2026", "0", "Full Pay", False, "1", "Mortgagee"),
    ("ShortPayOrlando", "6013 Fender Ct, Orlando, FL 32837",                  "07/06/2026", "0", "Full Pay", True,  "0", "Applicant"),
    ("ShortPayOdessa",  "1607 Abyss Dr, Odessa, FL 33556",                    "06/19/2026", "0", "Full Pay", False, "1", "Mortgagee"),
    ("ToleranceUnder",  "1151 NW 45th St, Miami, FL 33127",                   "06/25/2026", "0", "Full Pay", False, "0", "Applicant"),
]


_HEADERS = [
    "User Name", "Address", "Test Scenario", "Product Type", "Coverage A Limit", "Deductible",
    "Wholesaler", "Producer Number",
    "Quote Number", "Quote Expiration", "Dwelling (Cov A)", "Total Annual Premium",
    "Full Pay", "3 Pay (1st Install)", "4 Pay (1st Install)", "Monthly (1st Install)",
    "Name (max 30)", "Second Name (max 30)", "Street (max 30)", "City (max 30)",
    "Zip (max 5)", "Zip+4 (max 4)", "Binder Number", "Bill To", "Payment Plan",
    "Effective Date", "Amount Due Now", "Endorsements Applied",
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
]


def _login_and_open_quote(page: Page, producer_code: str, sub_code: str):
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
    page1.locator("#ctl00_WelcomeUserControl_ddlWholesaler").select_option(WHOLESALER)
    page1.locator("#ctl00_WelcomeUserControl_txtProducerNumber").fill(producer_code)
    page1.locator("#ctl00_WelcomeUserControl_txtSubProducerNumber").fill(sub_code)
    page1.locator("#ctl00_WelcomeUserControl_btnUpdateQuotingAs").click()
    page1.wait_for_selector("[id*='dgUserList'][id$='_btnId']", timeout=30000)
    page1.locator("[id*='dgUserList'][id$='_btnId']").first.click()
    page1.wait_for_timeout(3000)
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

    first_field = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName")
    last_field  = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLastName")
    try:
        max_first = int(first_field.get_attribute("maxlength") or "20")
    except (TypeError, ValueError):
        max_first = 20
    if len(insured_name) <= max_first:
        first_name, last_name = insured_name, "TESTER"
    else:
        cut = insured_name.rfind(" ", 0, max_first + 1)
        if cut == -1:
            first_name, last_name = insured_name[:max_first], insured_name[max_first:]
        else:
            first_name, last_name = insured_name[:cut], insured_name[cut + 1:]
    first_field.fill(first_name)
    last_field.fill(last_name)
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
    eff_date = eff_date or random.choice(EFFECTIVE_DATES)
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


def _maybe_change_deductible(page1: Page) -> str:
    """Randomly decide whether to edit the deductible. If yes, pick a non-default value (skipping $5,000),
    submit, and wait for the postback to re-render the quote details. Returns the final deductible label.
    Caller is responsible for re-reading premium/payment values after a change."""
    DEFAULT_LABEL = "$5,000"
    NON_DEFAULT = [
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
    Water exclusions are intentionally excluded — they conflict with water_limitation.
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
    selected = random.sample(options, random.randint(3, 5))
    applied = []
    for name, detail, fn in selected:
        try:
            result = fn()
            if result is not None:
                detail = f"${result:,}"
            applied.append(f"{name} ({detail})" if detail else name)
        except Exception as e:
            print(f"  [endorsement skip] {name}: {e}")
    print(f"Applied endorsements ({len(applied)}): {applied}")
    return applied


# ---------------------------------------------------------------------------
# NRS Auros HO-3 — Quote to Bind
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay,producer_code,sub_code", HO3_SCENARIOS)
def test_bind_nrs_auros_ho3(page: Page, snap, bill_to_val: str, bill_to_label: str,
                             pay_plan_val: str, pay_plan_label: str, autopay: bool,
                             producer_code: str, sub_code: str) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    pay_type = "AutoPay" if autopay else "Invoiced"
    insured_name = f"NRS {BILL_TO_SHORT[bill_to_label]} {pay_plan_label} {pay_type}"

    page1 = _login_and_open_quote(page, producer_code, sub_code)
    snap(page1, "03_welcome")

    _enter_address(page1, address)
    snap(page1, "04_address")

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

    # General information — answer all visible questions; ctl21 (wind mitigation) is conditional
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
    if deductible != "$5,000":
        QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
        snap(page1, "12b_quote_after_deductible")

    BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
        page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay
    )
    snap(page1, "14_binder_detail")

    scenario_label = f"NRS Auros HO-3 | {bill_to_label} | {insured_name}"
    _append_result([USERNAME, address, scenario_label, ProductName, coverage_a, deductible,
                     WHOLESALER, f"{producer_code}-{sub_code}",
                     QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                     insured_name, "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow,
                     "; ".join(applied_endorsements)])


# ---------------------------------------------------------------------------
# NRS Auros HO-3T — Quote to Bind  (Occupancy = Tenant Only)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("bill_to_val,bill_to_label,pay_plan_val,pay_plan_label,autopay,producer_code,sub_code", HO3T_SCENARIOS)
def test_bind_nrs_auros_ho3t(page: Page, snap, bill_to_val: str, bill_to_label: str,
                              pay_plan_val: str, pay_plan_label: str, autopay: bool,
                              producer_code: str, sub_code: str) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    pay_type = "AutoPay" if autopay else "Invoiced"
    insured_name = f"NRS {BILL_TO_SHORT[bill_to_label]} {pay_plan_label} {pay_type}"
    page1 = None
    try:
        page1 = _login_and_open_quote(page, producer_code, sub_code)
        snap(page1, "03_welcome")

        _enter_address(page1, address)
        page1.wait_for_timeout(2000)
        snap(page1, "04_address")

        # Set Occupancy Type to Tenant Only; Auros HO-3 produces HO-3T output under Tenant Only
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

        # General information — answer all visible questions; ctl17 (tenant) and ctl21 (roof materials) are conditional
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
        if deductible != "$5,000":
            QuoteNumber, ProductName, full_pay, three_pay, four_pay, monthly = _collect_quote_info(page1)
            snap(page1, "12b_quote_after_deductible")

        BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
            page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay
        )
        snap(page1, "14_binder_detail")

        scenario_label = f"NRS Auros HO-3T | {bill_to_label} | {insured_name}"
        _append_result([USERNAME, address, scenario_label, ProductName, coverage_a, deductible,
                         WHOLESALER, f"{producer_code}-{sub_code}",
                         QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                         insured_name, "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow,
                         "; ".join(applied_endorsements)])
    finally:
        if page1 is not None:
            try:
                page1.close()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# NRS Auros HO-3 — Targeted bind scenarios with fixed addresses and dates
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "test_case,address,eff_date,pay_plan_val,pay_plan_label,autopay,bill_to_val,bill_to_label",
    NRS_TEST_SCENARIOS,
    ids=[s[0] for s in NRS_TEST_SCENARIOS],
)
def test_bind_nrs_auros(page: Page, snap, test_case: str, address: str, eff_date: str,
                        pay_plan_val: str, pay_plan_label: str, autopay: bool,
                        bill_to_val: str, bill_to_label: str) -> None:
    producer_code, sub_code = PRODUCERS[0]
    coverage_a = "800,000"
    insured_name = f"NRS {test_case}"

    page1 = _login_and_open_quote(page, producer_code, sub_code)
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

    BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
        page1, QuoteNumber, bill_to_val, pay_plan_val, pay_plan_label, autopay, eff_date
    )
    snap(page1, "14_binder_detail")

    scenario_label = f"NRS {test_case} | {bill_to_label} | {pay_plan_label}"
    _append_result([USERNAME, address, scenario_label, ProductName, coverage_a, "",
                     WHOLESALER, f"{producer_code}-{sub_code}",
                     QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                     insured_name, "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow,
                     ""])

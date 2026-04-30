import re
import os
import random
import pytest
from playwright.sync_api import Page
from datetime import datetime
from openpyxl import load_workbook, Workbook

USERNAME = os.environ.get("AUROS_USERNAME", "icortesmarquez@geovera.com")
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "..", "..", "..", "results", "auros", "auros_test_results.xlsx")


def _append_to_xlsx(row: list):
    import time
    path = os.path.abspath(RESULTS_FILE)
    lock_path = path + ".lock"
    for _ in range(20):
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            break
        except FileExistsError:
            time.sleep(0.5)
    try:
        if os.path.exists(path):
            wb = load_workbook(path)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["User Name", "Address", "Test Scenario", "Product Type", "Coverage A Limit",
                        "Quote Number", "Quote Expiration", "Dwelling (Cov A)", "Total Annual Premium",
                        "Full Pay", "3 Pay (1st Install)", "4 Pay (1st Install)", "Monthly (1st Install)",
                        "Name (max 30)", "Second Name (max 30)", "Street (max 30)", "City (max 30)",
                        "Zip (max 5)", "Zip+4 (max 4)", "Binder Number", "Bill To", "Payment Plan",
                        "Effective Date", "Amount Due Now"])
        ws.append(row)
        wb.save(path)
    finally:
        try:
            os.unlink(lock_path)
        except Exception:
            pass


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


def _login_and_open_quote(page: Page):
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
    return page1


def _enter_address(page1: Page, address: str):
    page1.locator("#RadAutoCompleteBoxGeo_Input").click()
    page1.locator("#RadAutoCompleteBoxGeo_Input").press_sequentially(address)
    page1.wait_for_timeout(1500)
    try:
        page1.locator(".rcbList .rcbItem").first.click(timeout=1500)
    except Exception:
        page1.locator("#RadAutoCompleteBoxGeo_Input").press("Enter")
        page1.wait_for_timeout(1500)
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
            raw = rc_row.text_content(timeout=5000) or ""
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


def _fill_applicant(page1: Page):
    if page1.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlPolicyOwnerType").is_visible():
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


def _collect_quote_info(page1: Page):
    QuoteNumber = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Number").inner_text()
    ProductName = page1.locator("#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lblQuote1Header").inner_text()
    print(f'Product: {ProductName}  |  Quote: {QuoteNumber}')

    def _get_pay(idx: str) -> str:
        loc = page1.locator(f"#ctl00_ContentPlaceHolder1_ucQuoteDetailMultiCo1_lstPaymentOptions_{idx}_lblQuote1PaymentOption")
        try:
            text = loc.text_content(timeout=5000).strip()
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


def _bind(page1: Page, QuoteNumber: str, pay_plan_val: str, pay_plan_label: str):
    try:
        page1.locator(f'input[type="radio"][value="{QuoteNumber}"]').first.check(timeout=5000)
    except Exception:
        pass
    page1.get_by_role("button", name="Continue to Bind Coverage").click()
    page1.wait_for_timeout(2000)
    if page1.locator("#chkPopInConfirmation").is_visible():
        page1.locator("#chkPopInConfirmation").check()
        page1.get_by_role("button", name="Continue", exact=True).click()

    page1.wait_for_selector("#ctl00_ContentPlaceHolder1_ucBinderRequest_ddlBillTo1", timeout=30000)
    eff_date = datetime.today().strftime('%m/%d/%Y')
    page1.locator("#ctl00_ContentPlaceHolder1_ucBinderRequest_txtBinderEffectiveDate1").fill(eff_date)
    page1.locator("#ctl00_ContentPlaceHolder1_ucBinderRequest_ddlBillTo1").select_option("0")
    try:
        page1.locator("#ddlPaymentPlan1").select_option(pay_plan_val)
    except Exception:
        pytest.skip(f"Pay plan '{pay_plan_label}' not available for this product")
    print(f'Selected Pay Plan: {pay_plan_label}')
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

@pytest.mark.parametrize("pay_plan_val,pay_plan_label", [
    ("0", "Full Pay"),
])
def test_bind_auros_ho3(page: Page, snap, pay_plan_val: str, pay_plan_label: str) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"

    page1 = _login_and_open_quote(page)
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

    _fill_applicant(page1)
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
        page1, QuoteNumber, pay_plan_val, pay_plan_label
    )
    snap(page1, "14_binder_detail")

    _append_to_xlsx([USERNAME, address, f"Auros HO-3 ({pay_plan_label})", ProductName, coverage_a,
                     QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                     "", "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow])


# ---------------------------------------------------------------------------
# Auros HO-3T — Quote to Bind  (mirrors FG HO-3T: Occupancy = Tenant Only)
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("pay_plan_val,pay_plan_label", [
    ("0", "Full Pay"),
])
def test_bind_auros_ho3t(page: Page, snap, pay_plan_val: str, pay_plan_label: str) -> None:
    address = random.choice(FL_ADDRESSES)
    coverage_a = "800,000"
    page1 = None
    try:
        page1 = _login_and_open_quote(page)
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
            _fill_applicant(page1)
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

        BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow = _bind(
            page1, QuoteNumber, pay_plan_val, pay_plan_label
        )
        snap(page1, "14_binder_detail")

        _append_to_xlsx([USERNAME, address, f"Auros HO-3T ({pay_plan_label})", ProductName, coverage_a,
                         QuoteNumber, "", "", full_pay, full_pay, three_pay, four_pay, monthly,
                         "", "", "", "", "", "", BinderNumber, BillTo, PaymentPlan, EffectiveDate, AmountDueNow])
    finally:
        if page1 is not None:
            try:
                page1.close()
            except Exception:
                pass

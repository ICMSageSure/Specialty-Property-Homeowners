import os
from playwright.sync_api import Page

UPLOAD_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "test_data", "testfile.pdf"))


class AddressStep:
    def __init__(self, page: Page):
        self.page = page

    def enter_address(self, address: str):
        self.page.locator("#RadAutoCompleteBoxGeo_Input").click()
        self.page.locator("#RadAutoCompleteBoxGeo_Input").press_sequentially(address)
        self.page.wait_for_timeout(3000)
        self.page.locator("#RadAutoCompleteBoxGeo_Input").press("Enter")
        self.page.wait_for_timeout(2000)
        self.page.get_by_role("button", name="Next -->").click()


class CoveragesStep:
    def __init__(self, page: Page):
        self.page = page

    def select_wind_only(self, coverage_a: str):
        """Uncheck both HO-3 products, keeping Wind Only checked."""
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_chkProductType_0").uncheck()  # Auros HO-3
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_chkProductType_1").uncheck()  # GeoVera Specialty HO-3
        self.page.locator("#txtRequestedHomeownersCoverageALimit").fill(coverage_a)
        self.page.get_by_role("button", name="Next -->").click()

    def select_ho3(self, coverage_a: str):
        """Uncheck Auros HO-3 and Wind Only, keeping GeoVera Specialty HO-3 checked."""
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_chkProductType_0").uncheck()  # Auros HO-3
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_chkProductType_2").uncheck()  # GeoVera Specialty HO Wind Only
        self.page.locator("#txtRequestedHomeownersCoverageALimit").fill(coverage_a)
        self.page.get_by_role("button", name="Next -->").click()


class PropertyStep:
    def __init__(self, page: Page):
        self.page = page

    def fill(self, construction: str = "FR", siding: str = "ST", restricted_roof_shapes: str = "False"):
        self.page.locator("#ddlConstructionType").select_option(construction)
        self.page.locator("#ddlSidingTypeDominant").select_option(siding)
        self.page.locator("#ddlRestrictedRoofShapes").select_option(restricted_roof_shapes)
        if self.page.locator("#ddlRoofShapeDominant").is_visible():
            self.page.locator("#ddlRoofShapeDominant").select_option("2")
        self.page.get_by_role("button", name="Next -->").click()


class ProtectionStep:
    def __init__(self, page: Page):
        self.page = page

    def fill(self):
        self.page.wait_for_timeout(3000)
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralFireAlarm").select_option("False")
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlCentralBurglarAlarm").select_option("False")
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlGatedGuardedCommunity").select_option("False")
        if self.page.locator("#ddlSwimmingPool").is_visible():
            self.page.locator("#ddlSwimmingPool").select_option(label="No Swimming Pool")
        self.page.get_by_role("button", name="Next -->").click()


class PlumbingStep:
    def __init__(self, page: Page):
        self.page = page

    def fill(self):
        self.page.locator("#ddlPolybutylenePipes").select_option("False")
        self.page.locator("#ddlAutomaticWaterShutoffValve").select_option("False")
        self.page.get_by_role("button", name="Next -->").click()


class WindMitigationStep:
    def __init__(self, page: Page):
        self.page = page

    def fill(self, month: str = "01", day: str = "01", year: str = "2022"):
        ddl = self.page.locator(
            "#ctl00_ContentPlaceHolder1_ctl05_ddlWindMitigationVerificationInspection"
        )
        if ddl.is_visible():
            ddl.select_option("Yes")
            self.page.locator(
                "#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionMonth"
            ).fill(month)
            self.page.locator(
                "#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionDay"
            ).fill(day)
            self.page.locator(
                "#ctl00_ContentPlaceHolder1_ctl05_txtWindMitigationVerificationInspectionYear"
            ).fill(year)
            self.page.locator("#fileWindMitigationVerificationInspectionForm").set_input_files(UPLOAD_PATH)
        self.page.get_by_role("button", name="Next -->").click()


class ApplicantStep:
    def __init__(self, page: Page):
        self.page = page

    def fill(
        self,
        owner_type: str = "Individual",
        first: str = "ICMTEST",
        last: str = "TESTER",
        dob_month: str = "01",
        dob_day: str = "01",
        dob_year: str = "1998",
        ssn1: str = "123",
        ssn2: str = "11",
        ssn3: str = "1122",
        prior_insurance: str = "False",
    ):
        owner_loc = self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_ddlPolicyOwnerType")
        if owner_loc.is_visible():
            owner_loc.select_option(owner_type)
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName").click()
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName").press("CapsLock")
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtFirstName").press_sequentially(first)
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLastName").click()
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtLastName").press_sequentially(last)
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBMonth").click()
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBMonth").press_sequentially(dob_month)
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBDay").press_sequentially(dob_day)
        self.page.locator("#ctl00_ContentPlaceHolder1_ctl05_txtDOBYear").press_sequentially(dob_year)
        self.page.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN1\"]").press_sequentially(ssn1)
        self.page.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN2\"]").press_sequentially(ssn2)
        self.page.locator("input[name=\"ctl00\\$ContentPlaceHolder1\\$ctl05\\$txtSSN3\"]").press_sequentially(ssn3)
        self.page.get_by_role("checkbox", name="Is Mailing Address the same").check()
        self.page.locator("#ddlPriorHomeownersInsurance").select_option(prior_insurance)
        self.page.wait_for_timeout(3000)
        self.page.locator("#ddlPriorHomeownersInsurance").select_option(prior_insurance)
        self.page.get_by_role("button", name="Next -->").click()


class QuestionsStep:
    _QUESTION_BASE = (
        "#ctl00_ContentPlaceHolder1_ctl05_ucGenInfoQuestDDL_rptrQuestions_ctl{:02d}_ddlQuestions"
    )
    _FLOOD_CELL = (
        "Applicant acknowledges that flood coverage is excluded for the homeowners, "
        "wind-only, and vacant policy."
    )
    _OPTIONAL_CHECKBOXES = [
        "#ctl00_ContentPlaceHolder1_ctl05_chkRoofPaymentSchedule",
        "#ctl00_ContentPlaceHolder1_ctl05_chkFloodExclusion",
        "#ctl00_ContentPlaceHolder1_ctl05_chkRoofExclusion",
        "#ctl00_ContentPlaceHolder1_ctl05_chkAcknowledgedWaterDamageLimitation",
    ]

    def __init__(self, page: Page):
        self.page = page

    def answer_all_false_and_submit(self, question_count: int = 21):
        for i in range(question_count):
            self.page.locator(self._QUESTION_BASE.format(i)).select_option("False")
        self.page.get_by_role("cell", name=self._FLOOD_CELL, exact=True).locator("span").first.click()
        for checkbox_id in self._OPTIONAL_CHECKBOXES:
            loc = self.page.locator(checkbox_id)
            if loc.is_visible():
                loc.check()
        self.page.get_by_role("button", name="Submit Quote Request").click()

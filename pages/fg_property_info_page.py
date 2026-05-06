from playwright.sync_api import Page

from pages.fg_policy_base_page import FGPolicyBasePage


class FGPropertyInfoPage(FGPolicyBasePage):
    _PREFIX = "#ctl00_ctl00_ContentPlaceHolder1_PolicyContentPlaceHolder2_"
    _EFF_DATE_PREFIX = f"{_PREFIX}ucRequestedEffectiveDate_"

    # -- Buttons --
    _BTN_EDIT1    = f"{_PREFIX}btnEdit1"
    _BTN_CONTINUE = f"{_PREFIX}btnContinue"

    # -- Effective date section --
    _TXT_RECEIVED_DATE    = f"{_EFF_DATE_PREFIX}txtReceivedDate"
    _TXT_EFFECTIVE_DATE   = f"{_EFF_DATE_PREFIX}txtChangeEffectiveDate"
    _DDL_CHANGE_REQUESTOR = f"{_EFF_DATE_PREFIX}ddlChangeRequestor"

    # -- Building / units --
    _DDL_UNIT_TYPE        = f"{_PREFIX}ddlUnitType"
    _DDL_UNITS_IN_BUILDING = f"{_PREFIX}ddlNumberOfUnitsInBuilding"
    _DDL_UNITS_INSURED    = f"{_PREFIX}ddlNumberOfUnitsInsured"

    # -- Construction --
    _DDL_CONSTRUCTION_TYPE        = f"{_PREFIX}ddlConstructionType"
    _DDL_DWELLING_TYPE            = f"{_PREFIX}ddlDwellingType"
    _DDL_OCCUPANCY_TYPE           = f"{_PREFIX}ddlOccupancyType"
    _DDL_VACANCY_REASON           = f"{_PREFIX}ddlVacancyReason"
    _TXT_EXPECTED_OCCUPANCY_DATE  = f"{_PREFIX}txtExpectedOccupancyDate"

    # -- Year built --
    _TXT_YEAR_BUILT          = "#txtYearBuilt"
    _CHK_OVERRIDE_YEAR_BUILT = "#chkOverrideApprovalYearBuilt"

    # -- Building area --
    _TXT_LIVING_AREA          = f"{_PREFIX}txtLivingArea"
    _CHK_OVERRIDE_SQUARE_FEET = "#chkOverrideApprovalSquareFeet"

    # -- Foundation --
    _DDL_FOUNDATION_DOMINANT = "#ddlFoundationTypeDominant"

    # -- Structure --
    _DDL_LEVELS                = f"{_PREFIX}ddlLevels"
    _DDL_ROOF_SHAPE_DOMINANT   = "#ddlRoofShapeDominant"
    _DDL_RESTRICTED_ROOF_SHAPES = "#ddlRestrictedRoofShapes"
    _DDL_ROOF_MATERIAL_DOMINANT = "#ddlRoofMaterialDominant"
    _DDL_SIDING_TYPE_DOMINANT   = "#ddlSidingTypeDominant"

    # -- Garage --
    _DDL_GARAGE_TYPE = f"{_PREFIX}ddlGarageType"
    _DDL_GARAGE_SIZE = f"{_PREFIX}ddlGarageSize"

    # -- Protection / wind --
    _DDL_PROTECTION_CLASS          = "#ddlProtectionClass"
    _CHK_OVERRIDE_PROTECTION_CLASS = "#chkOverrideApprovalProtectionClass"
    _DDL_GABLE_BRACING             = "#ddlGableBracing"
    _DDL_HURRICANE_STRAPS          = f"{_PREFIX}ddlHurricaneStraps"
    _DDL_HURRICANE_SHUTTERS        = f"{_PREFIX}ddlHurricaneShutters"
    _DDL_WIND_MITIGATION           = "#ddlWindMitigationVerificationInspection"

    # -- Alarms & community --
    _DDL_CENTRAL_FIRE_ALARM      = f"{_PREFIX}ddlCentralFireAlarm"
    _DDL_CENTRAL_BURGLAR_ALARM   = f"{_PREFIX}ddlCentralBurglarAlarm"
    _DDL_GATED_GUARDED_COMMUNITY = f"{_PREFIX}ddlGatedGuardedCommunity"
    _DDL_SPRINKLERS              = "#ddlSprinklers"
    _DDL_SWIMMING_POOL           = "#ddlSwimmingPool"

    # -- Safety features --
    _DDL_SMOKE_DETECTORS  = f"{_PREFIX}ddlSmokeDetectors"
    _DDL_DEADBOLT_LOCKS   = f"{_PREFIX}ddlDeadboltLocks"
    _DDL_FIRE_EXTINGUISHER = f"{_PREFIX}ddlFireExtinguisher"

    # -- Distances --
    _TXT_DISTANCE_FIRE_STATION = f"{_PREFIX}txtDistanceToFireStation"
    _TXT_DISTANCE_FIRE_HYDRANT = f"{_PREFIX}txtDistanceToFireHydrant"

    # -- Update years --
    _TXT_WIRING_YEAR       = f"{_PREFIX}txtWiringYear"
    _TXT_PLUMBING_YEAR     = f"{_PREFIX}txtPlumbingYear"
    _TXT_HEATING_YEAR      = f"{_PREFIX}txtHeatingYear"
    _TXT_ROOF_YEAR         = f"{_PREFIX}txtRoofYear"
    _CHK_OVERRIDE_ROOF_YEAR = "#chkOverrideApprovalRoofYear"

    # -- Renovation --
    _DDL_WILL_RENOVATE    = f"{_PREFIX}ddlWillRenovate"
    _DDL_RENOVATION_TYPE  = f"{_PREFIX}ddlRenovationType"

    def __init__(self, page: Page) -> None:
        super().__init__(page)

    # ------------------------------------------------------------------ #
    # Navigation / edit mode                                               #
    # ------------------------------------------------------------------ #

    def click_edit(self) -> None:
        self.page.locator(self._BTN_EDIT1).click()

    def click_continue(self) -> None:
        self.page.locator(self._BTN_CONTINUE).click()

    # ------------------------------------------------------------------ #
    # Effective date section                                               #
    # ------------------------------------------------------------------ #

    def fill_effective_date_section(self, requestor: str) -> None:
        self.page.locator(self._TXT_RECEIVED_DATE).click()
        self.page.locator(self._TXT_EFFECTIVE_DATE).click()
        self.page.locator(self._DDL_CHANGE_REQUESTOR).select_option(requestor)

    # ------------------------------------------------------------------ #
    # Building / units                                                     #
    # ------------------------------------------------------------------ #

    def select_unit_type(self, value: str) -> None:
        self.page.locator(self._DDL_UNIT_TYPE).select_option(value)

    def select_units_in_building(self, value: str) -> None:
        self.page.locator(self._DDL_UNITS_IN_BUILDING).select_option(value)

    def select_units_insured(self, value: str) -> None:
        self.page.locator(self._DDL_UNITS_INSURED).select_option(value)

    # ------------------------------------------------------------------ #
    # Construction                                                         #
    # ------------------------------------------------------------------ #

    def select_construction_type(self, value: str) -> None:
        self.page.locator(self._DDL_CONSTRUCTION_TYPE).select_option(value)

    def select_dwelling_type(self, value: str) -> None:
        self.page.locator(self._DDL_DWELLING_TYPE).select_option(value)

    def select_occupancy_type(self, value: str) -> None:
        self.page.locator(self._DDL_OCCUPANCY_TYPE).select_option(value)

    def is_vacancy_reason_visible(self) -> bool:
        return self.page.locator(self._DDL_VACANCY_REASON).is_visible()

    def select_vacancy_reason(self, value: str) -> None:
        self.page.locator(self._DDL_VACANCY_REASON).select_option(value)

    def click_expected_occupancy_date(self) -> None:
        self.page.locator(self._TXT_EXPECTED_OCCUPANCY_DATE).click()

    # ------------------------------------------------------------------ #
    # Year built                                                           #
    # ------------------------------------------------------------------ #

    def click_year_built(self) -> None:
        self.page.locator(self._TXT_YEAR_BUILT).click()

    def check_override_year_built(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_YEAR_BUILT).check()

    def uncheck_override_year_built(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_YEAR_BUILT).uncheck()

    # ------------------------------------------------------------------ #
    # Building area                                                        #
    # ------------------------------------------------------------------ #

    def click_living_area(self) -> None:
        self.page.locator(self._TXT_LIVING_AREA).click()

    def check_override_square_feet(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_SQUARE_FEET).check()

    def uncheck_override_square_feet(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_SQUARE_FEET).uncheck()

    # ------------------------------------------------------------------ #
    # Foundation                                                           #
    # ------------------------------------------------------------------ #

    def select_foundation_dominant(self, value: str) -> None:
        self.page.locator(self._DDL_FOUNDATION_DOMINANT).select_option(value)

    # ------------------------------------------------------------------ #
    # Structure                                                            #
    # ------------------------------------------------------------------ #

    def select_levels(self, value: str) -> None:
        self.page.locator(self._DDL_LEVELS).select_option(value)

    def select_roof_shape(self, value: str) -> None:
        self.page.locator(self._DDL_ROOF_SHAPE_DOMINANT).select_option(value)

    def select_restricted_roof_shapes(self, value: str) -> None:
        self.page.locator(self._DDL_RESTRICTED_ROOF_SHAPES).select_option(value)

    def select_roof_material(self, value: str) -> None:
        self.page.locator(self._DDL_ROOF_MATERIAL_DOMINANT).select_option(value)

    def select_siding_type(self, value: str) -> None:
        self.page.locator(self._DDL_SIDING_TYPE_DOMINANT).select_option(value)

    # ------------------------------------------------------------------ #
    # Garage                                                               #
    # ------------------------------------------------------------------ #

    def select_garage_type(self, value: str) -> None:
        self.page.locator(self._DDL_GARAGE_TYPE).select_option(value)

    def select_garage_size(self, value: str) -> None:
        self.page.locator(self._DDL_GARAGE_SIZE).select_option(value)

    # ------------------------------------------------------------------ #
    # Protection / wind                                                    #
    # ------------------------------------------------------------------ #

    def select_protection_class(self, value: str) -> None:
        self.page.locator(self._DDL_PROTECTION_CLASS).select_option(value)

    def check_override_protection_class(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_PROTECTION_CLASS).check()

    def uncheck_override_protection_class(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_PROTECTION_CLASS).uncheck()

    def select_gable_bracing(self, value: str) -> None:
        self.page.locator(self._DDL_GABLE_BRACING).select_option(value)

    def select_hurricane_straps(self, value: str) -> None:
        self.page.locator(self._DDL_HURRICANE_STRAPS).select_option(value)

    def select_hurricane_shutters(self, value: str) -> None:
        self.page.locator(self._DDL_HURRICANE_SHUTTERS).select_option(value)

    def select_wind_mitigation(self, value: str) -> None:
        self.page.locator(self._DDL_WIND_MITIGATION).select_option(value)

    def fill_wind_mitigation_date(self, mm: str, dd: str, yyyy: str) -> None:
        self.page.get_by_role("textbox", name="mm").click()
        self.page.get_by_role("textbox", name="mm").fill(mm)
        self.page.get_by_role("textbox", name="dd").click()
        self.page.get_by_role("textbox", name="dd").fill(dd)
        self.page.get_by_role("textbox", name="yyyy").click()
        self.page.get_by_role("textbox", name="yyyy").fill(yyyy)

    # ------------------------------------------------------------------ #
    # Alarms & community                                                   #
    # ------------------------------------------------------------------ #

    def select_central_fire_alarm(self, value: str) -> None:
        self.page.locator(self._DDL_CENTRAL_FIRE_ALARM).select_option(value)

    def select_central_burglar_alarm(self, value: str) -> None:
        self.page.locator(self._DDL_CENTRAL_BURGLAR_ALARM).select_option(value)

    def select_gated_guarded_community(self, value: str) -> None:
        self.page.locator(self._DDL_GATED_GUARDED_COMMUNITY).select_option(value)

    def select_sprinklers(self, value: str) -> None:
        self.page.locator(self._DDL_SPRINKLERS).select_option(value)

    def select_swimming_pool(self, value: str) -> None:
        self.page.locator(self._DDL_SWIMMING_POOL).select_option(value)

    # ------------------------------------------------------------------ #
    # Safety features                                                      #
    # ------------------------------------------------------------------ #

    def select_smoke_detectors(self, value: str) -> None:
        self.page.locator(self._DDL_SMOKE_DETECTORS).select_option(value)

    def select_deadbolt_locks(self, value: str) -> None:
        self.page.locator(self._DDL_DEADBOLT_LOCKS).select_option(value)

    def select_fire_extinguisher(self, value: str) -> None:
        self.page.locator(self._DDL_FIRE_EXTINGUISHER).select_option(value)

    # ------------------------------------------------------------------ #
    # Distances                                                            #
    # ------------------------------------------------------------------ #

    def click_distance_to_fire_station(self) -> None:
        self.page.locator(self._TXT_DISTANCE_FIRE_STATION).click()

    def click_distance_to_fire_hydrant(self) -> None:
        self.page.locator(self._TXT_DISTANCE_FIRE_HYDRANT).click()

    # ------------------------------------------------------------------ #
    # Update years                                                         #
    # ------------------------------------------------------------------ #

    def click_wiring_year(self) -> None:
        self.page.locator(self._TXT_WIRING_YEAR).click()

    def click_plumbing_year(self) -> None:
        self.page.locator(self._TXT_PLUMBING_YEAR).click()

    def click_heating_year(self) -> None:
        self.page.locator(self._TXT_HEATING_YEAR).click()

    def click_roof_year(self) -> None:
        self.page.locator(self._TXT_ROOF_YEAR).click()

    def check_override_roof_year(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_ROOF_YEAR).check()

    def uncheck_override_roof_year(self) -> None:
        self.page.locator(self._CHK_OVERRIDE_ROOF_YEAR).uncheck()

    # ------------------------------------------------------------------ #
    # Renovation                                                           #
    # ------------------------------------------------------------------ #

    def select_will_renovate(self, value: str) -> None:
        self.page.locator(self._DDL_WILL_RENOVATE).select_option(value)

    def is_renovation_type_visible(self) -> bool:
        return self.page.locator(self._DDL_RENOVATION_TYPE).is_visible()

    def select_renovation_type(self, value: str) -> None:
        self.page.locator(self._DDL_RENOVATION_TYPE).select_option(value)

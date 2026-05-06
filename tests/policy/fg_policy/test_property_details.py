import os
import pytest
from playwright.sync_api import Page
from pages.login_page import LoginPage
from pages.fg_policy_overview_page import FGPolicyOverviewPage
from pages.fg_property_info_page import FGPropertyInfoPage

USERNAME   = "icortesmarquez@geovera"
POLICY_NUM = "GC60015733"

# ================================================================
# PROPERTY INFORMATION CHANGE LEGEND
# Edit values here to drive the test. The comment next to each
# constant shows the human-readable label for that option code.
# ================================================================

# -- Effective date --
CHANGE_REQUESTOR = "I"          # Insured

# -- Building / units --
UNIT_TYPE         = "3"         # Condominium
UNITS_IN_BUILDING = "B"         # 2 Family Units
UNITS_INSURED     = "2"         # 2 Units

# -- Construction --
CONSTRUCTION_TYPE = "MV"        # Masonry Veneer
DWELLING_TYPE     = "P"         # Primary
OCCUPANCY_TYPE    = "O"         # Owner Only
# NOTE: Vacancy Reason and Expected Occupancy Date are only visible
#       when Occupancy Type is Vacant or Unoccupied — conditional logic
#       in the test handles this automatically.
VACANCY_REASON    = "3"         # Sold but not occupied

# -- Foundation --
FOUNDATION_DOMINANT = "R"       # Crawlspace

# -- Structure --
LEVELS           = "3"          # 2 Levels
ROOF_SHAPE       = "3"          # Hip
RESTRICTED_ROOFS = "True"       # Yes
ROOF_MATERIAL    = "1"          # Asphalt Shingle / Fiberglass Shingle / Cement Fiber Shingle
SIDING_TYPE      = "VY"         # Vinyl

# -- Garage --
GARAGE_TYPE = "D"               # Detached Garage
GARAGE_SIZE = "1"               # 1 Car

# -- Protection --
PROTECTION_CLASS   = "5"        # 5
GABLE_BRACING      = "2"        # Braced
HURRICANE_STRAPS   = "False"    # No
HURRICANE_SHUTTERS = "2"        # Basic Protection

# -- Wind mitigation --
WIND_MITIGATION    = "True"     # Yes
WIND_MIT_DATE_MM   = "01"
WIND_MIT_DATE_DD   = "01"
WIND_MIT_DATE_YYYY = "2025"

# -- Alarms & community --
CENTRAL_FIRE_ALARM    = "True"  # Yes
CENTRAL_BURGLAR_ALARM = "True"  # Yes
GATED_COMMUNITY       = "True"  # Yes
SPRINKLERS            = "1"     # Partial sprinkler system
SWIMMING_POOL         = "IG"    # In-Ground Swimming Pool

# -- Safety features --
SMOKE_DETECTORS   = "True"      # Yes
DEADBOLT_LOCKS    = "True"      # Yes
FIRE_EXTINGUISHER = "True"      # Yes

# -- Renovation --
WILL_RENOVATE   = "True"        # Yes
RENOVATION_TYPE = "B"           # Minor


def _open_property_details(page: Page) -> tuple[Page, FGPropertyInfoPage]:
    LoginPage(page).login(USERNAME, os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("textbox", name="Quote/Binder/Policy #:").click()
    page.get_by_role("textbox", name="Quote/Binder/Policy #:").fill(POLICY_NUM)
    with page.expect_popup() as popup_info:
        page.get_by_role("button", name="Search").click()
    policy_page = popup_info.value
    if FGPolicyOverviewPage(policy_page).has_pending_changes():
        pytest.skip(
            f"Policy {POLICY_NUM} has pending changes — "
            "clear all pending transactions before running this test."
        )
    return policy_page, FGPropertyInfoPage(policy_page)


def test_property_information_edit(page: Page) -> None:
    policy_page, prop = _open_property_details(page)

    # Navigate to the Property Information edit form
    policy_page.get_by_role("link", name="Property Details").click()
    policy_page.get_by_role("link", name="Property Information").click()
    prop.click_edit()

    # ── Effective date ────────────────────────────────────────────────────
    prop.fill_effective_date_section(requestor=CHANGE_REQUESTOR)  # Insured
    prop.click_continue()

    # ── Building / units ─────────────────────────────────────────────────
    prop.select_unit_type(UNIT_TYPE)                   # Condominium
    prop.select_units_in_building(UNITS_IN_BUILDING)   # 2 Family Units
    prop.select_units_insured(UNITS_INSURED)           # 2 Units

    # ── Construction ─────────────────────────────────────────────────────
    prop.select_construction_type(CONSTRUCTION_TYPE)   # Masonry Veneer
    prop.select_dwelling_type(DWELLING_TYPE)           # Primary
    prop.select_occupancy_type(OCCUPANCY_TYPE)         # Owner Only
    if prop.is_vacancy_reason_visible():
        prop.select_vacancy_reason(VACANCY_REASON)     # Sold but not occupied
        prop.click_expected_occupancy_date()

    # ── Year built (click to confirm field is editable; override tested) ─
    prop.click_year_built()
    prop.check_override_year_built()
    prop.uncheck_override_year_built()

    # ── Building area (click to confirm field is editable; override tested) ─
    prop.click_living_area()
    prop.check_override_square_feet()
    prop.uncheck_override_square_feet()

    # ── Foundation ───────────────────────────────────────────────────────
    prop.select_foundation_dominant(FOUNDATION_DOMINANT)  # Crawlspace

    # ── Structure ────────────────────────────────────────────────────────
    prop.select_levels(LEVELS)                            # 2 Levels
    prop.select_roof_shape(ROOF_SHAPE)                    # Hip
    prop.select_restricted_roof_shapes(RESTRICTED_ROOFS)  # Yes
    prop.select_roof_material(ROOF_MATERIAL)              # Asphalt Shingle / Fiberglass / Cement Fiber
    prop.select_siding_type(SIDING_TYPE)                  # Vinyl

    # ── Garage ───────────────────────────────────────────────────────────
    prop.select_garage_type(GARAGE_TYPE)  # Detached Garage
    prop.select_garage_size(GARAGE_SIZE)  # 1 Car

    # ── Protection class ─────────────────────────────────────────────────
    prop.select_protection_class(PROTECTION_CLASS)  # 5
    prop.check_override_protection_class()
    prop.uncheck_override_protection_class()

    # ── Wind ─────────────────────────────────────────────────────────────
    prop.select_gable_bracing(GABLE_BRACING)          # Braced
    prop.select_hurricane_straps(HURRICANE_STRAPS)    # No
    prop.select_hurricane_shutters(HURRICANE_SHUTTERS)  # Basic Protection
    prop.select_wind_mitigation(WIND_MITIGATION)      # Yes
    prop.fill_wind_mitigation_date(WIND_MIT_DATE_MM, WIND_MIT_DATE_DD, WIND_MIT_DATE_YYYY)

    # ── Alarms & community ───────────────────────────────────────────────
    prop.select_central_fire_alarm(CENTRAL_FIRE_ALARM)          # Yes
    prop.select_central_burglar_alarm(CENTRAL_BURGLAR_ALARM)    # Yes
    prop.select_gated_guarded_community(GATED_COMMUNITY)        # Yes
    prop.select_sprinklers(SPRINKLERS)                          # Partial sprinkler system
    prop.select_swimming_pool(SWIMMING_POOL)                    # In-Ground Swimming Pool

    # ── Safety features ──────────────────────────────────────────────────
    prop.select_smoke_detectors(SMOKE_DETECTORS)      # Yes
    prop.select_deadbolt_locks(DEADBOLT_LOCKS)        # Yes
    prop.select_fire_extinguisher(FIRE_EXTINGUISHER)  # Yes

    # ── Distances (click to confirm fields are editable) ─────────────────
    prop.click_distance_to_fire_station()
    prop.click_distance_to_fire_hydrant()

    # ── Update years (click to confirm fields are editable) ──────────────
    prop.click_wiring_year()
    prop.click_plumbing_year()
    prop.click_heating_year()
    prop.click_roof_year()
    prop.check_override_roof_year()
    prop.uncheck_override_roof_year()

    # ── Renovation ───────────────────────────────────────────────────────
    prop.select_will_renovate(WILL_RENOVATE)       # Yes
    if prop.is_renovation_type_visible():
        prop.select_renovation_type(RENOVATION_TYPE)  # Minor
    prop.select_will_renovate("False")             # No — revert to avoid rating impact

    # ── Submit ────────────────────────────────────────────────────────────
    prop.click_continue()
    policy_page.get_by_text("Property Information Amended").click()
    policy_page.get_by_text("Yes").click()
    policy_page.get_by_text("Date on Wind Mitigation Verification Inspection: 01/01/").click()
    policy_page.get_by_text("Gable Bracing Braced changed").click()
    policy_page.get_by_role("button", name="Submit Changes").click()

    # ── Confirm submission ────────────────────────────────────────────────
    policy_page.get_by_role(
        "cell",
        name="Change requests have been submitted and are pending processing. "
             "Additional edits are not allowed on this policy until all pending "
             "transactions have been processed.",
        exact=True,
    ).click()
    policy_page.get_by_role("link", name=f"Policy {POLICY_NUM}").click()
    policy_page.get_by_role("table").filter(has_text="Pending Changes Requests").nth(3).click()

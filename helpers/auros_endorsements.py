"""Shared endorsement helpers for the Auros quote flow.

These functions operate on the Optional Endorsements page that follows the
General Information screen. They are intentionally framework-agnostic beyond
Playwright — no pytest dependencies — so both quote and bind tests can use them.
"""

import random


def fill_address_endorsement(page1, idx: str, num_addrs: int = 1, use_max: bool = False, use_ho2470: bool = False):
    """ctl00=Additional Interest, ctl01=Additional Insured, ctl06=Additional Residence Occupied.
    All three use ucEndorsement and have Name + SecondName fields.
    use_ho2470=True selects ddlNumberOfAddresses_HO_24_70 (max 4) instead of ddlNumberOfAddresses (max 2)."""
    prefix = f"ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_{idx}_ucEndorsement"

    page1.locator(f"#{prefix}_chkInclude").check()
    addr_ddl = f"{prefix}_ddlNumberOfAddresses_HO_24_70" if use_ho2470 else f"{prefix}_ddlNumberOfAddresses"
    page1.locator(f"#{addr_ddl}").select_option(str(num_addrs))

    for i in range(1, num_addrs + 1):
        name_i   = ("A" * 29 + str(i)) if use_max else f"ICMTEST{i}"
        second_i = ("B" * 29 + str(i)) if use_max else f"TESTER{i}"
        street_i = ("C" * 29 + str(i)) if use_max else f"123 TEST ST {i}"
        city_i   = ("D" * 29 + str(i)) if use_max else f"ORLANDO{i}"
        page1.locator(f"#{prefix}_txtName{i}").fill(name_i)
        page1.locator(f"#{prefix}_txtSecondName{i}").fill(second_i)
        page1.locator(f"#{prefix}_txtStreet{i}").fill(street_i)
        page1.locator(f"#{prefix}_txtCity{i}").fill(city_i)
        page1.locator(f"#{prefix}_ddlState{i}").select_option("FL")
        page1.locator(f"#{prefix}_txtZip{i}").fill(f"1234{i}")
        page1.locator(f"#{prefix}_txtZip4_{i}").fill(f"678{i}")


def fill_occupied_address_endorsement(page1, num_addrs: int = 1, use_max: bool = False, use_ho2470: bool = False):
    """ctl06=Additional Residence - Occupied by Insured. No Name/SecondName fields (only Street/City/State/Zip/Zip+4).
    Supports both standard ddlNumberOfAddresses (max 2) and HO-24-70 ddl (max 4) via use_ho2470 flag.
    Distinct from fill_address_endorsement, which assumes Name/SecondName exist."""
    prefix = "ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_ctl06_ucEndorsement"

    page1.locator(f"#{prefix}_chkInclude").check()
    addr_ddl = f"{prefix}_ddlNumberOfAddresses_HO_24_70" if use_ho2470 else f"{prefix}_ddlNumberOfAddresses"
    page1.locator(f"#{addr_ddl}").select_option(str(num_addrs))

    for i in range(1, num_addrs + 1):
        street_i = ("C" * 29 + str(i)) if use_max else f"123 TEST ST {i}"
        city_i   = ("D" * 29 + str(i)) if use_max else f"ORLANDO{i}"
        page1.locator(f"#{prefix}_txtStreet{i}").fill(street_i)
        page1.locator(f"#{prefix}_txtCity{i}").fill(city_i)
        page1.locator(f"#{prefix}_ddlState{i}").select_option("FL")
        page1.locator(f"#{prefix}_txtZip{i}").fill(f"1234{i}")
        page1.locator(f"#{prefix}_txtZip4_{i}").fill(f"678{i}")


def fill_rented_address_endorsement(page1, num_addrs: int = 1, use_max: bool = False):
    """ctl04=Additional Residence Rented to Others. No Name/SecondName fields.
    This endorsement only has HO-24-70 form; ddlNumberOfAddresses_HO_24_70 is always the visible dropdown."""
    prefix = "ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_ctl04_ucEndorsement"

    page1.locator(f"#{prefix}_chkInclude").scroll_into_view_if_needed()
    page1.locator(f"#{prefix}_chkInclude").click()
    page1.wait_for_timeout(1000)
    page1.locator(f"#{prefix}_ddlNumberOfAddresses_HO_24_70").select_option(str(num_addrs))

    for i in range(1, num_addrs + 1):
        street_i = ("C" * 29 + str(i)) if use_max else f"123 TEST ST {i}"
        city_i   = ("D" * 29 + str(i)) if use_max else f"ORLANDO{i}"
        page1.locator(f"#{prefix}_txtStreet{i}").fill(street_i)
        page1.locator(f"#{prefix}_txtCity{i}").fill(city_i)
        page1.locator(f"#{prefix}_ddlState{i}").select_option("FL")
        page1.locator(f"#{prefix}_txtZip{i}").fill(f"1234{i}")
        page1.locator(f"#{prefix}_txtZip4_{i}").fill(f"678{i}")


def check_only_endorsement(page1, idx: str):
    """ctl02=Ordinance or Law, ctl03=Windstorm Or Hail Exclusion. Checkbox only."""
    page1.locator(f"#ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_{idx}_ucEndorsement_chkInclude").check()


def fill_pool_endorsement(page1, limit: int = None):
    """ctl05=Increased Swimming Pool. Uses ucEndorsement2 (not ucEndorsement).
    limit options: 10000, 15000, 20000, 25000, 30000.
    If limit is None, randomly selects from available options.
    Returns the selected limit value."""
    prefix = "ctl00_ContentPlaceHolder1_ctl05_rptrEndorsements_ctl05_ucEndorsement2"
    page1.locator(f"#{prefix}_chkInclude").check()
    if limit is None:
        limit = random.choice([10000, 15000, 20000, 25000, 30000])
    page1.locator(f"#{prefix}_ddlLimit").select_option(str(limit))
    return limit


def fill_water_limitation(page1, limit: int = None):
    """Water Damage Limitation (ARIE-US-WL-20). Checks the checkbox if it is unchecked and enabled
    (mandatory/grayed-out checkboxes are left alone). Selects a limit from the HomeFlex dropdown.
    limit options: 10000, 20000, 25000.
    If limit is None, randomly selects from available options.
    Returns the selected limit value."""
    chk = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl00_ucEndorsement2a_rptrProductEndorsements_ctl00_chkInclude")
    if chk.count() > 0 and chk.is_visible() and chk.is_enabled() and not chk.is_checked():
        chk.check()
    if limit is None:
        limit = random.choice([10000, 20000, 25000])
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl00_ucEndorsement2a_ddlLimit").select_option(str(limit))
    return limit


def click_water_update_if_visible(page1):
    """Clicks the water endorsement Update button if it is visible.
    Must be called after all water section interactions — checking any water checkbox
    from an unchecked state causes this button to appear."""
    btn = page1.locator("#ctl00_ContentPlaceHolder1_ctl05_btnUpdateWaterLimitEndorsements")
    if btn.is_visible():
        btn.click()
        page1.wait_for_load_state("networkidle")


def fill_water_excl_named_peril(page1):
    """Water Damage Exclusion With Named Peril Coverage (US-A-NP). Only ctl00 — ctl01 maps to US-A and conflicts."""
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl01_ucEndorsement1a_rptrProductEndorsements_ctl00_chkInclude").check()


def fill_water_excl_limited(page1):
    """Water Damage Exclusion And Limited Named Peril Coverage (US-A). Only ctl00 — avoids conflict with US-A-NP."""
    page1.locator("#ctl00_ContentPlaceHolder1_ctl05_rptrWaterEndorsements_ctl02_ucEndorsement1a_rptrProductEndorsements_ctl00_chkInclude").check()

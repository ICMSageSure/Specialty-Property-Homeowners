import re
import os
from playwright.sync_api import Page, expect

USERNAME = "icortesmarquez@geovera"

def test_example(page: Page) -> None:
    page.goto("https://tst.mygeosource.com/")
    page.get_by_role("textbox", name="Policyholder Email or Agent").click()
    page.get_by_role("textbox", name="Policyholder Email or Agent").fill(USERNAME)
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill(os.environ["INTERNAL_PASSWORD"])
    page.get_by_role("button", name="Log In", exact=True).click()
    page.get_by_role("button", name="Policies").click()
    page.get_by_text("Select Product").click()
    page.get_by_role("tree").get_by_text("Homeowners").click()
    page.get_by_text("Select State...").first.click()
    page.get_by_role("option", name="FL").click()
    page.get_by_text("Policy #").first.click()
    page.get_by_role("option", name="Policy #").click()
    page.get_by_role("textbox", name="Search").click()
    page.get_by_role("textbox", name="Search").fill("GC50018717")
    page.get_by_role("button", name="View").click()
    with page.expect_popup() as page1_info:
        page.get_by_role("gridcell", name="GC50018717").click()
    page1 = page1_info.value
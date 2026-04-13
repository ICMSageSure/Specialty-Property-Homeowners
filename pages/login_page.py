import os
from playwright.sync_api import Page


class LoginPage:
    URL = "https://tst.mygeosource.com/"

    def __init__(self, page: Page):
        self.page = page

    def login(self, username: str, password: str):
        self.page.goto(self.URL)
        self.page.get_by_role("textbox", name="Policyholder Email or Agent").fill(username)
        self.page.get_by_role("textbox", name="Password").fill(password)
        self.page.get_by_role("button", name="Log In", exact=True).click()

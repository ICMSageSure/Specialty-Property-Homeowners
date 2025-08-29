import pytest
from playwright.sync_api import Page, expect

@pytest.fixture(scope="session")
def browser_context_args(browser_context_args, playwright):
    return {"viewport": {"width":1920,"height":1080}}

@pytest.fixture(scope="session")
def setUp(page):
    login = Login(page)
    login.navigate_to_login()
    assert login.url, "BASE_URL is not set"
    yield
    page.close()
    


@pytest.fixture
def login(setUp, page):
    login = Login(page)
    login.enter_username(login.username)
    login.enter_password(login.password)
    return page
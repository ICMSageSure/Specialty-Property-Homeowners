import pytest
import shutil
from pathlib import Path
from datetime import datetime
from playwright.sync_api import Page, expect

SCREENSHOTS_ROOT = Path(__file__).parent / "screenshots"
MAX_RUNS = 3

@pytest.fixture(scope="session")
def browser_context_args(browser_context_args, playwright):
    return {"viewport": {"width":1920,"height":1080}}

@pytest.fixture(scope="session")
def run_dir():
    SCREENSHOTS_ROOT.mkdir(exist_ok=True)
    run_path = SCREENSHOTS_ROOT / datetime.now().strftime("run_%Y-%m-%d_%H-%M-%S")
    run_path.mkdir()
    runs = sorted(SCREENSHOTS_ROOT.iterdir(), key=lambda p: p.stat().st_mtime)
    for old in runs[:-MAX_RUNS]:
        shutil.rmtree(old, ignore_errors=True)
    return run_path

@pytest.fixture
def snap(run_dir, request):
    test_name = request.node.name
    counter = [0]
    def take(page, label):
        counter[0] += 1
        dest = run_dir / test_name / f"{counter[0]:02d}_{label}.png"
        dest.parent.mkdir(exist_ok=True)
        page.screenshot(path=str(dest))
    return take

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
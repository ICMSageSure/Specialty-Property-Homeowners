import base64
import pytest
import shutil
from pathlib import Path
from datetime import datetime
from playwright.sync_api import Page, expect

SCREENSHOTS_ROOT = Path(__file__).parent / "screenshots"
MAX_RUNS = 3

# Stores screenshot paths per test node ID so the HTML report hook can attach them
_test_screenshots: dict[str, list[Path]] = {}


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
    node_id = request.node.nodeid
    counter = [0]
    paths: list[Path] = []
    _test_screenshots[node_id] = paths

    def take(page, label):
        counter[0] += 1
        dest = run_dir / test_name / f"{counter[0]:02d}_{label}.png"
        dest.parent.mkdir(exist_ok=True)
        page.screenshot(path=str(dest))
        paths.append(dest)

    return take


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    report = outcome.get_result()

    if report.when != "call":
        return

    from pytest_html import extras
    extra = getattr(report, "extras", [])

    for path in _test_screenshots.get(item.nodeid, []):
        if path.exists():
            img_b64 = base64.b64encode(path.read_bytes()).decode()
            extra.append(extras.image(img_b64, mime_type="image/png", name=path.stem))

    report.extras = extra

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
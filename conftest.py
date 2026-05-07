import base64
import os
import pytest
import shutil
from pathlib import Path
from datetime import datetime
from pages.login_page import LoginPage

SCREENSHOTS_ROOT = Path(__file__).parent / "screenshots"
MAX_RUNS = 3
_INTERNAL_USERS = frozenset({"icortesmarquez@geovera", "tsicortesmarquez@geovera"})

_test_screenshots: dict[str, list[Path]] = {}


@pytest.fixture(scope="session")
def browser_context_args(browser_context_args):
    return {"viewport": {"width": 1920, "height": 1080}}


@pytest.fixture(scope="session")
def run_dir():
    SCREENSHOTS_ROOT.mkdir(exist_ok=True)
    run_path = SCREENSHOTS_ROOT / datetime.now().strftime("run_%Y-%m-%d_%H-%M-%S")
    run_path.mkdir(exist_ok=True)
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
        safe_name = test_name.replace("/", "_").replace("\\", "_")
        dest = run_dir / safe_name / f"{counter[0]:02d}_{label}.png"
        dest.parent.mkdir(exist_ok=True)
        page.screenshot(path=str(dest), full_page=True)
        paths.append(dest)

    return take


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    outcome = yield
    report = outcome.get_result()
    setattr(item, f"rep_{call.when}", report)

    if report.when != "call":
        return

    from pytest_html import extras
    extra = getattr(report, "extras", [])

    for path in _test_screenshots.get(item.nodeid, []):
        if path.exists():
            img_b64 = base64.b64encode(path.read_bytes()).decode()
            extra.append(extras.image(img_b64, mime_type="image/png", name=path.stem))

    report.extras = extra


@pytest.fixture
def failure_pages():
    """Register popup/extra pages here so they get captured on failure."""
    return []


@pytest.fixture(autouse=True)
def screenshot_on_failure(page, request, snap, run_dir, failure_pages):
    yield
    rep = getattr(request.node, "rep_call", None)
    if rep and rep.failed:
        try:
            snap(page, "FAILED")
        except Exception:
            pass
        for i, extra_page in enumerate(failure_pages, start=1):
            try:
                snap(extra_page, f"FAILED_popup_{i}")
            except Exception:
                pass
            try:
                url = extra_page.url
                title = extra_page.title()
                body_text = extra_page.locator("body").inner_text(timeout=5000)
                safe_name = request.node.name.replace("/", "_").replace("\\", "_")
                log_path = run_dir / safe_name / "failure_log.txt"
                log_path.parent.mkdir(exist_ok=True)
                log_path.write_text(
                    f"URL: {url}\nTitle: {title}\n\n{body_text}",
                    encoding="utf-8",
                )
                print(f"\n=== FAILURE PAGE TEXT ({url}) ===\n{body_text[:2000]}\n===")
            except Exception:
                pass


@pytest.fixture
def login(page):
    def _login(username: str, password: str | None = None) -> None:
        if password is None:
            if username not in _INTERNAL_USERS:
                raise ValueError(
                    f"No password provided for '{username}'. "
                    f"INTERNAL_PASSWORD only works for internal users. "
                    f"Pass a password explicitly: login('{username}', 'your_password')"
                )
            password = os.environ["INTERNAL_PASSWORD"]
        LoginPage(page).login(username, password)
    return _login

# P4 — FG/HL Parameterization Scoping Notes

Status: not started. Scoping written 2026-05-08.

## Executive summary

The FG and HL test trees are not 92% identical as the original audit claimed. After actual diff analysis:

- **Quote pairs: 95–97% similar.** Mergeable with carrier-config parameterization.
- **Bind pairs: ~82% similar.** Materially more divergence; merging is harder and may not pay off.

Recommendation: do **quote-pair merge first** as a phase 1. Defer bind-pair merge until quote-pair merge is proven stable in CI.

## Pairs in scope

10 file pairs (FG vs HL), all under `tests/quotes/`:

| Pair | FG file | HL file | Diff similarity |
|---|---|---|---|
| ho3 quote | `test_quote_fg_ho3.py` (156L) | `test_quote_hl_ho3.py` (153L) | 95.1% |
| ho3t quote | `test_quote_fg_ho3t.py` | `test_quote_hl_ho3t.py` | ~95% (assumed) |
| ho_wind quote | `test_quote_fg_ho_wind.py` (157L) | `test_quote_hl_ho_wind.py` (153L) | 96.7% |
| vh1 quote | `test_quote_fg_vh1.py` (180L) | `test_quote_hl_vh1.py` (181L) | 95.2% |
| vh2 quote | `test_quote_fg_vh2.py` | `test_quote_hl_vh2.py` | ~95% (assumed) |
| ho3 bind | `test_bind_fg_ho3.py` (233L) | `test_bind_hl_ho3.py` (211L) | 82.4% |
| ho3t bind | `test_bind_fg_ho3t.py` | `test_bind_hl_ho3t.py` | ~82% (assumed) |
| ho_wind bind | `test_bind_fg_ho_wind.py` | `test_bind_hl_ho_wind.py` | ~82% (assumed) |
| vh1 bind | `test_bind_fg_vh1.py` | `test_bind_hl_vh1.py` | ~82% (assumed) |
| vh2 bind | `test_bind_fg_vh2.py` | `test_bind_hl_vh2.py` | ~82% (assumed) |

## Carrier-specific values inventory

These are the parameterizable values that differ between FG and HL in the quote pairs (concrete values pulled from `test_quote_fg_ho3.py` vs `test_quote_hl_ho3.py`):

| Concept | FG value | HL value | Notes |
|---|---|---|---|
| `RESULTS_FILE` path | `results/fg/fg_test_results.xlsx` | `results/hl/hl_test_results.xlsx` | Different output |
| `username` | `icortesmarquez@geovera.com` | `tsicortesmarquez@geovera` | Different login user |
| Wholesaler dropdown value | `APPUND` | `HULL001` | Carrier-specific |
| Producer Number | `APPAGT` | `10506` | Carrier-specific |
| User-list grid row selector | `_ctl02_btnId` | `_ctl28_btnId` | Position differs in user dropdown |
| Welcome wait timeout | 3000ms / 2000ms | 1000ms | Timing — FG is slower |
| Quote-results page layout | Multi-carrier panel: `div[id$='_pnlHOQuotes']` filter by "GeoVera Specialty Insurance Company" → scoped selectors | Single-carrier flat: `#ctl00_ContentPlaceHolder1_ucQuoteDetailHOFlood_*` | **Structural divergence — biggest risk** |
| Scenario label written to xlsx | `"FG HO-3"` | `"HL HO-3"` | Cosmetic |

For bind pairs there are additional parameterizable values not yet inventoried (binder selectors may differ, mortgagee/AP block IDs may differ).

## Phase 1 design — quote pairs only

### Approach

Define a `Carrier` config object (dataclass or dict) that holds the 8 differing values. Single shared test function takes `carrier` as a `pytest.parametrize` axis.

```python
# helpers/carrier_config.py (new file)
from dataclasses import dataclass

@dataclass(frozen=True)
class Carrier:
    name: str                    # "fg" or "hl"
    username: str
    results_file: str
    wholesaler: str
    producer_number: str
    user_grid_row_selector: str  # the _ctl02_ or _ctl28_ selector
    welcome_wait_ms: int
    uses_panel_scoped_quote_results: bool

FG = Carrier(
    name="fg",
    username="icortesmarquez@geovera.com",
    results_file=".../results/fg/fg_test_results.xlsx",
    wholesaler="APPUND",
    producer_number="APPAGT",
    user_grid_row_selector="#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl02_btnId",
    welcome_wait_ms=3000,
    uses_panel_scoped_quote_results=True,
)

HL = Carrier(
    name="hl",
    username="tsicortesmarquez@geovera",
    results_file=".../results/hl/hl_test_results.xlsx",
    wholesaler="HULL001",
    producer_number="10506",
    user_grid_row_selector="#ctl00_WelcomeUserControl_ucWelcomeUserLookup_dgUserList_ctl28_btnId",
    welcome_wait_ms=1000,
    uses_panel_scoped_quote_results=False,
)
```

### Test function shape

```python
# tests/quotes/test_quote_ho3.py (new merged file)
import pytest
from helpers.carrier_config import FG, HL

@pytest.mark.parametrize("carrier", [FG, HL], ids=lambda c: c.name)
def test_quote_ho3(page, snap, carrier):
    # Body uses carrier.* throughout instead of hardcoded values
    ...
```

### Where the merge gets ugly

The `uses_panel_scoped_quote_results` flag is a **branching** difference, not a value substitution. The two carriers have different DOM layouts on the quote-results page:

- FG: multi-carrier results in a collapsible panel; you find the panel filtered by carrier name, then scope selectors inside.
- HL: single-carrier flat layout; selectors target the page directly.

This requires either:
- A `get_quote_results(page, carrier)` helper that branches on the flag, OR
- Two separate functions called via the carrier config

Either way, this is the part where the merge stops being a clean parameterization and starts looking like an abstraction. Worth keeping the helper small and well-named.

## Phase 2 design — bind pairs

The 82% similarity reflects ~17 lines of additional divergence in each bind file beyond the quote-test divergences. Pre-merging requires a separate diff inventory. Likely additional parameterizable values:
- Binder Effective Date selectors (probably identical, verify)
- Bill To dropdown values (probably identical)
- Mortgagee block prefixes (likely identical — they're framework-driven)
- Pay plan / autopay handling (likely identical)
- Result-grid scraping selectors (may differ — same risk as quote-results layout)

Proceed only after Phase 1 ships and is stable.

## File deletion plan (after merge succeeds)

After merge:
- Delete: `test_quote_fg_*.py` and `test_quote_hl_*.py` (10 files for quote phase)
- Add: `test_quote_ho3.py`, `test_quote_ho3t.py`, `test_quote_ho_wind.py`, `test_quote_vh1.py`, `test_quote_vh2.py` (5 files)
- Net: 10 → 5 files for quote tests
- pytest collects each parametrized once per carrier → same test count as before, just centralized code

For bind phase: same 10 → 5 reduction. Total 20 → 10 files when all phases complete.

## Risk inventory

| Risk | Severity | Mitigation |
|---|---|---|
| Quote-results layout divergence (panel vs flat) creates branching abstraction | Medium | Keep branch helper small; document why |
| `welcome_wait_ms` difference suggests other timing fragility hiding in FG | Low | Run merged tests against both carriers for a week before deleting originals |
| Test ID changes break CI selectors (e.g., `pytest -k "test_quote_fg_ho3"` won't match anymore) | Low | Update batch scripts after merge |
| Merge introduces typos in one of the 8 config values for one carrier | Medium | Run merged tests for both carriers before merging branch |
| Subtle behavioral differences not caught by 95%/82% similarity (e.g., a single `if` check that exists in HL but not FG) | Medium-High | **Manual line-by-line diff review of each pair before merging** |

## Realistic effort estimate

Plan finding #7 said "3-5 hrs (phased)". Revised based on actual diff analysis:

| Task | Hours |
|---|---|
| Carrier config dataclass + constants | 1 |
| Merge quote pairs (5 pairs × ~30 min each + manual diff review) | 3–4 |
| Run + debug merged quote tests against both carriers | 2–3 |
| Delete old quote files; update scripts/CI | 1 |
| **Phase 1 subtotal** | **7–9 hours** |
| Bind-pair diff inventory | 2 |
| Merge bind pairs | 4–6 |
| Run + debug merged bind tests | 3–4 |
| Delete old bind files; update scripts/CI | 1 |
| **Phase 2 subtotal** | **10–13 hours** |
| **Total** | **17–22 hours** |

Spread across sessions, this is a 3–5 day project for a focused contributor.

## Recommendation

1. **Don't start Phase 2 until Phase 1 is stable in CI for ≥1 week.**
2. **Don't run the bulk merge as a single PR.** Do one pair at a time, ship each, monitor.
3. **Before each pair merge, run `python -c "import difflib; ..."` to dump the unified diff** and read it line-by-line. Some 5% differences are timing-only; some are silent semantic changes.
4. **Keep `helpers/carrier_config.py` minimal.** If it grows past ~30 lines per carrier, the abstraction is leaking and it's a sign to back out.

## Pre-flight checklist before starting Phase 1

- [ ] Confirm both `INTERNAL_PASSWORD` envs work for both `icortesmarquez@geovera.com` and `tsicortesmarquez@geovera`
- [ ] Confirm both carriers' test envs are reachable (FG and HL share `tst.mygeosource.com` but use different wholesaler accounts)
- [ ] Verify `_ctl02_btnId` and `_ctl28_btnId` user-grid selectors are still valid (these depend on user list size — fragile)
- [ ] Decide: do `username`s belong in carrier config or as env vars (`AUROS_USERNAME` precedent exists in `test_bind_auros.py`)?
- [ ] Identify whether `bad_addresses.txt` should be per-carrier or shared

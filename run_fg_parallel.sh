#!/usr/bin/env bash
# Runs fg quote tests in two parallel streams, one per login.
# Sessions within each stream are sequential — no overlapping logins.

set -euo pipefail

REPO="$(cd "$(dirname "$0")" && pwd)"
cd "$REPO"

# icortesmarquez@geovera.com stream (6 files)
USER1_TESTS=(
  tests/quotes/fg_quote/test_quote_fg_ho3.py
  tests/quotes/fg_quote/test_quote_fg_ho3t.py
  tests/quotes/fg_quote/test_quote_fg_ho_wind.py
  tests/quotes/fg_quote/test_bind_fg_ho3.py
  tests/quotes/fg_quote/test_bind_fg_ho3t.py
  tests/quotes/fg_quote/test_v2fg_ho_wind_.py
)

# tsicortesmarquez@geovera.com stream (5 files)
USER2_TESTS=(
  tests/quotes/fg_quote/test_quote_fg_vh1.py
  tests/quotes/fg_quote/test_quote_fg_vh2.py
  tests/quotes/fg_quote/test_bind_fg_ho_wind.py
  tests/quotes/fg_quote/test_bind_fg_vh1.py
  tests/quotes/fg_quote/test_bind_fg_vh2.py
)

echo "[$(date '+%H:%M:%S')] Starting both streams..."

pytest "${USER1_TESTS[@]}" \
  --html=report_fg_user1.html --self-contained-html \
  -v 2>&1 | sed 's/^/[user1] /' &
PID1=$!

pytest "${USER2_TESTS[@]}" \
  --html=report_fg_user2.html --self-contained-html \
  -v 2>&1 | sed 's/^/[user2] /' &
PID2=$!

wait $PID1; STATUS1=$?
wait $PID2; STATUS2=$?

echo ""
echo "[$(date '+%H:%M:%S')] Both streams finished."
echo "  user1 (icortesmarquez): exit $STATUS1  → report_fg_user1.html"
echo "  user2 (tsicortesmarquez): exit $STATUS2 → report_fg_user2.html"

[ $STATUS1 -eq 0 ] && [ $STATUS2 -eq 0 ]

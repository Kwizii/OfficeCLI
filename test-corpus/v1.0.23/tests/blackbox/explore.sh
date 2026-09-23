#!/usr/bin/env bash
# Black-box exploratory tests for officecli
# Covers new bugs and boundary conditions discovered during exploration
# Usage: ./explore.sh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CLI="${CLI:-$SCRIPT_DIR/../../src/officecli/bin/Debug/net10.0/osx-arm64/officecli}"

PASS=0
FAIL=0
TMPDIR=$(mktemp -d)
trap 'rm -rf "$TMPDIR"' EXIT

pass() { echo "PASS: $1"; PASS=$((PASS + 1)); }
fail() { echo "FAIL: $1 — $2"; FAIL=$((FAIL + 1)); }

run() {
    _out=$("$@" 2>&1); _code=$?
}

check_exit() {
    local label="$1" expected="$2" actual="$3"
    if [ "$actual" -eq "$expected" ]; then
        pass "$label (exit $expected)"
    else
        fail "$label" "expected exit $expected, got $actual"
    fi
}

check_output() {
    local label="$1" pattern="$2" output="$3"
    if echo "$output" | grep -q "$pattern"; then
        pass "$label"
    else
        fail "$label" "output missing '$pattern'"
    fi
}

check_no_output() {
    local label="$1" pattern="$2" output="$3"
    if echo "$output" | grep -q "$pattern"; then
        fail "$label" "output should NOT contain '$pattern'"
    else
        pass "$label"
    fi
}

echo "=== officecli exploratory tests ==="
echo "CLI: $CLI"
echo ""

# ---------------------------------------------------------------------------
# BT-NEW-1: XLSX chart should appear in get DOM (currently BROKEN)
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/new1.xlsx"
run "$CLI" add "$TMPDIR/new1.xlsx" "/Sheet1" --type chart \
    --prop "type=bar" --prop "data=Sales:100,200,300"
run "$CLI" query "$TMPDIR/new1.xlsx" "chart"
check_output "BT-NEW-1 xlsx chart visible via query" "Matches: 1" "$_out"
run "$CLI" get "$TMPDIR/new1.xlsx" "/Sheet1"
if echo "$_out" | grep -q "chart"; then
    pass "BT-NEW-1 xlsx chart visible via get (FIXED)"
else
    fail "BT-NEW-1 xlsx chart visible via get" "REGRESSION: chart not shown in get /Sheet1 tree"
fi

# ---------------------------------------------------------------------------
# BT-NEW-2: XLSX multiple cells in same row auto-advance column (currently BROKEN)
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/new2.xlsx"
run "$CLI" add "$TMPDIR/new2.xlsx" "/Sheet1" --type row
run "$CLI" add "$TMPDIR/new2.xlsx" "/Sheet1/row[1]" --type cell --prop "value=A"
run "$CLI" add "$TMPDIR/new2.xlsx" "/Sheet1/row[1]" --type cell --prop "value=B"
run "$CLI" add "$TMPDIR/new2.xlsx" "/Sheet1/row[1]" --type cell --prop "value=C"
run "$CLI" get "$TMPDIR/new2.xlsx" "/Sheet1/row[1]"
if echo "$_out" | grep -q "B1" && echo "$_out" | grep -q "C1"; then
    pass "BT-NEW-2 xlsx cells auto-advance column (FIXED)"
else
    fail "BT-NEW-2 xlsx cells auto-advance column" "REGRESSION: expected A1=A B1=B C1=C, all went to A1"
fi

# Workaround with explicit ref= should always work
run "$CLI" create "$TMPDIR/new2b.xlsx"
run "$CLI" add "$TMPDIR/new2b.xlsx" "/Sheet1" --type row
run "$CLI" add "$TMPDIR/new2b.xlsx" "/Sheet1/row[1]" --type cell --prop "value=A" --prop "ref=A1"
run "$CLI" add "$TMPDIR/new2b.xlsx" "/Sheet1/row[1]" --type cell --prop "value=B" --prop "ref=B1"
run "$CLI" add "$TMPDIR/new2b.xlsx" "/Sheet1/row[1]" --type cell --prop "value=C" --prop "ref=C1"
run "$CLI" get "$TMPDIR/new2b.xlsx" "/Sheet1/row[1]"
check_output "BT-NEW-2 explicit ref=B1 workaround works" "B1" "$_out"

# ---------------------------------------------------------------------------
# BT-17: XLSX add row --index beyond max (1048576) should be rejected
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt17.xlsx"
run "$CLI" add "$TMPDIR/bt17.xlsx" "/Sheet1" --type row --index 1048577
if [ $_code -eq 1 ] && echo "$_out" | grep -qi "range\|exceed\|invalid\|max"; then
    pass "BT-17 add row beyond max range rejected (exit 1)"
else
    fail "BT-17 add row beyond max range" "expected exit 1 with range error, got exit $_code: $_out"
fi

# Row at max boundary should succeed
run "$CLI" create "$TMPDIR/bt17b.xlsx"
run "$CLI" add "$TMPDIR/bt17b.xlsx" "/Sheet1" --type row --index 1048576
check_exit "BT-17 add row at max index 1048576 accepted" 0 $_code

# ---------------------------------------------------------------------------
# BT-5 regression: PPTX query text= filter
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt5reg.pptx"
run "$CLI" add "$TMPDIR/bt5reg.pptx" "/" --type slide
run "$CLI" add "$TMPDIR/bt5reg.pptx" "/slide[1]" --type shape --prop "text=Hello"
run "$CLI" add "$TMPDIR/bt5reg.pptx" "/slide[1]" --type shape --prop "text=World"

run "$CLI" query "$TMPDIR/bt5reg.pptx" "shape[text=Hello]"
if echo "$_out" | grep -q "Matches: 1"; then
    pass "BT-5 query shape[text=Hello] exact match (FIXED)"
else
    fail "BT-5 query shape[text=Hello]" "REGRESSION: expected Matches: 1, got: $(echo "$_out" | grep Matches)"
fi

run "$CLI" query "$TMPDIR/bt5reg.pptx" "shape[text~=NOTHING]"
if echo "$_out" | grep -q "Matches: 0"; then
    pass "BT-5 query shape[text~=NOTHING] returns 0 (FIXED)"
else
    fail "BT-5 query shape[text~=NOTHING]" "REGRESSION: ~= operator broken, got: $(echo "$_out" | grep Matches)"
fi

run "$CLI" query "$TMPDIR/bt5reg.pptx" "shape[text~=Hello]"
if echo "$_out" | grep -q "Matches: 1"; then
    pass "BT-5 query shape[text~=Hello] contains 1 match (FIXED)"
else
    fail "BT-5 query shape[text~=Hello]" "expected Matches: 1, got: $(echo "$_out" | grep Matches)"
fi

# ---------------------------------------------------------------------------
# Boundary: PPTX shape at zero size validates
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/zero.pptx"
run "$CLI" add "$TMPDIR/zero.pptx" "/" --type slide
run "$CLI" add "$TMPDIR/zero.pptx" "/slide[1]" --type shape \
    --prop "text=Zero" --prop "width=0" --prop "height=0"
run "$CLI" validate "$TMPDIR/zero.pptx"
check_output "boundary: PPTX zero-size shape validates" "Validation passed" "$_out"

# ---------------------------------------------------------------------------
# Boundary: XLSX cell at row 0 is rejected
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/row0.xlsx"
run "$CLI" add "$TMPDIR/row0.xlsx" "/Sheet1/row[1]" --type cell \
    --prop "value=X" --prop "ref=A0"
check_exit   "boundary: XLSX cell ref A0 (row 0) rejected" 1 $_code
check_output "boundary: XLSX cell ref A0 error message" "out of range" "$_out"

# ---------------------------------------------------------------------------
# Boundary: XLSX set beyond max row is rejected
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/maxrow.xlsx"
run "$CLI" set "$TMPDIR/maxrow.xlsx" "/Sheet1/A1048577" --prop "value=X"
check_exit "boundary: XLSX set beyond max row rejected" 1 $_code

# ---------------------------------------------------------------------------
# Feature: XLSX numFmt works
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/numfmt.xlsx"
run "$CLI" add "$TMPDIR/numfmt.xlsx" "/Sheet1" --type row
run "$CLI" add "$TMPDIR/numfmt.xlsx" "/Sheet1/row[1]" --type cell \
    --prop "value=1234.56" --prop "ref=A1"
run "$CLI" set "$TMPDIR/numfmt.xlsx" "/Sheet1/A1" --prop "numFmt=#,##0.00"
run "$CLI" get "$TMPDIR/numfmt.xlsx" "/Sheet1/A1"
check_output "feature: XLSX numFmt applied" "numberformat" "$_out"

# ---------------------------------------------------------------------------
# Feature: PPTX slide copy (--from)
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/copy.pptx"
run "$CLI" add "$TMPDIR/copy.pptx" "/" --type slide --prop "title=Original"
run "$CLI" add "$TMPDIR/copy.pptx" "/" --type slide --from "/slide[1]"
run "$CLI" get "$TMPDIR/copy.pptx" "/"
check_output "feature: PPTX slide copy --from creates slide[2]" "/slide\[2\]" "$_out"
run "$CLI" validate "$TMPDIR/copy.pptx"
check_output "feature: PPTX slide copy validates" "Validation passed" "$_out"

# ---------------------------------------------------------------------------
# Feature: XLSX merge cells validates
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/merge.xlsx"
run "$CLI" add "$TMPDIR/merge.xlsx" "/Sheet1" --type row
run "$CLI" add "$TMPDIR/merge.xlsx" "/Sheet1/row[1]" --type cell --prop "value=Merged" --prop "ref=A1"
run "$CLI" add "$TMPDIR/merge.xlsx" "/Sheet1/row[1]" --type cell --prop "value=X" --prop "ref=B1"
run "$CLI" set "$TMPDIR/merge.xlsx" "/Sheet1/A1:B1" --prop "merge=true"
run "$CLI" validate "$TMPDIR/merge.xlsx"
check_output "feature: XLSX merge cells validates" "Validation passed" "$_out"

# ---------------------------------------------------------------------------
# Feature: DOCX complex formatting (bold + italic + underline)
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/fmt.docx"
run "$CLI" add "$TMPDIR/fmt.docx" "/body" --type paragraph \
    --prop "text=Formatted" --prop "bold=true" \
    --prop "italic=true" --prop "underline=single"
run "$CLI" get "$TMPDIR/fmt.docx" "/body/p[1]"
check_output "feature: DOCX bold" "bold: True" "$_out"
check_output "feature: DOCX italic" "italic: True" "$_out"
check_output "feature: DOCX underline single" "underline: single" "$_out"

# ---------------------------------------------------------------------------
# Feature: DOCX heading styles
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/heading.docx"
run "$CLI" add "$TMPDIR/heading.docx" "/body" --type paragraph \
    --prop "text=H1" --prop "style=Heading1"
run "$CLI" add "$TMPDIR/heading.docx" "/body" --type paragraph \
    --prop "text=H2" --prop "style=Heading2"
run "$CLI" get "$TMPDIR/heading.docx" "/body"
check_output "feature: DOCX Heading1 style" "Heading1" "$_out"
check_output "feature: DOCX Heading2 style" "Heading2" "$_out"

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
echo ""
echo "=== Results ==="
echo "PASS: $PASS"
echo "FAIL: $FAIL"
echo "TOTAL: $((PASS + FAIL))"
[ "$FAIL" -eq 0 ]

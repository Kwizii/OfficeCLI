#!/usr/bin/env bash
# Black-box regression tests for officecli
# Covers all previously fixed bugs: BT-1 through BT-21
# Usage: ./regression.sh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CLI="${CLI:-$SCRIPT_DIR/../../src/officecli/bin/Debug/net10.0/osx-arm64/officecli}"

PASS=0
FAIL=0
TMPDIR=$(mktemp -d)
trap 'rm -rf "$TMPDIR"' EXIT

pass() { echo "PASS: $1"; PASS=$((PASS + 1)); }
fail() { echo "FAIL: $1 — $2"; FAIL=$((FAIL + 1)); }

run() {
    # Run command, capture stdout+stderr and exit code safely
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

echo "=== officecli regression tests ==="
echo "CLI: $CLI"
echo ""

# ---------------------------------------------------------------------------
# BT-1: create with unsupported file type gives clear error + exit 1
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/test.txt"
check_output "BT-1 error message" "Unsupported file type" "$_out"
check_exit   "BT-1 exit code" 1 $_code

# ---------------------------------------------------------------------------
# BT-2: invalid command exits 1
# ---------------------------------------------------------------------------
run "$CLI" invalidcmd
check_exit "BT-2 invalid command exit code" 1 $_code

# ---------------------------------------------------------------------------
# BT-2-ext: batch with failures exits 1
# ---------------------------------------------------------------------------
cat > "$TMPDIR/batch_fail.json" << 'EOF'
[
  {"command": "add", "parent": "/", "type": "slide"},
  {"command": "add", "parent": "/nonexistent", "type": "shape", "props": {"text": "Error"}}
]
EOF
run "$CLI" create "$TMPDIR/bt2ext.pptx"
run "$CLI" batch "$TMPDIR/bt2ext.pptx" --input "$TMPDIR/batch_fail.json"
check_exit   "BT-2-ext batch failure exit code" 1 $_code
check_output "BT-2-ext batch error summary" "failed" "$_out"

# ---------------------------------------------------------------------------
# BT-3: XLSX formula does not display double equals
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt3.xlsx"
run "$CLI" add "$TMPDIR/bt3.xlsx" "/Sheet1" --type row
run "$CLI" add "$TMPDIR/bt3.xlsx" "/Sheet1/row[1]" --type cell --prop "formula==SUM(A1:A5)" --prop "ref=A1"
run "$CLI" get "$TMPDIR/bt3.xlsx" "/Sheet1/A1"
check_no_output "BT-3 formula no double ==" "formula: =SUM" "$_out"
check_output    "BT-3 formula correct display" "formula: SUM" "$_out"

# ---------------------------------------------------------------------------
# BT-5: PPTX query format property filtering works (fill=)
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt5.pptx"
run "$CLI" add "$TMPDIR/bt5.pptx" "/" --type slide
run "$CLI" add "$TMPDIR/bt5.pptx" "/slide[1]" --type shape --prop "text=Hello" --prop "fill=FF0000"
run "$CLI" add "$TMPDIR/bt5.pptx" "/slide[1]" --type shape --prop "text=World" --prop "fill=0000FF"
run "$CLI" query "$TMPDIR/bt5.pptx" "shape[fill=FF0000]"
check_output "BT-5 query fill= filter returns 1 match" "Matches: 1" "$_out"

# BT-5: text= filter (regression — tracked)
run "$CLI" query "$TMPDIR/bt5.pptx" "shape[text=Hello]"
if echo "$_out" | grep -q "Matches: 1"; then
    pass "BT-5 query text= filter (FIXED)"
else
    fail "BT-5 query text= filter" "REGRESSION: shape[text=Hello] got 0 matches, expected 1"
fi

# ---------------------------------------------------------------------------
# BT-6: DOCX header type shows "default"
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt6.docx"
run "$CLI" add "$TMPDIR/bt6.docx" "/" --type header --prop "type=default"
run "$CLI" query "$TMPDIR/bt6.docx" "header"
check_output "BT-6 header type=default" "type: default" "$_out"

# ---------------------------------------------------------------------------
# BT-7: XLSX delete last sheet is blocked
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt7.xlsx"
run "$CLI" remove "$TMPDIR/bt7.xlsx" "/Sheet1"
check_exit   "BT-7 delete last sheet exit code" 1 $_code
check_output "BT-7 delete last sheet error" "Cannot remove the last sheet" "$_out"

# ---------------------------------------------------------------------------
# BT-8: invalid color rejected, existing fill preserved
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt8.pptx"
run "$CLI" add "$TMPDIR/bt8.pptx" "/" --type slide --prop "title=T"
run "$CLI" add "$TMPDIR/bt8.pptx" "/slide[1]" --type shape --prop "text=Test" --prop "fill=FF0000"
run "$CLI" set "$TMPDIR/bt8.pptx" "/slide[1]/shape[2]" --prop "fill=INVALIDCOLOR"
check_exit   "BT-8 invalid color exit code" 1 $_code
check_output "BT-8 invalid color error message" "Invalid color" "$_out"
run "$CLI" get "$TMPDIR/bt8.pptx" "/slide[1]/shape[2]"
check_output "BT-8 fill preserved after invalid set" "fill: FF0000" "$_out"

# ---------------------------------------------------------------------------
# BT-10: PPTX notes passes validation
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt10.pptx"
run "$CLI" add "$TMPDIR/bt10.pptx" "/" --type slide
run "$CLI" add "$TMPDIR/bt10.pptx" "/slide[1]" --type notes --prop "text=Speaker notes"
run "$CLI" validate "$TMPDIR/bt10.pptx"
check_output "BT-10 notes validation passes" "Validation passed" "$_out"

# ---------------------------------------------------------------------------
# BT-11: failed chart add leaves no empty part (file validates)
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt11.pptx"
run "$CLI" add "$TMPDIR/bt11.pptx" "/" --type slide
run "$CLI" add "$TMPDIR/bt11.pptx" "/slide[1]" --type chart --prop "type=INVALID"
run "$CLI" validate "$TMPDIR/bt11.pptx"
check_output "BT-11 validation after failed chart" "Validation passed" "$_out"

# ---------------------------------------------------------------------------
# BT-13: PPTX custom geometry passes validation
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt13.pptx"
run "$CLI" add "$TMPDIR/bt13.pptx" "/" --type slide
run "$CLI" add "$TMPDIR/bt13.pptx" "/slide[1]" --type shape \
    --prop "geometry=custom" --prop "custGeom=M 0 0 L 100 0 L 100 100 Z"
run "$CLI" validate "$TMPDIR/bt13.pptx"
check_output "BT-13 custom geometry validation" "Validation passed" "$_out"

# ---------------------------------------------------------------------------
# BT-14: XLSX invalid color is rejected
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt14.xlsx"
run "$CLI" add "$TMPDIR/bt14.xlsx" "/Sheet1" --type row
run "$CLI" add "$TMPDIR/bt14.xlsx" "/Sheet1/row[1]" --type cell --prop "value=Test" --prop "ref=A1"
run "$CLI" set "$TMPDIR/bt14.xlsx" "/Sheet1/A1" --prop "fill=GGGGGG"
check_exit   "BT-14 invalid XLSX color exit code" 1 $_code
check_output "BT-14 invalid XLSX color error" "Invalid color" "$_out"

# ---------------------------------------------------------------------------
# BT-15: DOCX move --index works
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt15.docx"
run "$CLI" add "$TMPDIR/bt15.docx" "/body" --type paragraph --prop "text=Para 1"
run "$CLI" add "$TMPDIR/bt15.docx" "/body" --type paragraph --prop "text=Para 2"
run "$CLI" add "$TMPDIR/bt15.docx" "/body" --type paragraph --prop "text=Para 3"
run "$CLI" move "$TMPDIR/bt15.docx" "/body/p[3]" --index 0
run "$CLI" get "$TMPDIR/bt15.docx" "/body"
check_output "BT-15 move --index para 3 to first" "Para 3" "$(echo "$_out" | head -4)"

# ---------------------------------------------------------------------------
# BT-16: batch add supports both "path" and "parent" keys
# ---------------------------------------------------------------------------
cat > "$TMPDIR/bt16_batch.json" << 'EOF'
[
  {"command": "add", "parent": "/", "type": "slide", "props": {"title": "From parent key"}},
  {"command": "add", "path": "/slide[1]", "type": "shape", "props": {"text": "From path key"}}
]
EOF
run "$CLI" create "$TMPDIR/bt16.pptx"
run "$CLI" batch "$TMPDIR/bt16.pptx" --input "$TMPDIR/bt16_batch.json"
check_exit   "BT-16 batch exit code" 0 $_code
check_output "BT-16 batch 2 succeeded" "2 succeeded" "$_out"

# ---------------------------------------------------------------------------
# BT-21: PPTX connector appears in DOM
# ---------------------------------------------------------------------------
run "$CLI" create "$TMPDIR/bt21.pptx"
run "$CLI" add "$TMPDIR/bt21.pptx" "/" --type slide
run "$CLI" add "$TMPDIR/bt21.pptx" "/slide[1]" --type connector
run "$CLI" get "$TMPDIR/bt21.pptx" "/slide[1]"
check_output "BT-21 connector in DOM" "connector" "$_out"

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
echo ""
echo "=== Results ==="
echo "PASS: $PASS"
echo "FAIL: $FAIL"
echo "TOTAL: $((PASS + FAIL))"
[ "$FAIL" -eq 0 ]

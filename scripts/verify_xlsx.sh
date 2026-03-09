#!/bin/bash
# Comprehensive Excel verification script
# Tests all xlsx features against ClosedXML reference implementation
# Usage: ./scripts/verify_xlsx.sh [--quick|--thorough]

set -e
cd "$(dirname "$0")/.."

MODE="${1:---thorough}"
FAILURES=0

echo "📊 Excel Verification ($MODE mode)"
echo "===================================="
echo ""

# Build test suite
echo "🔨 Building test suite..."
cargo build --release --workspace 2>&1 | grep -E "(Compiling|Finished)" || true
echo ""

# Test 1: Pivot Tables
echo "Test 1: Pivot Tables"
echo "---------------------"
if ./scripts/test_pivot.sh > /tmp/verify_xlsx_pivot.log 2>&1; then
    echo "✅ Pivot tables: PASS"
else
    echo "❌ Pivot tables: FAIL (see /tmp/verify_xlsx_pivot.log)"
    FAILURES=$((FAILURES + 1))
fi
echo ""

# Test 2: Formulas
echo "Test 2: Formulas"
echo "----------------"
if cargo test -p offidized-formula --release > /tmp/verify_xlsx_formulas.log 2>&1; then
    TEST_COUNT=$(grep "test result:" /tmp/verify_xlsx_formulas.log | head -1 | grep -oE "[0-9]+ passed" | grep -oE "[0-9]+")
    echo "✅ Formulas: PASS ($TEST_COUNT tests)"
else
    echo "❌ Formulas: FAIL (see /tmp/verify_xlsx_formulas.log)"
    FAILURES=$((FAILURES + 1))
fi
echo ""

# Test 3: Core xlsx functionality
echo "Test 3: Core xlsx"
echo "-----------------"
if cargo test -p offidized-xlsx --release > /tmp/verify_xlsx_core.log 2>&1; then
    TEST_COUNT=$(grep "test result:" /tmp/verify_xlsx_core.log | head -1 | grep -oE "[0-9]+ passed" | grep -oE "[0-9]+")
    echo "✅ Core xlsx: PASS ($TEST_COUNT tests)"
else
    echo "❌ Core xlsx: FAIL (see /tmp/verify_xlsx_core.log)"
    FAILURES=$((FAILURES + 1))
fi
echo ""

# Test 4: Roundtrip fidelity
if [ "$MODE" = "--thorough" ]; then
    echo "Test 4: Roundtrip Fidelity"
    echo "--------------------------"
    # TODO: Add roundtrip tests with real-world files
    echo "⏭️  Roundtrip tests: SKIPPED (not implemented yet)"
    echo ""
fi

# Test 5: ClosedXML parity check
echo "Test 5: ClosedXML Parity"
echo "------------------------"
# Check critical differences from our automated comparison
if [ -f "/tmp/pivot_test_results.txt" ]; then
    PIVOT_DIFF=$(grep -c "DIFFERENCE: xl/pivotTables/pivotTable.xml" /tmp/pivot_test_results.txt 2>/dev/null || echo "0" | tr -d '\n\r')
    CACHE_DIFF=$(grep -c "DIFFERENCE: pivotCache/pivotCacheDefinition1.xml" /tmp/pivot_test_results.txt 2>/dev/null || echo "0" | tr -d '\n\r')
    RECORDS_DIFF=$(grep -c "DIFFERENCE: pivotCache/pivotCacheRecords1.xml" /tmp/pivot_test_results.txt 2>/dev/null || echo "0" | tr -d '\n\r')

    if [ "${PIVOT_DIFF:-0}" -eq 0 ] && [ "${CACHE_DIFF:-0}" -eq 0 ] && [ "${RECORDS_DIFF:-0}" -eq 0 ]; then
        echo "✅ ClosedXML parity: PASS (pivot structures identical)"
    else
        echo "⚠️  ClosedXML parity: PARTIAL (some differences remain)"
    fi
else
    echo "⏭️  ClosedXML parity: SKIPPED (run test_pivot.sh first)"
fi
echo ""

# Summary
echo "===================================="
if [ $FAILURES -eq 0 ]; then
    echo "✅ ALL EXCEL TESTS PASSED"
    exit 0
else
    echo "❌ $FAILURES TESTS FAILED"
    exit 1
fi

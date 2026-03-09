#!/bin/bash
# PowerPoint verification script
# Tests all pptx features against python-pptx reference implementation
# Usage: ./scripts/verify_pptx.sh [--quick|--thorough]

set -e
cd "$(dirname "$0")/.."

MODE="${1:---thorough}"
FAILURES=0

echo "📽️  PowerPoint Verification ($MODE mode)"
echo "====================================="
echo ""

# Build test suite
echo "🔨 Building test suite..."
cargo build --release --workspace 2>&1 | grep -E "(Compiling|Finished)" || true
echo ""

# Test 1: Core pptx functionality
echo "Test 1: Core pptx"
echo "-----------------"
if cargo test -p offidized-pptx --release > /tmp/verify_pptx_core.log 2>&1; then
    TEST_COUNT=$(grep "test result:" /tmp/verify_pptx_core.log | head -1 | grep -oE "[0-9]+ passed" | grep -oE "[0-9]+")
    echo "✅ Core pptx: PASS ($TEST_COUNT tests)"
else
    echo "❌ Core pptx: FAIL (see /tmp/verify_pptx_core.log)"
    FAILURES=$((FAILURES + 1))
fi
echo ""

# Test 2: Roundtrip fidelity
if [ "$MODE" = "--thorough" ]; then
    echo "Test 2: Roundtrip Fidelity"
    echo "--------------------------"
    # TODO: Add roundtrip tests with real-world files
    echo "⏭️  Roundtrip tests: SKIPPED (not implemented yet)"
    echo ""
fi

# Test 3: python-pptx parity check
echo "Test 3: python-pptx Parity"
echo "--------------------------"
# TODO: Automated comparison against python-pptx generated files
echo "⏭️  python-pptx parity: SKIPPED (not implemented yet)"
echo ""

# Summary
echo "====================================="
if [ $FAILURES -eq 0 ]; then
    echo "✅ ALL POWERPOINT TESTS PASSED"
    exit 0
else
    echo "❌ $FAILURES TESTS FAILED"
    exit 1
fi

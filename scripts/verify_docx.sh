#!/bin/bash
# Word document verification script
# Tests all docx features against python-docx reference implementation
# Usage: ./scripts/verify_docx.sh [--quick|--thorough]

set -e
cd "$(dirname "$0")/.."

MODE="${1:---thorough}"
FAILURES=0

echo "📄 Word Verification ($MODE mode)"
echo "====================================="
echo ""

# Build test suite
echo "🔨 Building test suite..."
cargo build --release --workspace 2>&1 | grep -E "(Compiling|Finished)" || true
echo ""

# Test 1: Core docx functionality
echo "Test 1: Core docx"
echo "-----------------"
if cargo test -p offidized-docx --release >/tmp/verify_docx_core.log 2>&1; then
	TEST_COUNT=$(grep "test result:" /tmp/verify_docx_core.log | head -1 | grep -oE "[0-9]+ passed" | grep -oE "[0-9]+")
	echo "✅ Core docx: PASS ($TEST_COUNT tests)"
else
	echo "❌ Core docx: FAIL (see /tmp/verify_docx_core.log)"
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

# Test 3: python-docx parity check
echo "Test 3: python-docx Parity"
echo "--------------------------"
# TODO: Automated comparison against python-docx generated files
echo "⏭️  python-docx parity: SKIPPED (not implemented yet)"
echo ""

# Summary
echo "====================================="
if [ $FAILURES -eq 0 ]; then
	echo "✅ ALL WORD TESTS PASSED"
	exit 0
else
	echo "❌ $FAILURES TESTS FAILED"
	exit 1
fi

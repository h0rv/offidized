#!/bin/bash
# Master verification script - runs all format verifications
# Usage: ./scripts/verify_all.sh [--quick|--thorough]

set -e
cd "$(dirname "$0")/.."

MODE="${1:---thorough}"

echo "========================================="
echo "OFFIDIZED COMPREHENSIVE VERIFICATION"
echo "Mode: $MODE"
echo "========================================="
echo ""

# Track results
XLSX_PASS=0
DOCX_PASS=0
PPTX_PASS=0

# Run xlsx verification
echo "📊 Verifying Excel (xlsx)..."
if ./scripts/verify_xlsx.sh "$MODE"; then
	XLSX_PASS=1
	echo "✅ Excel verification PASSED"
else
	echo "❌ Excel verification FAILED"
fi
echo ""

# Run docx verification
echo "📄 Verifying Word (docx)..."
if ./scripts/verify_docx.sh "$MODE"; then
	DOCX_PASS=1
	echo "✅ Word verification PASSED"
else
	echo "❌ Word verification FAILED"
fi
echo ""

# Run pptx verification
echo "📽️  Verifying PowerPoint (pptx)..."
if ./scripts/verify_pptx.sh "$MODE"; then
	PPTX_PASS=1
	echo "✅ PowerPoint verification PASSED"
else
	echo "❌ PowerPoint verification FAILED"
fi
echo ""

# Summary
echo "========================================="
echo "FINAL RESULTS"
echo "========================================="
echo "Excel:      $([ $XLSX_PASS -eq 1 ] && echo '✅ PASS' || echo '❌ FAIL')"
echo "Word:       $([ $DOCX_PASS -eq 1 ] && echo '✅ PASS' || echo '❌ FAIL')"
echo "PowerPoint: $([ $PPTX_PASS -eq 1 ] && echo '✅ PASS' || echo '❌ FAIL')"
echo "========================================="

# Exit with error if any failed
TOTAL_PASS=$((XLSX_PASS + DOCX_PASS + PPTX_PASS))
if [ $TOTAL_PASS -eq 3 ]; then
	echo "🎉 ALL VERIFICATIONS PASSED"
	exit 0
else
	echo "⚠️  SOME VERIFICATIONS FAILED"
	exit 1
fi

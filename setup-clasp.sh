#!/bin/bash
# Quick setup script for clasp development

set -e

echo "🚀 QA Test Management - Clasp Setup"
echo "===================================="
echo ""

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "❌ clasp is not installed"
    echo "Installing clasp globally..."
    npm install -g @google/clasp
else
    echo "✅ clasp is already installed ($(clasp --version))"
fi

echo ""

# Check if user is logged in
if [ ! -f ~/.clasprc.json ]; then
    echo "🔐 Logging in to Google Account..."
    clasp login
else
    echo "✅ Already logged in to clasp"
fi

echo ""
echo "⚠️  IMPORTANT: Enable Apps Script API"
echo "   Before you can push/pull, you must enable the Apps Script API:"
echo "   1. Visit: https://script.google.com/home/usersettings"
echo "   2. Toggle ON 'Google Apps Script API'"
echo "   3. Wait 1-2 minutes for propagation"
echo ""
read -p "Have you enabled the Apps Script API? (y/n) " -n 1 -r
echo ""
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "⚠️  Please enable the API first, then re-run this script"
    exit 1
fi

echo ""
echo "📋 Verifying project configurations..."

# Verify Dashboard
if [ -f projects/qa-dashboard/.clasp.json ]; then
    echo "✅ Dashboard config found"
    cd projects/qa-dashboard
    clasp status 2>/dev/null || echo "⚠️  Dashboard: Run 'cd projects/qa-dashboard && clasp status' to verify"
    cd ../..
else
    echo "❌ Dashboard .clasp.json not found"
fi

# Verify Test Management
if [ -f projects/qa-test-management/.clasp.json ]; then
    echo "✅ Test Management config found"
    cd projects/qa-test-management
    clasp status 2>/dev/null || echo "⚠️  Test Management: Run 'cd projects/qa-test-management && clasp status' to verify"
    cd ../..
else
    echo "❌ Test Management .clasp.json not found"
fi

# Verify MOM Rolling
if [ -f projects/mom-rolling-pic/.clasp.json ]; then
    echo "✅ MOM Rolling PIC config found"
    cd projects/mom-rolling-pic
    clasp status 2>/dev/null || echo "⚠️  MOM Rolling: Run 'cd projects/mom-rolling-pic && clasp status' to verify"
    cd ../..
else
    echo "❌ MOM Rolling .clasp.json not found"
fi

echo ""
echo "✨ Setup complete!"
echo ""
echo "Next steps:"
echo "1. Edit files in projects/qa-dashboard/src/ or projects/qa-test-management/src/"
echo "2. Push changes: cd projects/qa-dashboard && clasp push"
echo "3. Or commit to GitHub for auto-deploy"
echo ""
echo "See CLASP_WORKFLOW.md for detailed instructions"

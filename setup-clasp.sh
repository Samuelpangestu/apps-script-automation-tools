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
echo "📋 Verifying project configurations..."

# Verify Dashboard
if [ -f qa-dashboard/.clasp.json ]; then
    echo "✅ Dashboard config found"
    cd qa-dashboard
    clasp status 2>/dev/null || echo "⚠️  Dashboard: Run 'cd qa-dashboard && clasp status' to verify"
    cd ..
else
    echo "❌ Dashboard .clasp.json not found"
fi

# Verify Test Management
if [ -f qa-test-management/.clasp.json ]; then
    echo "✅ Test Management config found"
    cd qa-test-management
    clasp status 2>/dev/null || echo "⚠️  Test Management: Run 'cd qa-test-management && clasp status' to verify"
    cd ..
else
    echo "❌ Test Management .clasp.json not found"
fi

echo ""
echo "✨ Setup complete!"
echo ""
echo "Next steps:"
echo "1. Edit files in qa-dashboard/src/ or qa-test-management/src/"
echo "2. Push changes: cd qa-dashboard && clasp push"
echo "3. Or commit to GitHub for auto-deploy"
echo ""
echo "See CLASP_WORKFLOW.md for detailed instructions"

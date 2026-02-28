#!/bin/bash
# Repository Rename Helper Script
# Renames repository from qa-test-management-template to apps-script-automation-tools

set -e

OLD_NAME="qa-test-management-template"
NEW_NAME="apps-script-automation-tools"
GITHUB_USER="Samuelpangestu"

echo "🔄 Repository Rename Helper"
echo "============================"
echo ""
echo "Old name: $OLD_NAME"
echo "New name: $NEW_NAME"
echo ""

# Step 1: Manual GitHub rename
echo "📋 STEP 1: Rename on GitHub (Manual)"
echo "-----------------------------------"
echo "1. Open: https://github.com/$GITHUB_USER/$OLD_NAME/settings"
echo "2. In 'Repository name' field, change to: $NEW_NAME"
echo "3. Click 'Rename' button"
echo "4. GitHub will show a warning - click 'I understand, rename repository'"
echo ""
read -p "✅ Have you completed the GitHub rename? (y/n) " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "⚠️  Please complete GitHub rename first, then re-run this script"
    exit 1
fi

echo ""
echo "📋 STEP 2: Update Local Git Remote"
echo "-----------------------------------"

# Get current remote URL
CURRENT_REMOTE=$(git remote get-url origin)
echo "Current remote: $CURRENT_REMOTE"

# Update remote URL
NEW_REMOTE="https://github.com/$GITHUB_USER/$NEW_NAME.git"
echo "New remote: $NEW_REMOTE"

git remote set-url origin "$NEW_REMOTE"
echo "✅ Git remote updated"

echo ""
echo "📋 STEP 3: Verify Connection"
echo "-----------------------------------"

if git ls-remote origin &>/dev/null; then
    echo "✅ Successfully connected to new repository URL"
else
    echo "❌ Cannot connect to new repository. Please check:"
    echo "   - Repository was renamed correctly on GitHub"
    echo "   - New name is: $NEW_NAME"
    exit 1
fi

echo ""
echo "✅ Repository rename complete!"
echo ""
echo "Next steps:"
echo "1. Update documentation references (will be done automatically)"
echo "2. Commit and push changes"
echo ""

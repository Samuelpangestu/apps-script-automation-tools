#!/bin/bash
# Sync script - Pull from Apps Script and push to Git

set -e

echo "🔄 Syncing from Apps Script to Git"
echo "===================================="
echo ""

# Check if in qa-dashboard directory
if [ ! -f "qa-dashboard/.clasp.json" ]; then
    echo "❌ Error: Run this script from project root"
    exit 1
fi

echo "📥 Step 1: Pulling from Apps Script..."
cd qa-dashboard
clasp pull

if [ $? -ne 0 ]; then
    echo "❌ clasp pull failed. You may need to re-login:"
    echo "   Run: clasp login"
    exit 1
fi

echo ""
echo "✅ Pull complete!"
echo ""

cd ..

# Check if there are changes (including untracked files)
if git diff --quiet && git diff --cached --quiet && [ -z "$(git ls-files --others --exclude-standard qa-dashboard/src/)" ]; then
    echo "ℹ️  No changes detected. Already in sync."
    exit 0
fi

echo "📋 Changes detected:"
git status --short

echo ""
read -p "💾 Commit and push these changes? (y/n) " -n 1 -r
echo ""

if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo ""
    read -p "📝 Commit message (or press Enter for default): " commit_msg

    if [ -z "$commit_msg" ]; then
        commit_msg="Sync latest changes from Apps Script Editor"
    fi

    git add qa-dashboard/src/
    git commit -m "$commit_msg

Synced via clasp pull from Apps Script Editor.

🤖 Auto-synced using sync-from-appscript.sh"

    echo ""
    echo "⬆️  Pushing to GitHub..."
    git push origin main

    echo ""
    echo "✅ Sync complete!"
    echo "   - Pulled from Apps Script"
    echo "   - Committed to Git"
    echo "   - Pushed to GitHub"
else
    echo ""
    echo "⏭️  Skipped commit. Changes are in working directory."
    echo "   Run 'git status' to see what changed."
fi

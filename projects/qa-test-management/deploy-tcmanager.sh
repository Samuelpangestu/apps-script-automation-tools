#!/bin/bash
# Deploy TC Manager to Apps Script Project
# Usage: ./deploy-tcmanager.sh <SCRIPT_ID>

set -e

SCRIPT_ID="$1"

if [ -z "$SCRIPT_ID" ]; then
  echo "❌ Error: Script ID required"
  echo ""
  echo "Usage: ./deploy-tcmanager.sh <SCRIPT_ID>"
  echo ""
  echo "Example:"
  echo "  ./deploy-tcmanager.sh 1EMrV7bdWIWN7jRMDe9TyAMRjpnBwz5MsCWqWkezgtI8uJuETIlA38mDD"
  echo ""
  echo "How to get Script ID:"
  echo "  1. Open target QATM spreadsheet"
  echo "  2. Click: Extensions → Apps Script"
  echo "  3. Click: ⚙️ Project Settings (gear icon)"
  echo "  4. Copy 'Script ID'"
  exit 1
fi

echo "🚀 Deploying TC Manager to Script ID: $SCRIPT_ID"
echo ""

# Create temporary deployment folder
DEPLOY_DIR="/tmp/qatm-deploy-$(date +%s)"
mkdir -p "$DEPLOY_DIR/src"

# Create .clasp.json
cat > "$DEPLOY_DIR/.clasp.json" << EOF
{
  "scriptId": "$SCRIPT_ID",
  "rootDir": "./src"
}
EOF

# Copy TC Manager files
echo "📦 Copying TC Manager files..."
cp "$(dirname "$0")/src/TCManager.js" "$DEPLOY_DIR/src/"
cp "$(dirname "$0")/src/appsscript.json" "$DEPLOY_DIR/src/" 2>/dev/null || {
  # Create appsscript.json if not exists
  cat > "$DEPLOY_DIR/src/appsscript.json" << 'EOF'
{
  "timeZone": "Asia/Jakarta",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
EOF
}

echo "✅ Files ready:"
ls -lh "$DEPLOY_DIR/src/"
echo ""

# Deploy with clasp
echo "🚀 Deploying to Apps Script..."
cd "$DEPLOY_DIR"
clasp push --force

if [ $? -eq 0 ]; then
  echo ""
  echo "✅ Deployment successful!"
  echo ""
  echo "📍 Next steps:"
  echo "  1. Open target spreadsheet"
  echo "  2. Refresh (Ctrl+Shift+R or F5)"
  echo "  3. Check menu bar for: 🔧 TC Manager"
  echo "  4. Test features: Insert, Bulk Insert, Delete, Mark as Deprecated"
  echo ""
  echo "⚠️ If old menus still appear:"
  echo "  1. Extensions → Apps Script"
  echo "  2. Delete all files EXCEPT TCManager.gs and appsscript.json"
  echo "  3. Save and refresh spreadsheet"
  echo ""
else
  echo ""
  echo "❌ Deployment failed!"
  echo ""
  echo "Common issues:"
  echo "  - Authentication error: Run 'clasp login' first"
  echo "  - Wrong Script ID: Verify Script ID from Apps Script editor"
  echo "  - Permission error: Check Apps Script project permissions"
  exit 1
fi

# Cleanup
echo "🧹 Cleaning up temporary files..."
rm -rf "$DEPLOY_DIR"
echo "✅ Done!"

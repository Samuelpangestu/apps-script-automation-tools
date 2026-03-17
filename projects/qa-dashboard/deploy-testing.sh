#!/bin/bash

# Deploy QA Dashboard to TESTING environment
# Automatically restores to PRODUCTION config after deploy

TESTING_SCRIPT_ID="1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3"
PRODUCTION_SCRIPT_ID="1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO"

echo "======================================"
echo "Deploying to TESTING environment"
echo "======================================"
echo ""
echo "Script ID: $TESTING_SCRIPT_ID"
echo ""

# Backup current config
cp .clasp.json .clasp.json.bak

# Switch to testing script ID
cat > .clasp.json <<EOF
{
  "scriptId": "$TESTING_SCRIPT_ID",
  "rootDir": "./src"
}
EOF

echo "Pushing to Apps Script..."
clasp push

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Successfully deployed to TESTING!"
    echo ""
    echo "Open in Apps Script Editor:"
    echo "https://script.google.com/d/$TESTING_SCRIPT_ID/edit"

    # Restore production config
    echo ""
    echo "Restoring .clasp.json to PRODUCTION..."
    cat > .clasp.json <<EOF
{
  "scriptId": "$PRODUCTION_SCRIPT_ID",
  "rootDir": "./src"
}
EOF
    rm -f .clasp.json.bak
    echo "✅ Restored to production config"
else
    echo ""
    echo "❌ Deployment failed!"
    # Restore backup
    mv .clasp.json.bak .clasp.json
    exit 1
fi

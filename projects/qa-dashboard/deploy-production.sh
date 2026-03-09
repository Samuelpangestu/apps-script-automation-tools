#!/bin/bash

# Deploy QA Dashboard to PRODUCTION environment

PRODUCTION_SCRIPT_ID="1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO"
TESTING_SCRIPT_ID="1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3"

echo "======================================"
echo "⚠️  DEPLOYING TO PRODUCTION  ⚠️"
echo "======================================"
echo ""
echo "Script ID: $PRODUCTION_SCRIPT_ID"
echo ""
read -p "Are you sure you want to deploy to PRODUCTION? (yes/no): " confirm

if [ "$confirm" != "yes" ]; then
    echo "Deployment cancelled."
    exit 0
fi

echo ""
echo "Updating .clasp.json to production..."

# Temporarily change to production script ID
cat > .clasp.json <<EOF
{
  "scriptId": "$PRODUCTION_SCRIPT_ID",
  "rootDir": "./src"
}
EOF

echo "Pushing to Apps Script..."
clasp push

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Successfully deployed to PRODUCTION!"
    echo ""
    echo "Open in Apps Script Editor:"
    echo "https://script.google.com/d/$PRODUCTION_SCRIPT_ID/edit"
else
    echo ""
    echo "❌ Deployment failed!"
fi

# Revert back to testing
echo ""
echo "Reverting .clasp.json back to TESTING..."
cat > .clasp.json <<EOF
{
  "scriptId": "$TESTING_SCRIPT_ID",
  "rootDir": "./src"
}
EOF

echo "✅ Reverted to testing script ID"

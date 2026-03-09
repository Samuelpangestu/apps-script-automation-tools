#!/bin/bash

# Deploy QA Dashboard to TESTING environment

TESTING_SCRIPT_ID="1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3"

echo "======================================"
echo "Deploying to TESTING environment"
echo "======================================"
echo ""
echo "Script ID: $TESTING_SCRIPT_ID"
echo ""

# Ensure .clasp.json has testing script ID
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
else
    echo ""
    echo "❌ Deployment failed!"
    exit 1
fi

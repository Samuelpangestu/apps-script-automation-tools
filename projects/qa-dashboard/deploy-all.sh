#!/bin/bash

# Deploy QA Dashboard to BOTH Testing and Production environments

echo "======================================"
echo "Deploy to BOTH environments"
echo "======================================"
echo ""
echo "This will deploy to:"
echo "1. Testing environment"
echo "2. Production environment"
echo ""
read -p "Continue? (yes/no): " confirm

if [ "$confirm" != "yes" ]; then
    echo "Deployment cancelled."
    exit 0
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "Step 1/2: Deploying to TESTING..."
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

./deploy-testing.sh

if [ $? -ne 0 ]; then
    echo ""
    echo "❌ Testing deployment failed. Stopping."
    exit 1
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "Step 2/2: Deploying to PRODUCTION..."
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Deploy to production (will ask for confirmation again)
./deploy-production.sh

if [ $? -eq 0 ]; then
    echo ""
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    echo "✅ All deployments completed!"
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
else
    echo ""
    echo "❌ Production deployment failed."
    exit 1
fi

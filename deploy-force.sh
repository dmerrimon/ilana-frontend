#!/bin/bash

echo "🚀 FORCE DEPLOYING TO AZURE STATIC WEB APPS..."

# Set the deployment token
AZURE_TOKEN="4008323677e9ae8f7fd605d4c986eb704f29b7846d92de4e4d61f6bbf240ca5c03-fb986235-99af-4e20-a4ee-157d2075be9800f14110bbaa4e0f"

# Create a timestamp for cache busting
TIMESTAMP=$(date +%s)

echo "📝 Adding cache-busting timestamp: $TIMESTAMP"

# Update CSS link with new timestamp
sed -i.bak "s/style.css?v=[0-9.]*/style.css?v=$TIMESTAMP/g" taskpane.html
sed -i.bak "s/style.css?v=[0-9.]*/style.css?v=$TIMESTAMP/g" index.html

# Use the Azure Static Web Apps CLI to deploy
npx @azure/static-web-apps-cli deploy \
  --deployment-token="$AZURE_TOKEN" \
  --app-location="." \
  --verbose

echo "✅ Force deployment completed!"
echo "🔄 Please wait 2-3 minutes for changes to propagate"
echo "🎯 Then refresh your Word add-in"
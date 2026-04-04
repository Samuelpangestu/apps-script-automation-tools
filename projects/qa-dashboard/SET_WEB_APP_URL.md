# Set Web App URL untuk Notifications

Web App URL baru: 
```
https://script.google.com/a/macros/inadigital.co.id/s/AKfycby2xYW-5VJmy-icA-joSGMsQEPyHpsiVqOZ_x3-0uMOkf_IcmR3Y-JWqq6-DLKWI4Ef2A/exec
```

Deployment ID untuk future deploys:
```
AKfycby2xYW-5VJmy-icA-joSGMsQEPyHpsiVqOZ_x3-0uMOkf_IcmR3Y-JWqq6-DLKWI4Ef2A
```

## Cara Set URL untuk Notifications:

### Option 1: Via Apps Script Console (Recommended)
1. Buka Apps Script: https://script.google.com/d/1LJ83OATTAp7ChDWGkrSTg0b9KmMhOABISBrAJrB54JksjQ7mi5oNB7C3/edit
2. Buka console (Ctrl/Cmd + Enter atau View > Logs)
3. Run function `setWebAppUrl` dengan parameter:
   ```javascript
   setWebAppUrl('https://script.google.com/a/macros/inadigital.co.id/s/AKfycby2xYW-5VJmy-icA-joSGMsQEPyHpsiVqOZ_x3-0uMOkf_IcmR3Y-JWqq6-DLKWI4Ef2A/exec')
   ```

### Option 2: Via Script Editor
1. Buka Notifications.js
2. Find function `setWebAppUrl(url)`
3. Temporarily modify it:
   ```javascript
   function setWebAppUrl(url) {
     url = 'https://script.google.com/a/macros/inadigital.co.id/s/AKfycby2xYW-5VJmy-icA-joSGMsQEPyHpsiVqOZ_x3-0uMOkf_IcmR3Y-JWqq6-DLKWI4Ef2A/exec';
     const scriptProps = PropertiesService.getScriptProperties();
     scriptProps.setProperty('WEB_APP_URL', url);
     Logger.log('✅ Web App URL set to: ' + url);
   }
   ```
4. Run `setWebAppUrl()`
5. Revert the modification after

## Deploy Command:
```bash
clasp deploy -i AKfycby2xYW-5VJmy-icA-joSGMsQEPyHpsiVqOZ_x3-0uMOkf_IcmR3Y-JWqq6-DLKWI4Ef2A -d "description"
```

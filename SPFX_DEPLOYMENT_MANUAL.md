# SPFx Deployment Manual Steps

## Status
- ✅ **Proxy deployed:** https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net (healthy)
- ✅ **SPFx built:** `spfx-sample/spfx-profile-agent/sharepoint/solution/spfx-profile-agent.sppkg` (19.2 KB)
- ⚠️ **PnP.PowerShell deployment blocked:** App registration lacks SharePoint Admin permissions

## Manual Deployment (UI)

### Option A: Upload to Tenant App Catalog (Recommended)

1. Navigate to your tenant's App Catalog site:
   ```
   https://mngenv168112.sharepoint.com/sites/appcatalog
   ```

2. Click **Manage Apps → Upload**

3. Upload the SPFx package:
   ```
   spfx-sample/spfx-profile-agent/sharepoint/solution/spfx-profile-agent.sppkg
   ```

4. Click **Deploy** in the dialog

5. Verify it shows as published in the App Catalog

### Option B: Upload Directly to Dev Site Collection

1. Navigate to your dev site:
   ```
   https://mngenv168112.sharepoint.com/sites/gbelelenky-dev-sp
   ```

2. Go to **Site Contents → Apps from Your Organization**

3. Click **New → Upload**

4. Upload the SPFx package

5. Once uploaded, add the **Profile Agent** web part to a page:
   - Edit the page
   - Click **+** to add a web part
   - Search for "Profile Agent"
   - Add to page
   - Publish page

## End-to-End Testing

After deployment, navigate to:
```
https://mngenv168112.sharepoint.com/sites/gbelelenky-dev-sp/SitePages/CollabHome.aspx
```

### Test Cases

1. **Greeting**
   - Click "Greeting you" button
   - Expected: One-sentence name greeting (e.g., "Hello, Jane!")
   - Check browser console for errors

2. **Profile Query**
   - Type: "What is my profile?"
   - Expected: Full user profile (name, email, country)
   - Verify profile is country-filtered FAQs

3. **FAQ Query (Country-Filtered)**
   - Type: "Where do I find vacation policies?"
   - Expected: Country-relevant FAQ answer based on user's location
   - Verify response comes from search_faq tool

4. **Multi-Turn Chat**
   - Type follow-up question
   - Expected: Response includes previous context
   - Verify no errors in browser console

### Logs to Monitor

**App Insights:**
1. Navigate to your App Insights resource in Azure
2. Check **Traces** for end-to-end flow:
   - SPFx calls proxy
   - Proxy validates JWT
   - Proxy acquires OBO token
   - Proxy calls Foundry agent
   - Agent calls search_faq tool

**Browser Console:**
- F12 → Console tab
- Look for any network errors or 4xx/5xx responses

## Troubleshooting

### If SPFx upload fails via UI
- Verify you have **Site Owner** permissions on the site
- Check that the `.sppkg` file is valid (not corrupted)
- Try uploading to a single site collection first (Option B) before tenant catalog

### If greeting fails
- Check browser console for network errors
- Verify proxy health: `https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net/liveness`
- Check App Insights traces for proxy errors

### If profile is empty
- First call may trigger Foundry Toolbox connection consent flow
- User may see sign-in prompt; accept and retry
- Second call should return full profile

### If FAQ queries return empty
- Verify `search_faq` tool exists in Foundry agent
- Verify Azure AI Search index contains FAQ data
- Check App Insights for tool call errors

## Programmatic Deployment (Future)

To automate SPFx deployment, the app registration needs:
1. SharePoint Admin role assignment in Azure AD
2. Tenant App Catalog admin permissions

Once added, PnP.PowerShell will work:
```powershell
Connect-PnPOnline -Url "https://mngenv168112-admin.sharepoint.com" `
  -ClientId "7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a" `
  -ClientSecret "<secret>"

Add-PnPApp -Path "spfx-profile-agent.sppkg" -Scope Tenant -Publish -Overwrite
```

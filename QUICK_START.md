# 🚀 Quick Setup Guide - Personal Outlook Account

## Time: ~15 minutes

---

## Step 1: Install Software (5 min)

### Node.js
1. Go to https://nodejs.org
2. Download **LTS** version
3. Install (click Next through everything)
4. Open terminal, type: `node --version` (should show v18 or higher)

### Python
1. Go to https://python.org/downloads
2. Download Python 3.11+
3. **CHECK THE BOX** "Add Python to PATH" during install
4. Open terminal, type: `python --version` (should show 3.11+)

---

## Step 2: Azure App Registration (5 min)

1. Go to https://portal.azure.com
2. Sign in with your **personal Microsoft account** (the @outlook.com one)
3. Search **"App registrations"** in the top search bar
4. Click **"+ New registration"**

Fill in:
- **Name:** `Hexa RFQ Demo`
- **Supported account types:** Select **"Personal Microsoft accounts only"**
- **Redirect URI:** 
  - Dropdown: **Single-page application (SPA)**
  - URL: `http://localhost:3000/auth-callback.html`

5. Click **Register**

6. **COPY THE APPLICATION (CLIENT) ID** - you need this!
   - It looks like: `12345678-abcd-1234-efgh-123456789abc`

7. Click **"API permissions"** in left sidebar
8. Click **"+ Add a permission"**
9. Select **"Microsoft Graph"**
10. Select **"Delegated permissions"**
11. Search and check:
    - ✅ `Mail.Send`
    - ✅ `Mail.ReadWrite`  
    - ✅ `User.Read`
12. Click **"Add permissions"**

---

## Step 3: Setup the Project (3 min)

1. Extract `rfq-outlook-addin.zip` to a folder

2. Open terminal in that folder:
```bash
cd rfq-outlook-addin
npm install
pip install openpyxl reportlab
```

3. **Edit the code** with your Azure App ID:

Open `src/taskpane/taskpane.js` and find this line (near the top):
```javascript
azureClientId: 'YOUR_AZURE_CLIENT_ID_HERE',
```

Replace with your actual ID:
```javascript
azureClientId: '12345678-abcd-1234-efgh-123456789abc',
```

---

## Step 4: Run It! (2 min)

```bash
npm start
```

You should see:
```
🚀 Hexa RFQ Manager Server Running
   Local:    http://localhost:3000
```

---

## Step 5: Test the Demo

1. Open http://localhost:3000/taskpane.html in Chrome/Edge

2. The demo RFQ will auto-load showing:
   - Specs extracted from email
   - Missing specs
   - Technical questions

3. Click **"Send All to Engineering"**
   - A Microsoft login popup will appear
   - Sign in with your @outlook.com account
   - Grant permissions when asked

4. **Check your Outlook inbox!** You'll see:
   - An email you "sent" to engineering
   - 3 seconds later: A "reply" from engineering (threaded!)

5. Click **"Send Combined Reply to Client"**
   - Check inbox again - more emails, properly threaded!

6. Click **"Generate Quote Documents"**
   - Download your Excel spec sheet
   - Download your PDF quote

---

## What's Happening

When you click buttons, the app:
1. Sends a real email via Microsoft Graph API (appears in your Sent folder)
2. Creates a "reply" email directly in your inbox (simulating a response)
3. Both emails are properly threaded in Outlook

All emails go to/from YOUR account - it's a demo showing the workflow.

---

## Troubleshooting

### "Login popup blocked"
- Allow popups for localhost in your browser

### "AADSTS50011: Reply URL mismatch"
- Go back to Azure Portal > App registrations > Your app > Authentication
- Make sure redirect URI is exactly: `http://localhost:3000/auth-callback.html`
- Make sure it's set as **SPA** (not Web)

### "Insufficient privileges"
- Go to API permissions and make sure Mail.Send and Mail.ReadWrite are added

### No emails appearing
- Check browser console (F12) for errors
- Make sure you completed the Microsoft login

---

## Next Steps

For a real production deployment:
1. Deploy server to Azure/AWS/etc
2. Get SSL certificate for HTTPS
3. Update redirect URI in Azure to production URL
4. Sideload as Outlook add-in (see main README)

---

## Files Overview

| File | What it does |
|------|--------------|
| `src/taskpane/taskpane.html` | The UI |
| `src/taskpane/taskpane.js` | Main logic + your Azure Client ID |
| `src/services/browserEmailService.js` | Sends emails via Graph API |
| `server.js` | Serves files, generates documents |

# Azure Setup Guide for RFQ Outlook Add-in

This guide walks you through setting up Azure AD so that the add-in can send real emails.

## What This Does

When fully configured:
1. You click "Send to Engineering" → Real email appears in your Outlook inbox
2. 3 seconds later → A reply "from Engineering" appears in the same thread
3. You click "Send to Client" → Real email appears
4. 3 seconds later → A reply "from Client" appears with complete specs
5. All emails are **properly threaded** in Outlook

---

## Step 1: Create an Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Search for **"App registrations"** in the top search bar
3. Click **"+ New registration"**
4. Fill in:
   - **Name:** `Hexa RFQ Add-in`
   - **Supported account types:** "Accounts in this organizational directory only"
   - **Redirect URI:** Leave blank for now
5. Click **Register**

**Copy these values** (you'll need them for `.env`):
- **Application (client) ID** - looks like `12345678-abcd-1234-efgh-123456789abc`
- **Directory (tenant) ID** - same format (click "Overview" to see it)

---

## Step 2: Create a Client Secret

1. In your app registration, click **"Certificates & secrets"** in the left sidebar
2. Click **"+ New client secret"**
3. Description: `RFQ Add-in Secret`
4. Expires: Choose 24 months
5. Click **Add**
6. **IMMEDIATELY copy the Value** (you can only see it once!)

This is your `AZURE_CLIENT_SECRET`.

---

## Step 3: Add API Permissions

1. Click **"API permissions"** in the left sidebar
2. Click **"+ Add a permission"**
3. Select **"Microsoft Graph"**
4. Select **"Application permissions"** (not Delegated!)
5. Search and check these permissions:
   - `Mail.Send`
   - `Mail.ReadWrite`
   - `User.Read.All`
6. Click **"Add permissions"**
7. Click **"Grant admin consent for [your org]"** (you need admin rights)
8. All permissions should now show green checkmarks ✓

---

## Step 4: Configure Your .env File

```bash
cd rfq-outlook-addin
cp .env.example .env
```

Edit `.env` with your values:

```env
# From Step 1
AZURE_CLIENT_ID=12345678-abcd-1234-efgh-123456789abc
AZURE_TENANT_ID=87654321-dcba-4321-hgfe-cba987654321

# From Step 2
AZURE_CLIENT_SECRET=your-secret-value-here

# Your email account (must be in your Azure AD tenant)
USER_EMAIL=you@yourcompany.com
EMAIL_DOMAIN=yourcompany.com

# Demo addresses (these will appear as senders)
ENGINEERING_EMAIL=engineering@yourcompany.com
DEMO_CLIENT_EMAIL=client@nrl.navy.mil
```

---

## Step 5: Test the Connection

```bash
npm install
npm start
```

You should see:
```
✅ Microsoft Graph client initialized
📧 Email service ready (Graph API mode)
```

If you see `📧 Email service ready (Demo mode)`, check your credentials.

---

## Step 6: Run the Demo Workflow

1. Open `https://localhost:3000/taskpane.html` in your browser
2. The demo RFQ will auto-load
3. Click **"Send All to Engineering"**
4. Watch your Outlook inbox - you'll see:
   - An email TO engineering (from you)
   - A reply FROM engineering (simulated, but real email)
5. Click **"Send Combined Reply to Client"**
6. Watch your inbox again:
   - An email TO client (from you)
   - A reply FROM client (with complete specs)
7. Click **"Generate Quote Documents"**
8. Download your Excel and PDF

---

## How Email Threading Works

```
Original RFQ (Message-ID: <abc123>)
  └── To Engineering (In-Reply-To: <abc123>)
        └── Engineering Reply (In-Reply-To: <def456>)
  └── To Client (In-Reply-To: <abc123>)
        └── Client Reply (In-Reply-To: <ghi789>)
```

The server uses `In-Reply-To` and `References` headers to ensure all emails appear in the same Outlook conversation thread.

---

## Troubleshooting

### "Access denied" or "Insufficient privileges"
- Make sure you clicked "Grant admin consent" in Step 3
- Verify all permissions show green checkmarks

### "No mailbox was found"
- The `USER_EMAIL` must be a real mailbox in your Azure AD tenant
- Can't use personal @outlook.com accounts with app-only auth

### Emails not threading
- Check server logs for the Message-ID values
- Outlook sometimes takes a few seconds to link threads

### "Client secret expired"
- Go back to Certificates & secrets and create a new one
- Update your `.env` file

---

## For Production

For a real deployment:
1. Use a dedicated service account for `USER_EMAIL`
2. Store secrets in Azure Key Vault, not `.env`
3. Add proper error handling and retry logic
4. Consider using Azure Functions instead of Express server
5. Add logging and monitoring

---

## Quick Reference

| Variable | Where to find it |
|----------|-----------------|
| `AZURE_CLIENT_ID` | App registration → Overview → Application (client) ID |
| `AZURE_TENANT_ID` | App registration → Overview → Directory (tenant) ID |
| `AZURE_CLIENT_SECRET` | App registration → Certificates & secrets → Client secrets |
| `USER_EMAIL` | Any mailbox in your Azure AD (must exist) |

# Send customer-1 email script – what you need

**App passwords:** Microsoft is retiring Basic Auth for SMTP (including app passwords). Many tenants already block it; full shutdown is around **April 2026**. So this script does **not** use SMTP or app passwords.

**This script uses Microsoft Graph** with **application permissions** (client credentials). No user password or app password is required.

---

## What you need

### 1. Azure AD app with Mail.Send (Application)

You can use your existing add-in app or create a separate one.

- **Azure Portal** → **Microsoft Entra ID** (or Azure Active Directory) → **App registrations** → your app (or **New registration**).
- **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.
- Add **Mail.Send**.
- Click **Grant admin consent for &lt;your org&gt;** (an admin must do this).

### 2. Client secret

- In the app: **Certificates & secrets** → **New client secret** → add a description, expiry, **Add**.
- Copy the **Value** once (it’s shown only once). This is `AZURE_CLIENT_SECRET`.

### 3. IDs

- **Application (client) ID** → Overview: **Application (client) ID** → `AZURE_CLIENT_ID`.
- **Directory (tenant) ID** → Overview: **Directory (tenant) ID** → `AZURE_TENANT_ID`.  
  You can use the tenant ID (GUID) or the tenant domain (e.g. `hexa729.onmicrosoft.com`).

### 4. .env (you set this; the script does not modify .env)

In your project root, in `.env`, set:

```env
AZURE_CLIENT_ID=your-client-id-guid
AZURE_CLIENT_SECRET=the-secret-value
AZURE_TENANT_ID=hexa729.onmicrosoft.com
TO_EMAIL=mannpatira@hexa729.onmicrosoft.com
```

Use your real tenant and recipient; `TO_EMAIL` is the inbox where you’ll open the email to start the add-in flow.

---

## Run the script

```bash
npm run send-customer1-email
```

or

```bash
node scripts/send-customer1-email.js
```

The script sends one message **from** `customer-1@hexa729.onmicrosoft.com` **to** `TO_EMAIL`. Open that message in Outlook as that recipient to use the add-in.

---

## If you get 403 / Access denied

- Ensure **Mail.Send** is added as an **Application** permission (not Delegated).
- Ensure **Admin consent** has been granted for that permission.
- The app must be allowed to send mail as users in the tenant; in some tenants, additional admin steps are required.

## If you get 404 / User not found

- Confirm `customer-1@hexa729.onmicrosoft.com` exists in the tenant.
- Confirm `AZURE_TENANT_ID` matches the tenant that contains that user (correct tenant ID or domain).

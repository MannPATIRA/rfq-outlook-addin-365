# Environment variables reference

## Vercel deployment (the add-in)

**The current add-in on Vercel does not use any environment variables.**  
The taskpane and `api/commands.js` are static/don’t read `process.env`, so the add-in works with **zero** env vars on Vercel.

You can safely **remove all old env vars** from the Vercel project if you want a clean setup. None of these are used by the deployed app right now:

- ~~PORT~~ (Vercel sets this automatically)
- ~~NODE_ENV~~
- ~~AZURE_*~~ (only used by local script below, not by the add-in)
- ~~USER_EMAIL, EMAIL_DOMAIN, ENGINEERING_EMAIL, DEMO_CLIENT_EMAIL, SALES_EMAIL~~
- ~~COMPANY_*, DEFAULT_UNIT_PRICE, CUSTOMS_DECLARATION, OFFER_NUMBER_PREFIX, NEXT_OFFER_NUMBER~~

## Local development / scripts

Used **only when you run locally** (e.g. `node server.js` or `node scripts/send-customer1-email.js`):

| Variable           | Used by                          | Required |
|--------------------|----------------------------------|----------|
| `PORT`             | server.js (optional; default 3000) | No       |
| `AZURE_CLIENT_ID`  | scripts/send-customer1-email.js  | Yes, for script |
| `AZURE_CLIENT_SECRET` | scripts/send-customer1-email.js | Yes, for script |
| `AZURE_TENANT_ID`  | scripts/send-customer1-email.js  | Yes, for script |
| `TO_EMAIL`         | scripts/send-customer1-email.js  | Optional (has default) |

Keep these in your **local** `.env` (and do not commit `.env`). You do **not** need to add them to Vercel unless you later add a serverless function that sends email or calls Graph.

## Importing .env into Vercel

Vercel does **not** support uploading or importing a `.env` file in the dashboard. You can:

1. **Add variables manually** in Vercel → Project → Settings → Environment Variables.
2. **Use the CLI** (one at a time):  
   `vercel env add AZURE_CLIENT_ID production`
3. **Bulk add via CLI** (from a file):  
   `vercel env pull .env.vercel` to pull from Vercel; to push, you’d add each from your `.env` with a small script or by hand.

For this project, since the add-in doesn’t use env on Vercel, you don’t need to import your `.env` for the deployment to work.

# Forever — GitHub/Vercel + Apps Script setup

## 1. Google Apps Script properties
In the Apps Script project, set these Script Properties:
- SHEET_ID = 1IBHeaaCZjgN6iwZ0gz-biRU9ymXl4sDMCnUlD3jXREA
- GOOGLE_CLIENT_ID = 470822665933-5g9960ptd660e6f3d8s03p3mdp50gm2c.apps.googleusercontent.com
- OWNER_EMAIL = h9302782@gmail.com

Run `setup()` once if the Sheets tabs/headers are not already present. Deploy the script as a Web App and use the `/exec` URL.

## 2. Vercel
Create/set this environment variable:
`APPS_SCRIPT_URL=<your Apps Script /exec URL>`

Deploy the repository. The public recipient URL format is:
`https://YOUR-DOMAIN/p/EXPERIENCE_ID`

## 3. Google OAuth
Add the deployed Vercel domain to the Google OAuth Client's Authorized JavaScript origins.
Do not commit OAuth secrets, service-account keys, or private credentials to GitHub.

## 4. Drive structure
The Apps Script uses the existing `Forever` Drive folder and creates one subfolder per experience. Photo/audio are optional. This MVP sends uploads through Apps Script as base64; keep each file under 12 MB.

## 5. Payments
Payment processing is not implemented in this MVP. `Orders` and revenue remain zero until a payment gateway/order writer is connected.

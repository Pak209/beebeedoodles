# BeeBee Market Manager Setup

The admin editor lives at `/admin.html`. It uses Google Sign-In and accepts only
`XOhellobeebee@gmail.com`. No password is stored in this repository.

## 1. Create the Google OAuth client

1. Open Google Cloud Console using the support Google account.
2. Create or select a project, configure the OAuth consent screen, and create a
   **Web application** OAuth client.
3. Add the live website origin and `http://127.0.0.1:4173` under **Authorized
   JavaScript origins**.
4. Copy the generated client ID.

## 2. Configure and deploy Apps Script

1. Open the Apps Script project attached to the BeeBee Orders spreadsheet.
2. Replace its code with `google-apps-script/Code.gs`.
3. In **Project Settings > Script Properties**, add `GOOGLE_CLIENT_ID` with the
   OAuth client ID from step 1.
4. Deploy a new Web App version using **Execute as me** and **Who has access:
   Anyone**. Admin writes are still protected by server-side Google-token and
   email verification.
5. Copy the `/exec` deployment URL.

## 3. Connect the website

Set `apiUrl` and `googleClientId` in `site-config.js`, deploy the website, then
visit `/admin.html` and sign in as `XOhellobeebee@gmail.com`.

Published events appear on the public website. Hidden events remain in the
manager but do not appear publicly.

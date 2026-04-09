# ARIS Approvals Dashboard

A permanent web app for your Gmail approval inbox. Deploy once to Vercel, open from any browser, no commands needed.

---

## One-time Setup (15 minutes)

### Step 1 — Google OAuth Credentials

1. Go to https://console.cloud.google.com
2. Create a new project called "ARIS Approvals"
3. Go to **APIs & Services → Library** → search "Gmail API" → Enable it
4. Go to **APIs & Services → OAuth consent screen**
   - User type: External
   - App name: ARIS Approvals
   - Support email: ronak@aris.in
   - Add scope: `https://www.googleapis.com/auth/gmail.modify`
   - Add test user: ronak@aris.in
   - Save
5. Go to **APIs & Services → Credentials → Create Credentials → OAuth Client ID**
   - Application type: Web application
   - Name: ARIS Approvals
   - Authorized redirect URIs: `https://YOUR-VERCEL-URL.vercel.app/api/gmail?action=auth-callback`
   - (You'll update this after deploying to Vercel)
   - Click Create → copy **Client ID** and **Client Secret**

### Step 2 — Deploy to Vercel

1. Install Vercel CLI: `npm i -g vercel`
2. In this folder run: `vercel`
   - Follow prompts, link to your account
   - Note your Vercel URL (e.g. `aris-approvals.vercel.app`)
3. Go back to Google Cloud Console → update the redirect URI to:
   `https://aris-approvals.vercel.app/api/gmail?action=auth-callback`

### Step 3 — Set Environment Variables in Vercel

In Vercel dashboard → your project → Settings → Environment Variables, add:

| Variable | Value |
|----------|-------|
| `GMAIL_CLIENT_ID` | Your Google Client ID |
| `GMAIL_CLIENT_SECRET` | Your Google Client Secret |
| `GMAIL_REDIRECT_URI` | `https://YOUR-VERCEL-URL.vercel.app/api/gmail?action=auth-callback` |

Then redeploy: `vercel --prod`

### Step 4 — Sign in

Open your Vercel URL → click "Sign in with Google" → choose ronak@aris.in → done.

Your tokens are saved in the browser. You'll never need to sign in again unless you clear browser data.

---

## Daily Use

Just open `https://aris-approvals.vercel.app` — the dashboard loads your live inbox instantly.

- **Approve** → sends "Approved." reply-all immediately
- **Reject** → sends "Rejected." reply-all immediately  
- **Mark as read** → removes FYI items (you were CC only)
- **Refresh** → re-fetches live from Gmail

---

## Refresh Data

The dashboard fetches live from Gmail every time you click Refresh or open the page. No stale data, no manual updates needed.

# Outlook Organizer

Fetches your Outlook inbox through the Microsoft Graph API and sorts recent
messages into projects — either with keyword rules or with AI classification.
Runs entirely on your machine; your email never passes through anyone else's server.

## What it does

- Signs you in with your own Microsoft account (OAuth, tokens cached locally).
- Pulls your most recent inbox messages.
- Assigns each to a project via `projects.json` rules or an AI call.
- Prints the grouped result, or shows it in a browser dashboard.
- Optionally writes the project name back as an Outlook **category** so the
  grouping shows up in Outlook itself.

## 1. Register an Azure app (one time, ~5 min)

You need a client ID so Microsoft will let the app read your mail.

1. Go to https://portal.azure.com → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it anything (e.g. "My Outlook Organizer").
3. Supported account types:
   - Personal Outlook.com / Hotmail → "Personal Microsoft accounts only" (or the combined option).
   - Work/school account → your organization.
4. **Redirect URI**: platform **Web**, value `http://localhost:3000/callback`.
5. Click **Register**. Copy the **Application (client) ID** and **Directory (tenant) ID**.
6. Under **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions** → add `Mail.Read`, `Mail.ReadWrite`, `User.Read`.
   (Delegated = acts as you; no admin consent needed for personal accounts.)

No client secret is required — this uses the public-client auth-code flow.

## 2. Configure

```bash
cp .env.example .env
```

Edit `.env` and set `CLIENT_ID` and `TENANT_ID`. For a personal account,
`TENANT_ID=common` is fine. Leave `APPLY_CATEGORIES=false` for a dry run.

## 3. Define your projects

Edit `projects.json`. Each project matches on `keywords` (found in subject or
preview) and `senders` (substring of the from-address). Highest score wins;
anything unmatched goes to `fallback`.

## 4. Install & run

```bash
npm install
```

CLI (prints grouped inbox to the terminal):

```bash
npm run categorize
```

Browser dashboard:

```bash
npm run server
# open http://localhost:3000 and click "Fetch & sort"
```

First run opens a browser to sign in. The token is cached in
`.token-cache.json` so later runs are silent until it expires.

## AI categorization (optional)

Set in `.env`:

```
CATEGORIZE_MODE=ai
ANTHROPIC_API_KEY=sk-ant-...
```

Each email is sent (from address, subject, preview only) to Claude Haiku, which
returns one of your project names. Costs a fraction of a cent per email. Use
`rules` mode if you'd rather keep everything local and free.

## Writing categories back to Outlook

Set `APPLY_CATEGORIES=true`. The project name is added to each message's Outlook
categories, so you can filter/group by it in Outlook. This is the only action
that modifies your mailbox; it never deletes, moves, or sends anything.

## Files

```
src/auth.js        MSAL sign-in + token cache
src/graph.js       fetch inbox, apply categories
src/categorize.js  rules + AI classifiers
src/index.js       CLI
src/server.js      dashboard API
public/index.html  dashboard UI
projects.json      your project rules
.env               your config (not committed)
```

## Notes

- `.token-cache.json` holds a live login — treat it like a password and don't
  commit it. A `.gitignore` is included.
- To reset auth, delete `.token-cache.json`.
- Tokens are delegated and scoped to the permissions above; the app can't do
  more than read your mail and set categories.

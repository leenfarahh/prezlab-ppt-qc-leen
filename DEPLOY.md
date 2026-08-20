# Deploying the demo/eval instance (Render + Supabase)

This is the **demo/eval** deployment: a public URL for showing the tool and
gathering feedback. It is **not for confidential client decks**. Client work
stays on the Windows LAN box, which renders with real PowerPoint and keeps
decks on-machine. This cloud instance renders with LibreOffice (lower Arabic
fidelity) and processes uploads on Render's servers, so treat everything you
upload here as non-confidential.

## What persists where

- **Supabase (Postgres):** users, sign-in sessions, audit history, comments,
  triage judgments. Survives restarts and redeploys.
- **Not persisted on the cloud instance:** profile edits made in the admin UI
  (the container filesystem resets on redeploy; seeded profiles ship in the
  image). Uploaded decks are never stored anywhere, on purpose.

## One-time setup

**1. Supabase**
- Create a project (choose the region closest to the UAE).
- Project settings, Database, copy the connection string (URI form). Use the
  connection **pooler** URI for a web app.
- No manual SQL needed: the app creates its tables on first boot. The same
  schema is in `supabase/schema.sql` if you prefer to run it yourself.

**2. Render**
- Push this repo to GitHub (see "First push" below).
- In Render: New, Blueprint, point it at the repo. It reads `render.yaml`.
- Set the two secret env vars in the dashboard:
  - `DATABASE_URL` = the Supabase connection string
  - `QC_BOOTSTRAP_ADMIN` = the first admin's name, e.g. `Sanad`
- Deploy. First boot builds the Docker image (LibreOffice makes this a few
  minutes), creates the tables, and creates the bootstrap admin.

**3. First sign-in**
- Open the Render URL, go to Sign in, pick the bootstrap admin name, and set a
  PIN (first sign-in sets it).
- Add the rest of the team on the Team page; each sets their own PIN on first
  sign-in.

## Environment variables

| Var | Purpose | Demo value |
|-----|---------|------------|
| `DATABASE_URL` | Supabase Postgres connection string | (secret) |
| `QC_AUTH_REQUIRED` | Require sign-in on every route | `1` |
| `QC_STRICT_SESSIONS` | Ignore the legacy name cookie | `1` |
| `QC_RENDERER` | Slide renderer | `libreoffice` |
| `QC_DEMO_BANNER` | Warning strip text | `DEMO INSTANCE - do not upload confidential client decks` |
| `QC_BOOTSTRAP_ADMIN` | First admin created on empty DB | e.g. `Sanad` |
| `QC_MAX_UPLOAD_MB` | Upload size cap in MB | `60` on the free tier (512 MB RAM); default `250` suits the LAN box |

Leave `DATABASE_URL` unset locally to keep using SQLite; leave `QC_RENDERER`
unset (or `auto`) on the Windows box to use PowerPoint.

## First push

The repo must be on GitHub for Render to deploy it:

```
git remote add origin <your-github-repo-url>
git push -u origin main
```

## Free plan note

The `free` plan sleeps after inactivity, so the first hit after idle takes ~30s
to wake. Switch `plan: free` to `plan: starter` in `render.yaml` for an
always-on demo.

## Not on this instance (by design)

Real client decks, Microsoft 365 SSO, and full-fidelity (PowerPoint / Graph)
rendering all belong to the production tier and remain gated on the
IT-security data-residency ruling per the PRD.

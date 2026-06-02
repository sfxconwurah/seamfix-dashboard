# Onboarding Guide: Seamfix Financial Dashboard

> **For**: Finance team members who will maintain and update the dashboard using Claude  
> **Last updated**: June 2026

---

## What This Project Is

The Seamfix Financial Intelligence Suite is a set of 5 interactive dashboards that show the company's financial health — revenue tracking, cash position, expenses, budget vs actual, and pipeline intelligence. It also includes "Bobby," an AI chat assistant that can answer questions about the data.

**Live dashboard**: https://seamfix-executive-dashboard.streamlit.app  
**GitHub repo**: https://github.com/sfxconwurah/seamfix-dashboard (branch: `main`)  
**Full technical docs**: See `CLAUDE.md` in this same folder

> **Note**: The repo currently lives under Chibuzor's personal GitHub account
> (`sfxconwurah/seamfix-dashboard`). A migration to the company org repo
> (`seamfix/finance-dashboard`) is planned but not done yet — it's waiting on
> CTO access. Until then, use the personal repo above, and ask Chibuzor to add
> you as a collaborator.

---

## Prerequisites

Before you start, you need three things:

1. **A GitHub account** with access to the `sfxconwurah/seamfix-dashboard` repository. Ask Chibuzor to add you as a collaborator.

2. **Claude Pro or Team plan** at https://claude.ai. Seamfix has a team plan — ask your admin to add your email.

3. **Claude Desktop app** installed on your computer. Download from https://claude.ai/download. You need the desktop app (not just the website) because it includes "Cowork mode" which can edit files and run code.

---

## Initial Setup (One-Time)

### Step 1: Install Claude Desktop

Download and install from https://claude.ai/download. Sign in with your Seamfix Claude account.

### Step 2: Enable Cowork Mode

In Claude Desktop, look for "Cowork" in the sidebar or settings. Enable it. This gives Claude the ability to read/write files and run commands on your computer.

### Step 3: Clone the Repository

Open a terminal (Mac: Terminal app, Windows: PowerShell) and run:

```bash
git clone git@github.com:sfxconwurah/seamfix-dashboard.git
cd seamfix-dashboard
```

If you get a permission error, you need to set up an SSH key. See the "SSH Key Setup" section below.

### Step 4: Connect Claude to the Project Folder

In Claude Desktop's Cowork mode, click "Select folder" and choose the `seamfix-dashboard` folder you just cloned. Claude will automatically read `CLAUDE.md` to understand the project.

---

## SSH Key Setup (If Needed)

If `git clone` fails with a permission error, you need an SSH key:

1. Open terminal and run:
   ```bash
   ssh-keygen -t ed25519 -C "your-email@seamfix.com"
   ```
   Press Enter to accept defaults. You can set a passphrase or leave it empty.

2. Copy the public key:
   - **Mac**: `pbcopy < ~/.ssh/id_ed25519.pub`
   - **Windows**: `cat ~/.ssh/id_ed25519.pub | clip`

3. Go to https://github.com/settings/keys → "New SSH key" → paste the key → Save

4. Test: `ssh -T git@github.com` — you should see "Hi username! You've successfully authenticated"

---

## How to Make Changes Using Claude

This is the core workflow. You talk to Claude in natural language, and it edits the code for you.

### Example 1: Fix a Bug

You notice the dashboard shows wrong numbers. Open Claude Cowork with the project folder selected and say:

> "The achievement percentage on the Revenue dashboard is showing 200% for deals that should be 100%. Can you investigate and fix this?"

Claude will:
1. Read the relevant code files
2. Identify the bug
3. Fix it
4. Test the fix locally
5. Tell you what it changed

### Example 2: Add a Feature

> "Add a new column to the Revenue Streams table that shows the monthly average revenue for each deal."

### Example 3: Update Business Logic

> "Change the FX rate from 1450 to 1500 across all dashboards."

### Key Tips

- **Be specific** about which dashboard has the issue (Cash Overview, Expense & Vendor, Budget vs Actual, Revenue & Fundability, or Pipeline Intelligence)
- **Share screenshots** if you see something wrong — you can paste them directly into Claude
- **Always test** before pushing. Ask Claude: "Can you test this locally to make sure it works?"
- **Quick local preview without Claude**: double-click `UPDATE_DASHBOARD.command` in the project folder. It regenerates all 5 dashboards from the Excel files in `data/` and opens them in your browser. (This is a local preview only — it does not change the live site. The live site updates when changes are pushed to GitHub.)

---

## How to Push Changes to GitHub

After Claude makes changes, you need to push them to GitHub so the live dashboard updates.

### Option A: Ask Claude to Push (Recommended)

Simply say:

> "Please commit and push these changes to GitHub."

Claude will handle the git commands for you. It will ask you to confirm before pushing.

### Option B: Push Manually

In terminal, from the project folder:

```bash
git add .
git commit -m "Description of what changed"
git push origin main
```

---

## How to Deploy Changes

After pushing to GitHub:

1. Go to https://share.streamlit.io
2. Find the Seamfix dashboard app
3. Click the three dots menu next to it
4. Click **Reboot**
5. Wait 30-60 seconds for the app to restart

**Important**: Just clicking "Regenerate Dashboards" inside the app only refreshes the data — it does NOT pick up code changes. You must reboot from share.streamlit.io for code changes to take effect.

---

## Common Tasks

### "The numbers look wrong on a dashboard"

Tell Claude which dashboard, what's wrong, and what you expect. Example:

> "On the Revenue & Fundability dashboard, the deal 'NIMC NIN Tokenization' shows 0% achievement but it should show some progress. It has revenue in the May column of the Google Sheet."

### "Finance added new deals to the Google Sheet"

No code change needed. Just click "Regenerate Dashboards" in the app sidebar (under the gear icon), or wait for the 24-hour auto-refresh.

### "Finance added a new month of data"

No code change needed. The dashboards automatically detect which months have data. Click "Regenerate Dashboards" or wait for auto-refresh.

### "We need to add a new user to the dashboard"

This requires updating Streamlit Cloud secrets (not code). You need access to share.streamlit.io:
1. Go to Settings → Secrets
2. Find the `allowed_emails` list under `[google_oauth]`
3. Add the new email address
4. Save

### "The FX rate changed significantly"

Tell Claude:

> "Update the FX rate from 1450 to [new rate] across all dashboard generators."

Then commit, push, and reboot.

### "We need to change the annual revenue target"

Tell Claude:

> "Update the annual revenue target from $8M to $[new amount]. It's used in the pipeline dashboard as LANDING_ZONE and in the revenue dashboard as annual_revenue_target_usd."

---

## Key Files (What Does What)

| File | Purpose |
|------|---------|
| `app.py` | Main application — handles login, data fetching, tabs, Bobby chat |
| `generate_dashboard.py` | Generates the Cash Overview dashboard |
| `generate_expense_dashboard.py` | Generates the Expense & Vendor Analysis dashboard |
| `generate_budget_dashboard.py` | Generates the Budget vs Actual dashboard |
| `generate_revenue_dashboard.py` | Generates the Revenue & Fundability dashboard |
| `generate_pipeline_dashboard.py` | Generates the Pipeline Intelligence dashboard |
| `CLAUDE.md` | Full technical documentation (Claude reads this automatically) |
| `requirements.txt` | Python package dependencies |
| `data/` | Bundled Excel files (fallback if Google Drive/Sheets unavailable) |

---

## Access & Permissions Summary

| Resource | Who Manages | How to Access |
|----------|-------------|---------------|
| Dashboard login | Streamlit secrets (`allowed_emails`) | Ask dashboard admin |
| GitHub repo | Seamfix GitHub org admin | Ask IT or Chibuzor |
| Streamlit Cloud | App owner | share.streamlit.io |
| Google Sheet (revenue data) | Finance team | Direct edit — auto-fetched by dashboard |
| Google Drive (cash reports) | Finance team | Upload to shared folder — auto-fetched |
| Bobby chat (Claude API) | Streamlit secrets (`ANTHROPIC_API_KEY`) | Ask dashboard admin |

---

## Getting Help

- **Dashboard bugs or feature requests**: Use Claude Cowork with the project folder
- **Can't log in**: Check if your email is in the `allowed_emails` list
- **Data questions**: Ask Bobby (the chat assistant in the dashboard sidebar)
- **Infrastructure issues**: Contact Chibuzor (conwurah@gmail.com)

---

## Important Reminders

1. **Always test locally before pushing.** Ask Claude to run the generator scripts against the data folder.
2. **Reboot after pushing code changes.** "Regenerate Dashboards" only refreshes data, not code.
3. **Don't modify the Google Sheet structure** (column order, sheet names) without coordinating — the code depends on specific column positions.
4. **Read CLAUDE.md** for deep technical details — it covers column mappings, business logic, and every gotcha we've learned.

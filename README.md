# Timesheet from git + Outlook to Jira, using a Claude Code Cloud session

A Cloud session can't reach on-prem GitLab or your laptop. It doesn't need to:
a timesheet needs your **git logs**, not the code. So the laptop exports the logs
and the calendar with a plain script (no Claude involved), uploads that small
folder to a Cloud session, and Claude does the analysis and drafting there.

```
laptop (no Claude usage)                 Cloud session (cloud credit)
───────────────────────                  ────────────────────────────
timesheet.py collect                     reads timeline.md, git-activity.json,
  git log + reflog of every checkout       calendar.csv; asks your rules;
  Outlook calendar (COM)            ──►    drafts weekly tables with you;
  → ~/timesheet-YYYY-MM (git repo,         writes worklogs.json
     no remote)                                       │
claude --cloud "..."  (uploads it)                    ▼
                                         posts to Jira (if allowed) or prints
timesheet.py post worklogs.json --apply ◄─ worklogs.json for you to post
```

## One-time setup (Windows)

- Python 3.9+ (`py --version`), Git, classic Outlook desktop.
- Claude Code CLI, signed in to the account that has the Cloud credit:
  ```powershell
  irm https://claude.ai/install.ps1 | iex
  claude auth login
  ```
  The CLI is only used to *start* the Cloud session; the work runs in the cloud.
- Get this folder onto the laptop (clone or download the zip).

## Each month

Use a **normal (not Administrator) PowerShell**: Outlook runs un-elevated and
an elevated shell can't connect to it (error `0x80080005`).

```powershell
py timesheet.py collect --month 2026-09 --root C:\dev --root D:\work --outlook
```

- `--root`: folders that contain your checkouts, searched 4 levels deep.
- `--author`: your git email/name; defaults to your global `git config`. Repeat it
  if you've committed under more than one address.
- Optional: `--notes notes.md` with leave days, default tickets, meeting
  mappings; `--include-body` to add the first 300 characters of meeting bodies
  (useful if Jira keys are in invitations).

Read `~\timesheet-2026-09\timeline.md` and remove anything you don't want to
send, then:

```powershell
cd ~\timesheet-2026-09
claude --cloud "Prepare my 2026-09 timesheet following CLAUDE.md"
```

Because that folder is a git repo with **no remote**, Claude Code uploads it as
a bundle instead of cloning. Open the session in the Desktop app (or
claude.ai/code / mobile), answer Claude's questions, review the weekly tables.

## Posting to Jira

`timesheet.py post` validates the file, shows totals per day and issue, checks
Jira for worklogs you already logged (same issue, start time and duration are
skipped), and posts only with `--apply`.

**On the laptop (recommended: the token never leaves it).** Copy the final
`worklogs.json` from the session, then:

```powershell
$env:JIRA_URL = "https://it-projects.just.fgov.be"
$env:JIRA_PAT = "<personal access token>"
py timesheet.py post worklogs.json --month 2026-09          # dry run + duplicate check
py timesheet.py post worklogs.json --month 2026-09 --apply  # post
```

**From the Cloud session.** Open the environment editor: in the Desktop app
with **Cloud** selected (or at claude.ai/code), click the cloud/environment
button above the message box, hover your environment (e.g. *Default*) and click
the gear icon. In **Edit cloud environment**:

1. Environment variables: add `JIRA_URL=https://it-projects.just.fgov.be`.
2. The token, one of:
   - **API credentials** section (Pro and Max plans only): *Add credential* →
     type *Bearer*, host `it-projects.just.fgov.be`, header `Authorization`,
     prefix `Bearer`, value = your PAT. Claude never sees the token, and the host
     is allowed automatically. Claude then posts with `--proxy-auth`.
   - Otherwise an environment variable `JIRA_PAT=...` (readable inside the
     session: give it a short expiry and revoke it afterwards), plus
     **Network access** → *Custom* → `it-projects.just.fgov.be` in **Allowed
     domains**, with *Also include default list of common package managers* ticked.

Changes apply to new sessions, so edit the environment before `claude --cloud`.

Credentials, depending on the Jira edition (open
`https://it-projects.just.fgov.be/rest/api/2/serverInfo` and look at
`deploymentType`):

- `Server` (this includes Data Center): a Personal Access Token from your avatar →
  *Profile* → *Personal Access Tokens* → *Create token* (set an expiry).
- `Cloud`: `JIRA_EMAIL` + `JIRA_API_TOKEN` from id.atlassian.com → Security → API tokens.

Worklogs go through the standard Jira REST API (`/rest/api/2/issue/{key}/worklog`).
If your organisation uses Tempo Timesheets, these worklogs normally show up there too.

## Calendar fallbacks

`Export-OutlookCalendar.ps1` needs **classic** Outlook (the "new Outlook" has no
COM interface) and PowerShell scripts being allowed. If either blocks you,
export an `.ics` file instead (classic Outlook: *File → Save Calendar*, choose a
date range and *Full details*) and pass it with `--calendar file.ics`; Claude
expands recurring meetings in the Cloud session.

## What gets uploaded

Commit subjects/bodies, branch names, repository names, file-change counts and
meeting subjects/times — no source code. That still is work data: check that
your employer allows sending it to the Claude account you use.

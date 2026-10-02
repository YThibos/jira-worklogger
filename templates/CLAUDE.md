# Timesheet session

This folder is a one-month activity export made by `timesheet.py collect` on the
user's work laptop. Your job is to turn it into Jira worklogs the user approves.

## Inputs

- `timeline.md`: day-by-day merge of commits, branch switches and meetings. Start here.
- `git-activity.json`: full detail per repository (commit messages, file stats,
  Jira keys found in messages and branch names, HEAD reflog).
- `calendar.csv` (Outlook export) and/or `*.ics`: meetings. Ignore items with
  `Response` = `Declined` and `ShowAs` = `Free`. An `OutOfOffice` all-day item is leave.
  If there is an `.ics` file, expand recurring events (RRULE) before using it.
- `notes.md`, if present: the user's own notes (leave, defaults, mappings). These win.

## Before allocating, ask the user (one message, skip what notes.md answers)

1. Hours per working day (Belgian federal default: 7h36, 38h/week) and whether
   to fill every working day to that total or log only evidenced time.
2. Leave, sick days, public holidays and training days that month.
3. Which Jira issue takes meetings, admin and other work without a ticket.
4. How to map recurring meetings (stand-up, refinement, ...) to issues.
5. Rounding: 15 or 30 minutes.

## Allocation rules

- Meetings: their real duration, on the issue whose key is in the subject or on
  the user's mapping, at their real start time.
- Development time: split the rest of the day across the issues with evidence
  that day (commits, branch checkouts), weighted by activity. Work on a branch
  usually starts at its checkout and ends at the last commit or the next checkout.
- Keys come from branch names (`feature/ABC-123-...`), commit messages and
  meeting subjects. When a repository has no keys, ask which issue it belongs to.
- Never invent work for days without evidence: list them and ask.
- Give each worklog a short comment summarising the commits or the meeting.

## Review, then write worklogs.json

Show the draft as one table per week (day, issue, hours, comment) with daily
totals, and iterate until the user approves it. Then write `worklogs.json`:

```json
{
  "timezone": "Europe/Brussels",
  "worklogs": [
    {"issue": "ABC-123", "date": "2026-09-01", "start": "09:00", "minutes": 90,
     "comment": "Refactor login flow, fix session timeout"}
  ]
}
```

Validate it with `python3 timesheet.py post worklogs.json --month YYYY-MM` (offline dry run).

## Posting

- If `JIRA_URL` and a credential (`JIRA_PAT`, or `JIRA_EMAIL` + `JIRA_API_TOKEN`)
  are set and the Jira host is reachable, run the same command to check for
  duplicates against Jira. Post with `--apply` only after the user explicitly says so.
- Otherwise print the final `worklogs.json` in a single code block so the user can
  save it on the laptop and run `py timesheet.py post worklogs.json --apply` there.
  This session was uploaded as a bundle and cannot push files back.

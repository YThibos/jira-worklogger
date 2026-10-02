#!/usr/bin/env python3
"""Monthly timesheet helper: collect git + calendar activity, post Jira worklogs.

Standard library only (Python 3.9+), works on Windows, macOS and Linux.

  collect  Run on the laptop. Scans local git checkouts (any host: GitLab,
           GitHub, ...) for your commits and branch switches in a month, adds an
           Outlook calendar export, and writes everything into a fresh folder
           that is its own git repo with no remote. Running `claude --cloud`
           from that folder uploads it to a Cloud session as a bundle.

  post     Validate a worklogs.json file and, with --apply, post each entry to
           Jira as a worklog. Without --apply nothing is written to Jira.
"""

import argparse
import base64
import calendar
import csv
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
JIRA_KEY = re.compile(r"\b[A-Z][A-Z0-9_]{1,15}-\d+\b")
SKIP_DIRS = {"node_modules", ".venv", "venv", "target", "build", "dist", ".gradle", ".idea", "__pycache__"}
US, RS, GS = "\x1f", "\x1e", "\x1d"


def utf8_stdout():
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def month_range(month):
    year, mon = (int(p) for p in month.split("-"))
    first = dt.date(year, mon, 1)
    last = dt.date(year, mon, calendar.monthrange(year, mon)[1])
    return first, last


def run_git(repo, *args):
    cmd = ["git", "-C", str(repo), "-c", "i18n.logOutputEncoding=utf-8", *args]
    res = subprocess.run(cmd, capture_output=True, encoding="utf-8", errors="replace")
    if res.returncode != 0:
        raise RuntimeError(f"{' '.join(cmd)}: {res.stderr.strip()}")
    return res.stdout


def find_repos(roots, max_depth):
    repos = []
    for root in roots:
        root = Path(root).expanduser().resolve()
        if not root.is_dir():
            print(f"warning: {root} is not a directory, skipped", file=sys.stderr)
            continue
        base_depth = len(root.parts)
        for dirpath, dirnames, _ in os.walk(root):
            path = Path(dirpath)
            if (path / ".git").exists():
                repos.append(path)
                dirnames[:] = []  # don't descend into a checkout (submodules included)
                continue
            if len(path.parts) - base_depth >= max_depth:
                dirnames[:] = []
                continue
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith(".")]
    return sorted(set(repos))


def default_authors():
    authors = []
    for key in ("user.email", "user.name"):
        res = subprocess.run(["git", "config", "--global", key], capture_output=True, encoding="utf-8")
        if res.returncode == 0 and res.stdout.strip():
            authors.append(res.stdout.strip())
    return authors


def in_range(iso, first, last):
    day = dt.datetime.fromisoformat(iso).date()
    return first <= day <= last


def jira_keys(*texts):
    return sorted({k for t in texts if t for k in JIRA_KEY.findall(t)})


def collect_commits(repo, authors, first, last):
    # --since/--until filter on committer date; widen the window so commits that
    # were rebased or merged later still show up, then keep those authored in range.
    since = (first - dt.timedelta(days=7)).isoformat()
    until = (last + dt.timedelta(days=8)).isoformat()
    fmt = RS + US.join(["%H", "%an", "%ae", "%aI", "%cI", "%S", "%D", "%P", "%s"]) + US + "%b" + GS
    # --source fills %S with the ref each commit was reached from, usually its branch.
    args = ["log", "--all", "--source", f"--since={since}", f"--until={until}", f"--format={fmt}", "--numstat"]
    args += [f"--author={a}" for a in authors]
    out = run_git(repo, *args)
    commits = []
    for record in out.split(RS)[1:]:
        head, _, stats = record.partition(GS)
        fields = head.split(US)
        if len(fields) < 10:
            continue
        sha, name, email, adate, cdate, source, refs, parents, subject, body = fields[:10]
        if not in_range(adate, first, last):
            continue
        added = deleted = files = 0
        for line in stats.strip().splitlines():
            parts = line.split("\t")
            if len(parts) == 3:
                files += 1
                added += int(parts[0]) if parts[0].isdigit() else 0
                deleted += int(parts[1]) if parts[1].isdigit() else 0
        commits.append({
            "sha": sha[:12],
            "author": f"{name} <{email}>",
            "authored": adate,
            "committed": cdate,
            "branch": re.sub(r"^refs/(heads|remotes)/", "", source),
            "refs": refs,
            "merge": len(parents.split()) > 1,
            "subject": subject,
            "body": body.strip(),
            "files": files,
            "added": added,
            "deleted": deleted,
            "jira": jira_keys(subject, body, source, refs),
        })
    return commits


def collect_reflog(repo, first, last):
    # HEAD reflog is local-only history of checkouts, commits, rebases and pulls:
    # it shows which branch you were on and when, even on days without commits.
    try:
        out = run_git(repo, "reflog", "show", "--date=iso-strict", f"--format=%gd{US}%gs{US}%h", "HEAD")
    except RuntimeError:
        return []
    entries = []
    for line in out.splitlines():
        parts = line.split(US)
        if len(parts) != 3:
            continue
        selector, message, sha = parts
        m = re.search(r"@\{(.+)\}", selector)
        if not m:
            continue
        try:
            when = dt.datetime.fromisoformat(m.group(1))
        except ValueError:
            continue
        if first <= when.date() <= last:
            entries.append({"time": when.isoformat(), "action": message, "sha": sha, "jira": jira_keys(message)})
    entries.reverse()
    return entries


def remote_name(repo):
    try:
        url = run_git(repo, "remote", "get-url", "origin").strip()
    except RuntimeError:
        return repo.name
    parts = [p for p in re.split(r"[:/]", re.sub(r"\.git$", "", url.rstrip("/"))) if p]
    return "/".join(parts[-2:]) if len(parts) >= 2 else repo.name


def export_outlook(first, last, out_csv, include_body):
    script = HERE / "Export-OutlookCalendar.ps1"
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
           "-Start", first.isoformat(), "-End", (last + dt.timedelta(days=1)).isoformat(),
           "-OutFile", str(out_csv)]
    if include_body:
        cmd.append("-IncludeBody")
    res = subprocess.run(cmd, capture_output=True, encoding="utf-8", errors="replace")
    if res.returncode != 0 or not out_csv.exists():
        raise RuntimeError(f"Outlook export failed: {res.stderr.strip() or res.stdout.strip()}")


def read_calendar_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def write_timeline(path, month, repos, events):
    days = defaultdict(list)
    for repo in repos:
        for c in repo["commits"]:
            t = dt.datetime.fromisoformat(c["authored"])
            stats = f"{c['files']} files +{c['added']}/-{c['deleted']}"
            days[t.date()].append((t.strftime("%H:%M"),
                                   f"commit  [{repo['name']} {c['branch']}] {c['subject']} ({stats})"))
        for r in repo["reflog"]:
            if r["action"].startswith(("checkout:", "merge", "rebase", "pull", "reset", "cherry-pick")):
                t = dt.datetime.fromisoformat(r["time"])
                days[t.date()].append((t.strftime("%H:%M"), f"git     [{repo['name']}] {r['action']}"))
    for e in events:
        try:
            start = dt.datetime.fromisoformat(e["Start"])
            end = dt.datetime.fromisoformat(e["End"])
        except (KeyError, ValueError):
            continue
        flags = []
        if e.get("AllDay") == "True":
            flags.append("all day")
        if e.get("Response") == "Declined":
            flags.append("DECLINED")
        if e.get("ShowAs") not in (None, "", "Busy"):
            flags.append(e["ShowAs"])
        extra = f" ({', '.join(flags)})" if flags else ""
        days[start.date()].append((start.strftime("%H:%M"),
                                   f"meeting {start:%H:%M}-{end:%H:%M} {e.get('Subject', '')}{extra}"))
    lines = [f"# Activity timeline {month}", "",
             "Generated by `timesheet.py collect`. Times are local to the laptop.", ""]
    for day in sorted(days):
        lines.append(f"## {day:%a %Y-%m-%d}")
        lines += [f"- {t} {text}" for t, text in sorted(days[day])]
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def cmd_collect(args):
    first, last = month_range(args.month)
    authors = args.author or default_authors()
    if not authors:
        sys.exit("No --author given and no global git user.email/user.name configured.")
    out = Path(args.out or Path.home() / f"timesheet-{args.month}").expanduser().resolve()
    if out.exists() and any(out.iterdir()) and not args.force:
        sys.exit(f"{out} already exists and is not empty; pass --force to overwrite.")
    out.mkdir(parents=True, exist_ok=True)

    print(f"Scanning {', '.join(args.root)} for git checkouts...")
    repos = []
    for path in find_repos(args.root, args.max_depth):
        try:
            commits = collect_commits(path, authors, first, last)
            reflog = collect_reflog(path, first, last)
        except RuntimeError as exc:
            print(f"  skipped {path}: {exc}", file=sys.stderr)
            continue
        if commits or reflog:
            repos.append({"name": remote_name(path), "path": str(path), "commits": commits, "reflog": reflog})
            print(f"  {path}: {len(commits)} commits, {len(reflog)} reflog entries")
    activity = {"month": args.month, "authors": authors, "generated": dt.datetime.now().astimezone().isoformat(),
                "repos": repos}
    (out / "git-activity.json").write_text(json.dumps(activity, indent=2, ensure_ascii=False), encoding="utf-8")

    events = []
    if args.outlook:
        print("Exporting Outlook calendar...")
        export_outlook(first, last, out / "calendar.csv", args.include_body)
        events = read_calendar_csv(out / "calendar.csv")
        print(f"  {len(events)} calendar items")
    for cal in args.calendar or []:
        src = Path(cal).expanduser()
        dest = out / src.name
        shutil.copyfile(src, dest)
        if src.suffix.lower() == ".csv":
            events += read_calendar_csv(dest)
        print(f"  copied {src.name}")

    write_timeline(out / "timeline.md", args.month, repos, events)
    shutil.copyfile(HERE / "timesheet.py", out / "timesheet.py")
    shutil.copyfile(HERE / "templates" / "CLAUDE.md", out / "CLAUDE.md")
    if args.notes:
        shutil.copyfile(Path(args.notes).expanduser(), out / "notes.md")

    # A repo with no remote is what makes `claude --cloud` upload a bundle instead of cloning.
    if not (out / ".git").exists():
        subprocess.run(["git", "-C", str(out), "init", "-q"], check=True)
    subprocess.run(["git", "-C", str(out), "add", "-A"], check=True)
    subprocess.run(["git", "-C", str(out), "-c", "user.name=timesheet", "-c", "user.email=timesheet@localhost",
                    "-c", "commit.gpgsign=false",
                    "commit", "-q", "--no-verify", "--allow-empty", "-m", f"Activity export {args.month}"], check=True)

    total = sum(len(r["commits"]) for r in repos)
    print(f"\nWrote {out}: {len(repos)} repos, {total} commits, {len(events)} calendar items.")
    print("Review timeline.md, then start the Cloud session from that folder:\n")
    print(f'  cd "{out}"')
    print(f'  claude --cloud "Prepare my {args.month} timesheet following CLAUDE.md"')


# ---------------------------------------------------------------- post


def tzinfo_for(name):
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo(name)
    except Exception:
        return None  # Windows without the tzdata package: fall back to the machine's local zone


def jira_started(date, start, tz):
    naive = dt.datetime.fromisoformat(f"{date}T{start}")
    aware = naive.replace(tzinfo=tz) if tz else naive.astimezone()
    return aware.strftime("%Y-%m-%dT%H:%M:%S.000%z")


def parse_started(value):
    # Jira's "2026-09-01T09:00:00.000+0200"; compared as instants so offsets don't matter.
    return dt.datetime.strptime(value, "%Y-%m-%dT%H:%M:%S.%f%z").astimezone(dt.timezone.utc)


class Jira:
    def __init__(self, base, pat=None, email=None, token=None):
        self.base = base.rstrip("/")
        if pat:
            self.auth = f"Bearer {pat}"
        elif email and token:
            self.auth = "Basic " + base64.b64encode(f"{email}:{token}".encode()).decode()
        else:
            raise SystemExit("Set JIRA_PAT (Server/Data Center) or JIRA_EMAIL + JIRA_API_TOKEN (Cloud).")

    def call(self, method, path, body=None):
        req = urllib.request.Request(self.base + path, method=method,
                                     data=json.dumps(body).encode() if body is not None else None,
                                     headers={"Authorization": self.auth, "Accept": "application/json",
                                              "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "replace")[:300]
            raise RuntimeError(f"{method} {path}: HTTP {exc.code} {detail}") from None
        return json.loads(raw) if raw else None


def load_worklogs(path, month):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    entries = data["worklogs"] if isinstance(data, dict) else data
    tz_name = data.get("timezone", "Europe/Brussels") if isinstance(data, dict) else "Europe/Brussels"
    first, last = month_range(month) if month else (None, None)
    errors = []
    for i, e in enumerate(entries, 1):
        where = f"entry {i} ({e.get('issue')} {e.get('date')})"
        if not JIRA_KEY.fullmatch(str(e.get("issue", ""))):
            errors.append(f"{where}: bad issue key")
        try:
            day = dt.date.fromisoformat(e["date"])
            dt.time.fromisoformat(e.get("start", "09:00"))
            if first and not first <= day <= last:
                errors.append(f"{where}: date outside {month}")
        except (KeyError, ValueError):
            errors.append(f"{where}: date must be YYYY-MM-DD and start HH:MM")
        if not isinstance(e.get("minutes"), int) or not 1 <= e["minutes"] <= 24 * 60:
            errors.append(f"{where}: minutes must be a whole number between 1 and 1440")
    return entries, tz_name, errors


def print_summary(entries):
    per_day, per_issue = defaultdict(int), defaultdict(int)
    for e in entries:
        per_day[e["date"]] += e["minutes"]
        per_issue[e["issue"]] += e["minutes"]
    fmt = lambda m: f"{m // 60}h{m % 60:02d}"
    print("Per day:")
    for day in sorted(per_day):
        weekday = dt.date.fromisoformat(day).strftime("%a")
        print(f"  {weekday} {day}  {fmt(per_day[day]):>6}")
    print("Per issue:")
    for issue, mins in sorted(per_issue.items(), key=lambda kv: -kv[1]):
        print(f"  {issue:<16} {fmt(mins):>7}")
    print(f"Total: {fmt(sum(per_day.values()))} in {len(entries)} worklogs")


def cmd_post(args):
    entries, tz_name, errors = load_worklogs(args.file, args.month)
    if errors:
        print("\n".join(errors), file=sys.stderr)
        sys.exit(1)
    print_summary(entries)
    base = args.url or os.environ.get("JIRA_URL")
    if not base:
        print("\nValid. Set JIRA_URL (and credentials) to check against Jira; add --apply to post.")
        return
    jira = Jira(base, os.environ.get("JIRA_PAT"), os.environ.get("JIRA_EMAIL"), os.environ.get("JIRA_API_TOKEN"))
    me = jira.call("GET", "/rest/api/2/myself")
    my_ids = {me.get(k) for k in ("key", "name", "accountId") if me.get(k)}
    print(f"\nSigned in to {base} as {me.get('displayName')}")
    tz = tzinfo_for(tz_name)

    existing = {}
    for issue in sorted({e["issue"] for e in entries}):
        try:
            logs = jira.call("GET", f"/rest/api/2/issue/{urllib.parse.quote(issue)}/worklog")["worklogs"]
        except RuntimeError as exc:
            sys.exit(f"Cannot read {issue}: {exc}")
        existing[issue] = {(parse_started(w["started"]), w["timeSpentSeconds"]) for w in logs
                           if my_ids & {w["author"].get(k) for k in ("key", "name", "accountId")}}

    todo = []
    for e in entries:
        started = jira_started(e["date"], e.get("start", "09:00"), tz)
        if (parse_started(started), e["minutes"] * 60) in existing[e["issue"]]:
            print(f"  already logged: {e['issue']} {e['date']} {e.get('start', '09:00')} {e['minutes']}m")
        else:
            todo.append((e, started))
    if not args.apply:
        print(f"\nDry run: {len(todo)} worklogs would be posted. Re-run with --apply to post them.")
        return
    for e, started in todo:
        body = {"started": started, "timeSpentSeconds": e["minutes"] * 60, "comment": e.get("comment", "")}
        jira.call("POST", f"/rest/api/2/issue/{urllib.parse.quote(e['issue'])}/worklog", body)
        print(f"  posted: {e['issue']} {e['date']} {e['minutes']}m")
    print(f"\nPosted {len(todo)} worklogs.")


def main():
    utf8_stdout()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    c = sub.add_parser("collect", help="export a month of git and calendar activity into a bundle folder")
    c.add_argument("--month", required=True, help="YYYY-MM")
    c.add_argument("--root", action="append", required=True,
                   help="folder that contains your checkouts (repeatable; searched recursively)")
    c.add_argument("--author", action="append",
                   help="git author email or name to match (repeatable; default: global git config)")
    c.add_argument("--outlook", action="store_true", help="export the default Outlook calendar (Windows, classic Outlook)")
    c.add_argument("--include-body", action="store_true", help="include the first 300 chars of meeting bodies")
    c.add_argument("--calendar", action="append", help="extra calendar file to include (.csv or .ics)")
    c.add_argument("--notes", help="optional notes file (leave days, ticket for meetings, ...)")
    c.add_argument("--out", help="output folder (default: ~/timesheet-YYYY-MM)")
    c.add_argument("--max-depth", type=int, default=4, help="how deep to search under each root (default 4)")
    c.add_argument("--force", action="store_true", help="reuse a non-empty output folder")
    c.set_defaults(func=cmd_collect)

    p = sub.add_parser("post", help="validate worklogs.json and post it to Jira")
    p.add_argument("file", help="worklogs.json")
    p.add_argument("--month", help="YYYY-MM; reject entries outside this month")
    p.add_argument("--url", help="Jira base URL (default: $JIRA_URL)")
    p.add_argument("--apply", action="store_true", help="actually post; without it this is a dry run")
    p.set_defaults(func=cmd_post)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

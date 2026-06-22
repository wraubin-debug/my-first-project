"""Personal Assistant — a daily, Claude-powered briefing for a team lead.

This module gathers three things that already live on your machine:

  * your pending tasks (tasks.json),
  * the flagged emails cached by get_flagged_emails.py, and
  * the previous day's Inbox emails (read once, via Outlook COM),

then asks the local `claude` command-line tool to act as your chief of staff.
Claude returns a plan split into MORNING work (your high-focus hours, so the
hard/important things go here) and AFTERNOON work (your low-energy hours, so
lighter, reactive, or administrative things go here), plus a heads-up on
upcoming high-priority tasks.

Design choices that match the rest of this project:

  * Local only. The `claude` CLI is the same one you already run on this
    machine; nothing new is sent anywhere you weren't already sending it.
  * Gentle on Outlook. We read the Inbox once a day with a single date-limited
    query, the same "fetch once, cache, throttle" approach as the flagged reader.
  * The briefing is cached to assistant_briefing_cache.json so the GUI opens
    instantly and only regenerates when you ask or once a day after 7 AM.
"""

import json
import os
import shutil
import subprocess
from datetime import datetime, timedelta

import get_flagged_emails  # reuse its cache reader for flagged emails

_HERE = os.path.dirname(os.path.abspath(__file__))
TASKS_FILE = os.path.join(_HERE, "tasks.json")
CACHE_FILE = os.path.join(_HERE, "assistant_briefing_cache.json")

# Like the flagged-email cache, the briefing is regenerated once a day after
# this hour. Before it, we reuse yesterday's briefing.
DAILY_HOUR = 7

# How many of yesterday's emails to hand to Claude, and how much of each body.
MAX_RECENT_EMAILS = 40
BODY_SNIPPET_CHARS = 280

# Claude can take a while to think; give it room but don't hang forever.
CLAUDE_TIMEOUT_SECONDS = 240


# ── Gathering the inputs ──────────────────────────────────────────────────────

def _load_pending_tasks():
    """Return the not-done tasks from tasks.json, newest-priority-first-ish.
    Returns [] if the file is missing or unreadable."""
    try:
        with open(TASKS_FILE, "r", encoding="utf-8") as f:
            tasks = json.load(f)
    except Exception:
        return []
    return [t for t in tasks if not t.get("done")]


def get_previous_day_emails(limit=MAX_RECENT_EMAILS):
    """Return (emails, error) for messages received *yesterday* from the Inbox.

    Each email is a dict: subject, sender, received (time string), snippet.
    This is one date-restricted read of the default Inbox — light on Outlook,
    in the same spirit as get_flagged_emails. error is None on success."""
    try:
        import win32com.client
    except ImportError:
        return [], "pywin32 not installed (pip install pywin32)"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    except Exception as error:
        return [], f"Couldn't connect to Outlook ({error})"

    # Yesterday, from midnight to midnight.
    today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    start = today_start - timedelta(days=1)
    end = today_start

    # Outlook's Restrict wants US-style date strings. Sorting first makes the
    # restricted scan cheaper on large inboxes.
    time_format = "%m/%d/%Y %I:%M %p"
    restriction = (
        f"[ReceivedTime] >= '{start.strftime(time_format)}' "
        f"AND [ReceivedTime] < '{end.strftime(time_format)}'"
    )

    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)  # newest first
        recent = items.Restrict(restriction)
    except Exception as error:
        return [], f"Couldn't read the Inbox ({error})"

    emails = []
    for item in recent:
        try:
            body = (getattr(item, "Body", "") or "").strip()
            snippet = " ".join(body.split())[:BODY_SNIPPET_CHARS]
            received_dt = getattr(item, "ReceivedTime", None)
            emails.append({
                "subject": (getattr(item, "Subject", "") or "(no subject)").strip(),
                "sender": (getattr(item, "SenderName", "") or "").strip(),
                "received": received_dt.strftime("%I:%M %p") if received_dt else "",
                "snippet": snippet,
            })
        except Exception:
            continue  # skip anything we can't read; keep the rest
        if len(emails) >= limit:
            break
    return emails, None


# ── Building the prompt ─────────────────────────────────────────────────────--

def _format_tasks(tasks):
    if not tasks:
        return "(no pending tasks)"
    lines = []
    for task in tasks:
        priority = task.get("priority", "Low")
        project = task.get("project", "General")
        lines.append(f"- [{priority}] {task.get('title', '').strip()} (project: {project})")
    return "\n".join(lines)


def _format_flagged(flagged):
    if not flagged:
        return "(no flagged emails)"
    lines = []
    for email in flagged:
        detail = email.get("sender", "")
        if email.get("received"):
            detail = f"{detail}, {email['received']}" if detail else email["received"]
        lines.append(f"- {email.get('subject', '(no subject)')} — {detail}")
    return "\n".join(lines)


def _format_recent_emails(emails):
    if not emails:
        return "(no emails received yesterday)"
    lines = []
    for email in emails:
        header = f"- {email.get('subject', '(no subject)')} — {email.get('sender', '')}"
        if email.get("received"):
            header += f" ({email['received']})"
        lines.append(header)
        if email.get("snippet"):
            lines.append(f"    {email['snippet']}")
    return "\n".join(lines)


def build_prompt(tasks, flagged, recent_emails, now=None):
    """Assemble the instruction + context we send to Claude."""
    now = now or datetime.now()
    today_label = now.strftime("%A, %B %d, %Y")

    return f"""You are the personal chief of staff for a team lead. Today is {today_label}.

Your job is to turn the information below into a focused daily plan. Apply solid
management principles:
- Prioritize ruthlessly with the Eisenhower idea (urgent vs. important); important-
  but-not-urgent work is what usually gets dropped, so protect it.
- A team lead's own blockers gate the whole team's throughput, so surface anything
  that unblocks teammates, answers a waiting person, or keeps a project moving.
- Flag work that should be delegated rather than done personally.
- Protect one block of deep-focus time; don't fragment the morning with small stuff.

Energy model for THIS person (important):
- They are MOST productive in the MORNING. Put the hardest, most cognitively
  demanding, and highest-priority/important work here — deep focus, decisions,
  writing, planning, anything that needs a clear head.
- They get SLEEPY in the AFTERNOON. Put lighter, reactive, or administrative work
  here — email replies, quick approvals, routine updates, meetings, low-stakes tasks.

== PENDING TASKS ==
{_format_tasks(tasks)}

== FLAGGED EMAILS (need follow-up) ==
{_format_flagged(flagged)}

== EMAILS RECEIVED YESTERDAY ==
{_format_recent_emails(recent_emails)}

Now write the plan as plain text (no markdown tables) using exactly these sections:

MORNING (high-focus work)
- 3 to 5 concrete action items, each one short line. Hardest/most important first.

AFTERNOON (lighter work)
- 3 to 5 concrete action items, each one short line. Reactive/admin/meetings here.

UPCOMING HIGH-PRIORITY
- Any high-priority tasks or commitments to watch over the coming days.

NOTES
- 1 to 3 brief management observations: what to delegate, what to protect, what's at risk.

Be specific and reference the actual tasks/emails above. Keep it tight and scannable."""


# ── Calling Claude ──────────────────────────────────────────────────────────--

def run_claude(prompt):
    """Send the prompt to the local `claude` CLI and return (text, error).

    Uses print mode (`claude -p`) with the prompt piped over stdin, so a long
    prompt never runs into command-line length limits. error is None on success."""
    claude_path = shutil.which("claude")
    if not claude_path:
        return "", "The 'claude' command wasn't found on PATH."
    try:
        result = subprocess.run(
            [claude_path, "-p"],
            input=prompt,
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=CLAUDE_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return "", "Claude took too long to respond."
    except Exception as error:
        return "", f"Couldn't run Claude ({error})"

    if result.returncode != 0:
        message = (result.stderr or "").strip() or "Claude returned an error."
        return "", message
    return (result.stdout or "").strip(), None


# ── Cache ─────────────────────────────────────────────────────────────────────

def read_cached_briefing():
    """Return (briefing_text, updated_label) from the cache, instantly.
    Returns ("", None) when there's no cache yet."""
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("briefing", ""), data.get("updated")
    except Exception:
        return "", None


def _most_recent_boundary(now):
    """The most recent time today's DAILY_HOUR has passed (or yesterday's if
    it's still early). Mirrors the throttle in get_flagged_emails."""
    boundary = now.replace(hour=DAILY_HOUR, minute=0, second=0, microsecond=0)
    if now < boundary:
        boundary -= timedelta(days=1)
    return boundary


def _cache_is_fresh():
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        last_iso = data.get("updated_iso")
        if not last_iso:
            return False
        return datetime.fromisoformat(last_iso) >= _most_recent_boundary(datetime.now())
    except Exception:
        return False


def generate_briefing(force=False):
    """Gather the inputs, ask Claude, and cache the result. Returns (text, error).

    Like refresh_cache() for emails, this only does real work when the cache is
    stale (older than today's 7 AM) or force=True; otherwise it returns the
    cached briefing untouched. On error the existing cache is left in place."""
    if not force and _cache_is_fresh():
        return read_cached_briefing()[0], None

    tasks = _load_pending_tasks()
    flagged, _ = get_flagged_emails.read_cached_emails()
    recent_emails, email_error = get_previous_day_emails()
    # A failure to read yesterday's mail isn't fatal — Claude can still plan from
    # tasks and flagged items. We just note it for context.
    if email_error:
        recent_emails = []

    prompt = build_prompt(tasks, flagged, recent_emails)
    briefing, error = run_claude(prompt)
    if error:
        return "", error

    now = datetime.now()
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "briefing": briefing,
                "updated": now.strftime("%b %d, %I:%M %p"),
                "updated_iso": now.isoformat(),
            }, f, indent=2)
    except Exception:
        pass
    return briefing, None


if __name__ == "__main__":
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    # `--refresh` : force-regenerate the briefing and cache it (the 7:05 AM task).
    # (no args)   : print the current briefing, generating it if the cache is stale.
    force = "--refresh" in sys.argv
    briefing, error = generate_briefing(force=force)
    if error:
        print("Briefing error:", error)
    elif not briefing:
        print("No briefing produced.")
    else:
        print(briefing)

"""Personal Assistant — a daily, Claude-powered briefing + chat for a team lead.

This module does three jobs, all locally via the `claude` command-line tool:

  1. BRIEFING (morning): turns your pending tasks, flagged emails, the rolling
     topic memory (below), and any overnight email into a morning/afternoon plan.
  2. ROLLING MEMORY (nightly): instead of storing each day's raw email, a 10 PM
     job asks Claude to fold the day's mail into a small set of TOPIC SUMMARIES
     and discards the raw messages. Claude adds to this memory a little each
     night; topics older than 30 days are flushed unless you marked them
     important or they're still tied to an open task.
  3. CHAT: you can talk to the assistant ("bump the China tax task to High; I
     finished the offboarding; the BM Wave-2 topic is important"). Claude replies
     and the changes are applied to tasks.json / the topic memory automatically.

Design choices that match the rest of this project:

  * Local only. The `claude` CLI is the same one you already run on this machine.
  * Gentle on Outlook. Email is read with a single date-limited query, the same
    "fetch once, cache, throttle" approach as the flagged reader.
  * Small JSON caches so the GUI opens instantly and rarely calls Claude on open.
"""

import json
import os
import re
import shutil
import subprocess
from datetime import datetime, timedelta

import get_flagged_emails  # reuse its cache reader for flagged emails

_HERE = os.path.dirname(os.path.abspath(__file__))
TASKS_FILE = os.path.join(_HERE, "tasks.json")
CACHE_FILE = os.path.join(_HERE, "assistant_briefing_cache.json")
MEMORY_FILE = os.path.join(_HERE, "assistant_memory.json")
CHAT_FILE = os.path.join(_HERE, "assistant_chat.json")

# The briefing regenerates once a day after this hour (the 7:05 AM task).
DAILY_HOUR = 7
# The rolling memory is rebuilt once a day after this hour (the 10:05 PM task).
NIGHTLY_HOUR = 22

# How much of the day's mail to hand Claude at the nightly summary, and how much
# of each body. This runs at 10 PM (off-hours), so a generous read is fine.
MAX_RECENT_EMAILS = 100
BODY_SNIPPET_CHARS = 1500

# The morning briefing's "overnight" read is deliberately small — it only covers
# the gap between last night's summary and now.
OVERNIGHT_EMAIL_LIMIT = 30
OVERNIGHT_SNIPPET_CHARS = 400

# Topics not touched in this many days are flushed (unless important / still open).
MEMORY_MAX_AGE_DAYS = 30
# How many chat turns to keep for continuity.
MAX_CHAT_TURNS = 20

VALID_PRIORITIES = ("High", "Medium", "Low")

# Claude can take a while to think; give it room but don't hang forever.
CLAUDE_TIMEOUT_SECONDS = 240


# ── Gathering task inputs ─────────────────────────────────────────────────────

def _load_all_tasks():
    """Return every task from tasks.json (with ids), or [] if unreadable."""
    try:
        with open(TASKS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _load_pending_tasks():
    """Return the not-done tasks from tasks.json. [] if missing/unreadable."""
    return [t for t in _load_all_tasks() if not t.get("done")]


# ── Reading email from Outlook ────────────────────────────────────────────────

def _find_primary_inbox(namespace):
    """Return the Inbox of the primary mailbox, or None.

    GetDefaultFolder(6) follows the *default delivery store*, which on this setup
    is a local .pst with an empty Inbox — the real mailbox isn't the DefaultStore
    (the same reason get_flagged_emails scans all stores). So we look across every
    store and pick the Inbox that actually has mail."""
    best_inbox = None
    best_count = -1
    for store in namespace.Stores:
        try:
            inbox = store.GetDefaultFolder(6)  # 6 = olFolderInbox
            count = inbox.Items.Count
        except Exception:
            continue  # some stores (SharePoint, online archive) have no usable Inbox
        if count > best_count:
            best_inbox = inbox
            best_count = count
    return best_inbox


def read_inbox_between(start, end, limit, snippet_chars):
    """Return (emails, error) for Inbox messages received in [start, end).

    Each email is a dict: subject, sender, received (time string), snippet.
    One date-restricted read of the primary mailbox's Inbox — light on Outlook,
    the same spirit as get_flagged_emails. error is None on success."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return [], "pywin32 not installed (pip install pywin32)"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = _find_primary_inbox(namespace)
    except Exception as error:
        return [], f"Couldn't connect to Outlook ({error})"
    if inbox is None:
        return [], "Couldn't find a mailbox Inbox in Outlook"

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
            snippet = " ".join(body.split())[:snippet_chars]
            received_dt = getattr(item, "ReceivedTime", None)
            emails.append({
                "subject": (getattr(item, "Subject", "") or "(no subject)").strip(),
                "sender": (getattr(item, "SenderName", "") or "").strip(),
                "received": received_dt.strftime("%b %d %I:%M %p") if received_dt else "",
                "snippet": snippet,
            })
        except Exception:
            continue  # skip anything we can't read; keep the rest
        if len(emails) >= limit:
            break
    return emails, None


def get_todays_emails():
    """The day's mail so far (midnight → now) — for the nightly summary."""
    today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    return read_inbox_between(today_start, datetime.now(),
                              MAX_RECENT_EMAILS, BODY_SNIPPET_CHARS)


def get_overnight_emails(since_iso=None):
    """Mail since the last memory update (or last 12h) — for the morning briefing.
    A small read that just covers the 10 PM → 7 AM gap."""
    now = datetime.now()
    start = now - timedelta(hours=12)
    if since_iso:
        try:
            start = datetime.fromisoformat(since_iso)
        except Exception:
            pass
    return read_inbox_between(start, now, OVERNIGHT_EMAIL_LIMIT, OVERNIGHT_SNIPPET_CHARS)


# ── Rolling topic memory ──────────────────────────────────────────────────────

def read_memory():
    """Return (topics, updated_label) from assistant_memory.json.
    topics is [] and updated_label is None when there's no memory yet."""
    try:
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("topics", []), data.get("updated")
    except Exception:
        return [], None


def _write_memory(topics):
    now = datetime.now()
    try:
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "topics": topics,
                "updated": now.strftime("%b %d, %I:%M %p"),
                "updated_iso": now.isoformat(),
            }, f, indent=2)
    except Exception:
        pass


def _memory_updated_iso():
    try:
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get("updated_iso")
    except Exception:
        return None


def format_memory_for_prompt(topics):
    if not topics:
        return "(no topic memory yet)"
    lines = []
    for topic in topics:
        tags = []
        if topic.get("important"):
            tags.append("IMPORTANT")
        if topic.get("still_open"):
            tags.append("open")
        tag_text = f" [{', '.join(tags)}]" if tags else ""
        dates = f"(since {topic.get('first_seen', '?')}, updated {topic.get('last_updated', '?')})"
        lines.append(f"- {topic.get('topic', '(untitled)')}{tag_text} {dates}")
        if topic.get("summary"):
            lines.append(f"    {topic['summary']}")
    return "\n".join(lines)


def prune_old_topics(topics, now=None):
    """Drop topics older than MEMORY_MAX_AGE_DAYS unless they're marked important
    or still tied to an open task. Pure function — easy to reason about/test."""
    now = now or datetime.now()
    cutoff = now - timedelta(days=MEMORY_MAX_AGE_DAYS)
    kept = []
    for topic in topics:
        if topic.get("important") or topic.get("still_open"):
            kept.append(topic)
            continue
        last_updated = topic.get("last_updated", "")
        try:
            last_dt = datetime.strptime(last_updated, "%Y-%m-%d")
        except Exception:
            kept.append(topic)  # undated/unparseable — keep it to be safe
            continue
        if last_dt >= cutoff:
            kept.append(topic)
    return kept


def _reconcile_topics(old_topics, new_topics, today):
    """Merge Claude's returned topics back over the old ones so we never lose
    user-controlled state: the `important` flag (set via chat) is preserved, and
    `first_seen` keeps the earliest date. Fills in any missing fields."""
    old_by_name = {t.get("topic", "").strip().lower(): t for t in old_topics}
    reconciled = []
    for topic in new_topics:
        name = topic.get("topic", "").strip()
        if not name:
            continue
        prior = old_by_name.get(name.lower(), {})
        reconciled.append({
            "topic": name,
            "summary": (topic.get("summary") or prior.get("summary") or "").strip(),
            "first_seen": prior.get("first_seen") or topic.get("first_seen") or today,
            "last_updated": topic.get("last_updated") or today,
            # important is user-controlled: once set, the nightly merge can't clear it.
            "important": bool(prior.get("important") or topic.get("important")),
            "still_open": bool(topic.get("still_open", True)),
        })
    return reconciled


# ── Chat history ──────────────────────────────────────────────────────────────

def read_chat():
    """Return the list of recent chat turns (oldest first)."""
    try:
        with open(CHAT_FILE, "r", encoding="utf-8") as f:
            return json.load(f).get("turns", [])
    except Exception:
        return []


def _append_chat_turns(user_text, assistant_text):
    turns = read_chat()
    turns.append({"role": "user", "text": user_text})
    turns.append({"role": "assistant", "text": assistant_text})
    turns = turns[-MAX_CHAT_TURNS:]
    try:
        with open(CHAT_FILE, "w", encoding="utf-8") as f:
            json.dump({"turns": turns}, f, indent=2)
    except Exception:
        pass


def _format_chat_history(turns):
    if not turns:
        return "(no earlier conversation)"
    lines = []
    for turn in turns:
        who = "You" if turn.get("role") == "user" else "Assistant"
        lines.append(f"{who}: {turn.get('text', '').strip()}")
    return "\n".join(lines)


# ── Pulling JSON out of Claude's reply ────────────────────────────────────────

def _extract_json(text):
    """Best-effort: return the JSON value embedded in Claude's reply, or None.

    Tolerates a plain JSON reply, a ```json fenced block, or JSON surrounded by
    a little prose (takes the outermost {...} or [...])."""
    if not text:
        return None
    text = text.strip()
    try:
        return json.loads(text)
    except Exception:
        pass

    fenced = re.search(r"```(?:json)?\s*(.+?)```", text, re.DOTALL)
    if fenced:
        try:
            return json.loads(fenced.group(1).strip())
        except Exception:
            pass

    # Fall back to the outermost bracketed span.
    for open_char, close_char in (("{", "}"), ("[", "]")):
        start = text.find(open_char)
        end = text.rfind(close_char)
        if start != -1 and end > start:
            try:
                return json.loads(text[start:end + 1])
            except Exception:
                continue
    return None


# ── Formatting helpers for prompts ────────────────────────────────────────────

def _format_pending_tasks(tasks):
    if not tasks:
        return "(no pending tasks)"
    lines = []
    for task in tasks:
        priority = task.get("priority", "Low")
        project = task.get("project", "General")
        lines.append(f"- [{priority}] {task.get('title', '').strip()} (project: {project})")
    return "\n".join(lines)


def _format_tasks_with_ids(tasks):
    if not tasks:
        return "(no tasks)"
    lines = []
    for task in tasks:
        state = "done" if task.get("done") else "pending"
        priority = task.get("priority", "Low")
        project = task.get("project", "General")
        lines.append(
            f"- id {task.get('id')}: {task.get('title', '').strip()} "
            f"[{priority}, {project}, {state}]"
        )
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


def _format_emails(emails):
    if not emails:
        return "(none)"
    lines = []
    for email in emails:
        header = f"- {email.get('subject', '(no subject)')} — {email.get('sender', '')}"
        if email.get("received"):
            header += f" ({email['received']})"
        lines.append(header)
        if email.get("snippet"):
            lines.append(f"    {email['snippet']}")
    return "\n".join(lines)


# ── The morning briefing ──────────────────────────────────────────────────────

def build_prompt(tasks, flagged, topics, overnight_emails, now=None):
    """Assemble the instruction + context we send to Claude for the briefing."""
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

Stay grounded — do NOT invent relationships between items. The tasks, emails, and
topics below are independent unless an input explicitly says otherwise. Never claim
one task "feeds", "relates to", or "should be done together with" another, and never
assume two items are the same project just because they sound similar. If you group
items, only group ones the inputs actually connect; otherwise treat each on its own.

Energy model for THIS person (important):
- They are MOST productive in the MORNING. Put the hardest, most cognitively
  demanding, and highest-priority/important work here — deep focus, decisions,
  writing, planning, anything that needs a clear head.
- They get SLEEPY in the AFTERNOON. Put lighter, reactive, or administrative work
  here — email replies, quick approvals, routine updates, meetings, low-stakes tasks.

== PENDING TASKS ==
{_format_pending_tasks(tasks)}

== FLAGGED EMAILS (need follow-up) ==
{_format_flagged(flagged)}

== TOPIC MEMORY (rolling summary of recent email themes) ==
{format_memory_for_prompt(topics)}

== EMAILS SINCE LAST NIGHT (overnight) ==
{_format_emails(overnight_emails)}

Now write the plan as plain text (no markdown tables) using exactly these sections:

MORNING (high-focus work)
- 3 to 5 concrete action items, each one short line. Hardest/most important first.

AFTERNOON (lighter work)
- 3 to 5 concrete action items, each one short line. Reactive/admin/meetings here.

UPCOMING HIGH-PRIORITY
- Any high-priority tasks or commitments to watch over the coming days.

NOTES
- 1 to 3 brief management observations: what to delegate, what to protect, what's at risk.

Be specific and reference the actual tasks/emails/topics above — but only as they
appear, without inventing links between them. Keep it tight and scannable."""


# ── Calling Claude ────────────────────────────────────────────────────────────

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


# ── Briefing cache ────────────────────────────────────────────────────────────

def read_cached_briefing():
    """Return (briefing_text, updated_label) from the cache, instantly.
    Returns ("", None) when there's no cache yet."""
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("briefing", ""), data.get("updated")
    except Exception:
        return "", None


def _most_recent_boundary(now, hour):
    """The most recent time today's `hour` has passed (or yesterday's if it's
    still earlier than that). Mirrors the throttle in get_flagged_emails."""
    boundary = now.replace(hour=hour, minute=0, second=0, microsecond=0)
    if now < boundary:
        boundary -= timedelta(days=1)
    return boundary


def _briefing_is_fresh():
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            last_iso = json.load(f).get("updated_iso")
        if not last_iso:
            return False
        return datetime.fromisoformat(last_iso) >= _most_recent_boundary(datetime.now(), DAILY_HOUR)
    except Exception:
        return False


def generate_briefing(force=False):
    """Gather the inputs, ask Claude, and cache the result. Returns (text, error).

    Only does real work when the cache is stale (older than today's 7 AM) or
    force=True; otherwise returns the cached briefing untouched. On error the
    existing cache is left in place."""
    if not force and _briefing_is_fresh():
        return read_cached_briefing()[0], None

    tasks = _load_pending_tasks()
    flagged, _ = get_flagged_emails.read_cached_emails()
    topics, _ = read_memory()
    overnight, email_error = get_overnight_emails(since_iso=_memory_updated_iso())
    if email_error:
        overnight = []  # not fatal — Claude can still plan from tasks/topics

    prompt = build_prompt(tasks, flagged, topics, overnight)
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


# ── Nightly memory update ─────────────────────────────────────────────────────

def _memory_is_fresh():
    """True if the memory was already rebuilt since today's 10 PM mark."""
    try:
        last_iso = _memory_updated_iso()
        if not last_iso:
            return False
        return datetime.fromisoformat(last_iso) >= _most_recent_boundary(datetime.now(), NIGHTLY_HOUR)
    except Exception:
        return False


def build_memory_prompt(old_topics, emails, open_task_titles, today):
    return f"""You maintain a rolling TOPIC MEMORY for a team lead — a small set of
running summaries of their email themes, so we never have to keep raw emails around.

Today is {today}. Below are the EXISTING topics, TODAY'S emails, and the lead's
currently OPEN tasks. Fold today's emails into the topics:
- Merge each email into an existing topic when it fits; otherwise create a new topic.
- Keep each summary tight (1-3 sentences) and factual; carry forward what still matters.
- For every topic, set "last_updated" to {today} if today's emails touched it, else
  leave its previous date. Keep "first_seen" as the earliest date you can.
- Set "still_open" to true if the topic maps to an open task or is an unresolved
  thread awaiting action; false if it looks resolved/informational.
- Do NOT invent topics with no basis in the emails or existing memory.

== EXISTING TOPICS ==
{format_memory_for_prompt(old_topics)}

== TODAY'S EMAILS ==
{_format_emails(emails)}

== OPEN TASKS ==
{open_task_titles or "(none)"}

Return ONLY a JSON array (no prose, no markdown fences) of topic objects, each:
{{"topic": "...", "summary": "...", "first_seen": "YYYY-MM-DD", "last_updated": "YYYY-MM-DD", "still_open": true}}"""


def update_memory_from_emails(force=False):
    """Fold the day's email into the rolling topic memory, then prune old topics.
    Returns (topics, error). Throttled to once a day after 10 PM unless force."""
    if not force and _memory_is_fresh():
        return read_memory()[0], None

    old_topics, _ = read_memory()
    emails, email_error = get_todays_emails()
    if email_error:
        # Can't read mail — still prune the existing memory so aging keeps working.
        pruned = prune_old_topics(old_topics)
        _write_memory(pruned)
        return pruned, email_error

    today = datetime.now().strftime("%Y-%m-%d")
    open_titles = "\n".join(
        f"- {t.get('title', '').strip()}" for t in _load_pending_tasks()
    )
    prompt = build_memory_prompt(old_topics, emails, open_titles, today)
    reply, error = run_claude(prompt)
    if error:
        return old_topics, error

    parsed = _extract_json(reply)
    if not isinstance(parsed, list):
        # Claude didn't return usable JSON — keep the old memory rather than wipe it.
        return old_topics, "Claude's memory update couldn't be parsed; memory kept as-is."

    reconciled = _reconcile_topics(old_topics, parsed, today)
    pruned = prune_old_topics(reconciled)
    _write_memory(pruned)
    return pruned, None


# ── Chat ──────────────────────────────────────────────────────────────────────

def build_chat_prompt(user_message, all_tasks, topics, history):
    today = datetime.now().strftime("%A, %B %d, %Y")
    return f"""You are the personal assistant and chief of staff for a team lead.
Today is {today}. Have a natural, brief conversation, and when they ask you to
change something, emit the corresponding actions so it actually happens.

You can:
- Update existing tasks (priority, mark done/undone, retitle, change project) — by id.
- Add new tasks.
- Mark a topic in the rolling memory important (so it's never auto-flushed) or
  flag whether it's still open.

== TASKS (id: title [priority, project, state]) ==
{_format_tasks_with_ids(all_tasks)}

== TOPIC MEMORY ==
{format_memory_for_prompt(topics)}

== RECENT CONVERSATION ==
{_format_chat_history(history)}

== USER'S NEW MESSAGE ==
{user_message}

Respond with ONLY a JSON object (no prose outside it, no markdown fences):
{{
  "reply": "your short, friendly reply to the user",
  "task_updates": [{{"id": 7, "set": {{"priority": "High"}}}}],
  "new_tasks": [{{"title": "...", "priority": "Medium", "project": "General"}}],
  "memory": [{{"topic": "exact topic name", "important": true, "still_open": true}}]
}}
Use only valid priorities: High, Medium, Low. "set" may include priority, done
(true/false), title, project. Omit any array you don't need (use []). Only include
changes the user actually asked for."""


def apply_chat_actions(actions):
    """Apply task/memory changes from a parsed chat action object.
    Returns a short human-readable summary of what changed ("" if nothing)."""
    summary = []

    task_updates = actions.get("task_updates") or []
    new_tasks = actions.get("new_tasks") or []
    if task_updates or new_tasks:
        # Re-read immediately before writing so we don't clobber other edits.
        tasks = _load_all_tasks()
        by_id = {t.get("id"): t for t in tasks}

        for update in task_updates:
            task = by_id.get(update.get("id"))
            if not task:
                continue
            changes = update.get("set") or {}
            if "priority" in changes and changes["priority"] in VALID_PRIORITIES:
                task["priority"] = changes["priority"]
                summary.append(f"{task.get('title', 'task')} → {changes['priority']}")
            if "done" in changes:
                task["done"] = bool(changes["done"])
                summary.append(f"{task.get('title', 'task')} → {'done' if task['done'] else 'reopened'}")
            if changes.get("title"):
                task["title"] = str(changes["title"]).strip()
                summary.append(f"renamed to “{task['title']}”")
            if changes.get("project"):
                task["project"] = str(changes["project"]).strip()
                summary.append(f"{task.get('title', 'task')} → project {task['project']}")

        next_id = max((t.get("id", 0) for t in tasks), default=0) + 1
        for new_task in new_tasks:
            title = (new_task.get("title") or "").strip()
            if not title:
                continue
            priority = new_task.get("priority")
            if priority not in VALID_PRIORITIES:
                priority = "Low"
            tasks.append({
                "id": next_id,
                "title": title,
                "priority": priority,
                "done": False,
                "project": (new_task.get("project") or "General").strip(),
            })
            summary.append(f"added “{title}”")
            next_id += 1

        try:
            with open(TASKS_FILE, "w", encoding="utf-8") as f:
                json.dump(tasks, f, indent=2)
        except Exception:
            pass

    memory_updates = actions.get("memory") or []
    if memory_updates:
        topics, _ = read_memory()
        by_name = {t.get("topic", "").strip().lower(): t for t in topics}
        today = datetime.now().strftime("%Y-%m-%d")
        for memory_update in memory_updates:
            name = (memory_update.get("topic") or "").strip()
            if not name:
                continue
            topic = by_name.get(name.lower())
            if not topic:
                topic = {
                    "topic": name, "summary": "", "first_seen": today,
                    "last_updated": today, "important": False, "still_open": True,
                }
                topics.append(topic)
                by_name[name.lower()] = topic
            if "important" in memory_update:
                topic["important"] = bool(memory_update["important"])
                if topic["important"]:
                    summary.append(f"marked “{name}” important")
            if "still_open" in memory_update:
                topic["still_open"] = bool(memory_update["still_open"])
        _write_memory(topics)

    return "; ".join(summary)


def chat_with_assistant(user_message):
    """Send the user's message + context to Claude, apply any actions it returns,
    and record the turn. Returns (reply_text, applied_summary, error)."""
    user_message = (user_message or "").strip()
    if not user_message:
        return "", "", None

    all_tasks = _load_all_tasks()
    topics, _ = read_memory()
    history = read_chat()

    prompt = build_chat_prompt(user_message, all_tasks, topics, history)
    reply, error = run_claude(prompt)
    if error:
        return "", "", error

    actions = _extract_json(reply)
    if isinstance(actions, dict):
        reply_text = (actions.get("reply") or "").strip() or "Done."
        applied = apply_chat_actions(actions)
    else:
        # Couldn't parse structured actions — just show whatever Claude said.
        reply_text = reply.strip() or "Done."
        applied = ""

    _append_chat_turns(user_message, reply_text)
    return reply_text, applied, None


if __name__ == "__main__":
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    # --refresh        : force-regenerate the briefing and cache it (7:05 AM task).
    # --update-memory  : fold today's email into the rolling memory (10:05 PM task).
    # --chat "message" : one-off chat turn from the command line (handy for testing).
    # (no args)        : print the current briefing, generating it if the cache is stale.
    if "--update-memory" in sys.argv:
        topics, error = update_memory_from_emails(force=True)
        if error:
            print("Memory update note:", error)
        print(f"Topic memory now has {len(topics)} topic(s).")
    elif "--chat" in sys.argv:
        index = sys.argv.index("--chat")
        message = " ".join(sys.argv[index + 1:]).strip()
        reply, applied, error = chat_with_assistant(message)
        if error:
            print("Chat error:", error)
        else:
            print(reply)
            if applied:
                print("\n[applied]", applied)
    else:
        force = "--refresh" in sys.argv
        briefing, error = generate_briefing(force=force)
        if error:
            print("Briefing error:", error)
        elif not briefing:
            print("No briefing produced.")
        else:
            print(briefing)

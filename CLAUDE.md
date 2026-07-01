# Morning Agent — Task Manager & Briefing

A personal productivity suite in Python for Windows. It started as a CLI to-do
list and grew into a desktop task manager, a Claude-powered daily briefing, a
phone-friendly view, and an Outlook flagged-email summary. Everything runs
locally — no cloud services, no third-party storage.

## Components

| File | Type | What it does |
|------|------|--------------|
| `task_manager.py` | Desktop GUI (CustomTkinter) | Main app (light theme): two tabs — **Tasks** (multiple projects, add/complete/delete tasks, **drag the ↕ handle to reorder projects**, live flagged-emails panel) and **Assistant** (the Claude-powered daily briefing **plus a chat box** to talk to the assistant). |
| `personal_assistant.py` | Library + CLI | Builds a daily plan: gathers pending tasks, cached flagged emails, the rolling topic memory, and any overnight Inbox mail, then asks the local `claude` CLI to act as a chief of staff and split the work into morning (high-focus) and afternoon (low-energy) action items. Also (a) maintains a **rolling topic memory** distilled nightly from the day's email (raw mail is discarded), and (b) powers the **chat** — instructions like "raise X to High / I finished Y / mark topic Z important" are sent to Claude, which replies and applies the changes to `tasks.json` / the memory. Caches the briefing to JSON. |
| `get_flagged_emails.py` | Library + CLI | Reads flagged emails from local Outlook (the "For Follow Up" search folder) via COM; caches them to JSON. Fetches once a day to keep Outlook responsive, and defers unflags to a nightly batch. Imported by `task_manager.py`. |
| `generate_mobile_view.py` | Library + CLI | Generates a self-contained `tasks.html` into OneDrive for phone viewing. |
| `schedule_briefing.py` | Setup script | Registers Windows scheduled tasks + creates a desktop shortcut. Run once. |
| `todo.py` | CLI (legacy) | The original terminal to-do app. Still works; reads/writes the same `tasks.json`. |

### Data files (created automatically)

- `tasks.json` — the task list (see structure below). Shared by **all** components.
- `projects.json` — ordered list of project names, e.g. `["General", "Mercury", "BM Wave-2"]`. Lets empty projects persist.
- `flagged_emails_cache.json` — cached flagged emails + a "last updated" label and machine-readable timestamp. **Gitignored** (contains real email subjects/senders).
- `pending_unflags.json` — emails you unflagged in the app that haven't been written back to Outlook yet; drained by the nightly 10 PM batch. **Gitignored** (contains Outlook entry IDs).
- `assistant_briefing_cache.json` — the latest Claude-generated briefing text + a "last updated" label and timestamp. **Gitignored** (built from real tasks/emails). Self-throttled to once a day after 7 AM, same as the flagged cache.
- `assistant_memory.json` — the **rolling topic memory**: a small list of Claude-maintained topic summaries (`topic`, `summary`, `first_seen`, `last_updated`, `important`, `still_open`, `awaiting`) distilled from the day's email each night. `awaiting` records **who owes the next step** — `"me"`, `"them"`, or a name — so the morning briefing can tell my action items from watchpoints without re-reading raw mail. Replaces storing raw daily mail. **Gitignored** (built from real email). Self-throttled to once a day after 10 PM.
- `assistant_chat.json` — the last ~20 chat turns between you and the assistant, kept for conversational continuity. **Gitignored** (may reference real tasks/emails).

### Generated output (outside the repo)

- `OneDrive - Accenture\TaskManager\tasks.html` — read-only phone dashboard, regenerated on every task change in `task_manager.py`.

## How the pieces connect

```
                       tasks.json  ◄──────────────┐
                          ▲                        │
   todo.py (CLI) ─────────┤                        │
                          │                        │
   task_manager.py ───────┼──► generate_mobile_view.py ──► OneDrive/tasks.html (phone)
        │                 │
        │ reads/writes    └──── personal_assistant.py (also reads tasks.json)
        ▼
   get_flagged_emails.py ──► Outlook (COM) ──► flagged_emails_cache.json
        ▲                                              ▲
        │ the app reads the cache instantly,           │
        └── then refreshes it on demand                │
                                                       │
   schedule_briefing.py registers a task that runs ────┘
   `get_flagged_emails.py --refresh` every 15 minutes
```

- `tasks.json` is the single source of truth for tasks. The desktop app, the CLI,
  the briefing, and the mobile view all read the same flat array, so they stay in sync.
- `task_manager.py` regenerates the phone `tasks.html` automatically after every save.
- Flagged emails are cached so the GUIs open instantly. To keep Outlook responsive,
  the cache is refreshed from Outlook only **once a day** (a scheduled task at 7 AM;
  `refresh_cache()` self-throttles to that mark, so the GUIs read the cache without
  poking Outlook). Unflags made in the app are queued and written back to Outlook in
  one **nightly 10 PM batch** rather than immediately.

## `tasks.json` data structure

A **flat JSON array** of task objects (kept flat for backward compatibility with
`todo.py`):

```json
[
  {
    "id": 4,
    "title": "Build Agent",
    "priority": "High",
    "done": false,
    "project": "General"
  }
]
```

| Field | Type | Notes |
|-------|------|-------|
| `id` | int | Unique. Auto-assigned; older tasks without an id are backfilled on load. |
| `title` | string | Required. |
| `priority` | string | `"High"`, `"Medium"`, or `"Low"`. Optional — missing priority is treated as `"Low"`. |
| `done` | bool | Completion state. |
| `project` | string | Project name. Missing project is backfilled to `"General"`. |

`projects.json` is a simple ordered array of project-name strings.

## How to run

**Install dependencies first** (see below), then:

```
python task_manager.py        # main desktop app
python todo.py                # legacy terminal version
python generate_mobile_view.py        # regenerate the phone HTML on demand
python get_flagged_emails.py          # print flagged emails (live read)
python get_flagged_emails.py --refresh        # force the daily fetch + cache write (7 AM task)
python get_flagged_emails.py --apply-unflags  # write queued unflags back to Outlook (10 PM task)
python personal_assistant.py            # print the current briefing (generates it if the cache is stale)
python personal_assistant.py --refresh  # force-regenerate the daily briefing via Claude (7:05 AM task)
python personal_assistant.py --update-memory   # fold today's email into the rolling topic memory (10:05 PM task)
python personal_assistant.py --chat "raise the China tax task to High"   # one-off chat turn (testing)
python schedule_briefing.py   # one-time setup: scheduled tasks + desktop shortcut
```

### One-time setup (`schedule_briefing.py`) registers:

- **"Flagged Email Cache Refresh"** — runs `get_flagged_emails.py --refresh` daily at 7:00 AM (via `pythonw`, no console window).
- **"Personal Assistant Briefing"** — runs `personal_assistant.py --refresh` daily at 7:05 AM (just after the flagged fetch, so the briefing sees a fresh flagged cache).
- **"Apply Unflagged Emails"** — runs `get_flagged_emails.py --apply-unflags` daily at 10:00 PM (writes queued unflags back to Outlook).
- **"Assistant Memory Update"** — runs `personal_assistant.py --update-memory` daily at 10:05 PM (just after the unflag batch, so the two nightly jobs don't overlap; folds the day's email into the rolling topic memory).
- A **"Task Manager"** shortcut on the Desktop.

These tasks run while the screen is **locked**, as long as you're still logged in and the
machine is **awake** (not asleep/hibernated/shut down) at the scheduled time. If the machine
is asleep at 7 AM, the fetch is skipped — but the next time you open an app, `refresh_cache()`
sees the cache is older than today's 7 AM and does a single catch-up read.

To remove: `schtasks /delete /tn "Flagged Email Cache Refresh" /f` (and likewise for `"Personal Assistant Briefing"`, `"Apply Unflagged Emails"`, and `"Assistant Memory Update"`).

### Phone access

Open `tasks.html` from the OneDrive app under `TaskManager/`. It's a self-contained,
read-only dashboard (data baked in) so it renders without fetching other files.

## Dependencies

- **`customtkinter`** — modern themed Tkinter UI. `pip install customtkinter`
- **`pywin32`** — Outlook COM access for flagged emails. `pip install pywin32`
- **`claude` CLI** — the Claude Code command-line tool, used by `personal_assistant.py` to generate the briefing. Already installed/authenticated on this machine; the module finds it via `shutil.which("claude")` and pipes the prompt over stdin (`claude -p`). If it's missing, the Assistant tab shows a friendly message and the rest of the app still works.

`todo.py`, `generate_mobile_view.py`, and the data layer use only the standard library.
If `pywin32` or Outlook is unavailable, the GUIs still run — the flagged-emails section
just shows a friendly message.

## Outlook flagged-email reading (important details)

- Reads from the **classic** Outlook desktop client only (COM). The "new Outlook" doesn't support COM automation.
- Reads only the **"For Follow Up" search folder** — a search folder the user set up in Outlook to gather every flagged item, including ones that live in subfolders (the built-in To-Do folder missed subfolder items) — instead of sweeping every folder of every mailbox. That's one cheap read, so Outlook stays responsive. The folder is found by name via `Store.GetSearchFolders()`, searching across all stores because the primary mailbox isn't necessarily the `DefaultStore` (that can be a local .pst). (Trade-off: it covers the primary mailbox's flagged items; flagged items in delegate/archive mailboxes aren't swept. Revisit if that's needed.)
- **Everything in the folder is shown** — no per-item filtering. Because "For Follow Up" is a search folder the user curates in Outlook, whatever lands in it is intentional, so we surface every item: flagged mail plus anything else routed in, such as meeting requests (`IPM.Schedule.Meeting.Request`). (Earlier versions read the built-in To-Do folder and filtered to `IPM.Note` + `FlagStatus == 2` + `IsMarkedAsTask == True` to drop sender-flagged bulk mail; that filtering is no longer needed because the user's search-folder criteria already decide what belongs.)
- **Fetched once a day.** `refresh_cache()` only reads Outlook when the cache is older than today's 7 AM mark (or `force=True`); otherwise it returns the cache untouched. The **task manager does not read Outlook on open** — it shows the cache and only fetches when you click its **Refresh** button (which calls `refresh_cache(force=True)`). The briefing likewise reads the cache.
- Outlook calls run on a **background thread** (with `pythoncom.CoInitialize()`); results return to the UI through a `queue.Queue` polled on the main thread (Tkinter is not thread-safe).
- Clicking an email row opens it in Outlook (`GetItemFromID(...).Display()`). **Unflag** is deferred: `queue_unflag()` drops the row from the cache/UI instantly and records the request in `pending_unflags.json`; the actual `ClearTaskFlag()` write happens in the nightly 10 PM batch (`apply_pending_unflags()`), so unflagging never pokes Outlook while you're using it. Failed writes stay queued and retry next night.

## Code style

- **Simple and readable over clever.** Straightforward logic a beginner can follow.
- **Clear variable and function names.** Avoid abbreviations or one-letter names.
- **No unnecessary abstractions.** Don't add classes, decorators, or patterns unless they genuinely simplify things.
- **Short comments only when the reason isn't obvious.** Don't comment what the code already says clearly.

## Personal Assistant briefing (how it works)

- The **Assistant** tab in `task_manager.py` reads the cached briefing on open (no Claude call) and only regenerates when you click **Regenerate** or via the 7:05 AM scheduled task. Generation runs on a background thread and posts the result back through a `queue.Queue`, the same Tkinter-safe pattern as the flagged-email read.
- `personal_assistant.py` builds the briefing prompt from four local inputs — pending tasks (`tasks.json`), the flagged-email cache, the **rolling topic memory** (`assistant_memory.json`), and a **light read of overnight Inbox mail** (the gap since last night's memory update; `get_overnight_emails`). It no longer re-reads a full day of raw Inbox at briefing time — that content now lives, distilled, in the topic memory. It then calls the local `claude` CLI.
- The prompt encodes the user's energy model: **mornings are peak focus** (hardest/most important/deep work goes there) and **afternoons are low-energy** (lighter, reactive, admin, meetings). It also applies team-lead management principles (Eisenhower prioritization, unblocking the team first, delegation, protecting a deep-focus block). Output sections: MORNING, AFTERNOON, **WATCHPOINTS**, UPCOMING HIGH-PRIORITY, NOTES.
- **Whose action is it?** The prompt tells Claude who "me" is (`USER_NAME` / `USER_EMAIL` in `personal_assistant.py`) and only treats an email as *my* action item when the ball is clearly in my court — I'm in the **`to:`** line, **@-mentioned** by name, or explicitly asked by name. An unanswered thread, someone else waiting on a third party, or a mail I'm only cc'd on becomes a **WATCHPOINT** (keep an eye on it), not a to-do. To support this, overnight/daily Inbox reads capture the email's **`to`** line and a **`mentions_me`** flag (`read_inbox_between` + `_mentions_user`, which scans the *full body* — not the truncated snippet — for `@`+`USER_MENTION_NAMES` at a word boundary). Emails that @-mention me are surfaced to Claude as `[@-MENTIONS ME]` and folded into a topic marked `awaiting: "me"`. The nightly topic-memory summary carries the `awaiting` field so the "whose action" distinction survives after raw mail is discarded.

### Rolling topic memory (nightly, `--update-memory`)

- Instead of storing each day's raw email, a **10:05 PM** job (`update_memory_from_emails`) reads the day's Inbox (a generous **100 emails × 1500 chars**, fine at off-hours), hands Claude the existing topics + today's mail + the current open-task titles, and asks it to **fold each email into an existing topic or create a new one**. Claude returns the updated topic list as JSON; the raw emails are then discarded. Self-throttled to once a day after 10 PM (like the flagged cache), `force=True` from the scheduled task.
- **Aging:** `prune_old_topics()` drops any topic not touched in **30 days** *unless* it's marked `important` or `still_open`. `still_open` is set by Claude during the nightly merge (true if the topic maps to an open task or an unresolved thread); `important` is set by you via chat and is **never cleared by the nightly merge** (`_reconcile_topics` preserves it, along with the earliest `first_seen`).

### Chat (`chat_with_assistant`)

- The Assistant tab has a **chat box** below the briefing (transcript + entry + Send; Enter sends). It uses the same background-thread + `queue.Queue` pattern; recent turns persist in `assistant_chat.json` and reload on open.
- `chat_with_assistant()` sends Claude the user's message plus context: **all tasks with their ids**, the topic memory, and recent chat turns. Claude replies with a JSON object — a `reply` plus optional `task_updates` (by **id**, so auto-apply is unambiguous), `new_tasks`, and `memory` flags. `apply_chat_actions()` validates and writes the changes straight to `tasks.json` / the memory (priority must be High/Medium/Low; new tasks get a fresh `id`). Per the user's choice, changes **apply automatically** (no confirm gate), but the reply states what changed and the GUI shows an "✓ Applied: …" line, then reloads the Tasks tab and regenerates the phone HTML.

## Known gaps / planned

- **Teams chat analysis (Phase 2, deferred):** the original request included analyzing Teams chats when the system locks. Deferred because "new" Teams has no COM automation and Graph is blocked on this tenant (see below), so there's no supported local way to read chat content. Lock *detection* is feasible (Windows session-lock events); reading the *chats* is the blocker. Revisit if a viable path appears (or wire a manual-paste fallback). *(Planned — not yet done.)*

- **Task filter default:** the task list currently defaults to the **All** filter; the **Active** list should become the default. *(Planned — not yet done.)*
- **Phone view is read-only:** editing tasks from the phone isn't possible (no backend, by design — keeps everything local).
- **Phone view freshness:** `tasks.html` only updates when `task_manager.py` saves a change; it isn't regenerated by the CLI or the briefing.
- **Only the primary mailbox's flagged items:** reading the "For Follow Up" search folder covers all folders of the primary mailbox, but **not** separate delegate/archive mailboxes. This was a deliberate trade-off to keep Outlook responsive (the old code swept every store, which made Outlook sluggish). Re-add per-store reads if those mailboxes are ever needed.
- **Cloud / Graph API path is blocked:** reading flagged emails from Office 365 via Microsoft Graph was evaluated (would avoid Outlook entirely), but Accenture's tenant blocks it — Azure app registration returns 401 and any app needs admin consent. Revisit only if IT approves an app. The COM-based approach above is the supported path.
- **Briefing reads the cache:** the morning briefing shows the cached flagged emails; after the 7 AM fetch its `refresh_cache()` call is a no-op (throttled), so it doesn't poke Outlook.

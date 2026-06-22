# Morning Agent — Task Manager & Briefing

A personal productivity suite in Python for Windows. It started as a CLI to-do
list and grew into a desktop task manager, a daily morning-briefing popup, a
phone-friendly view, and an Outlook flagged-email summary. Everything runs
locally — no cloud services, no third-party storage.

## Components

| File | Type | What it does |
|------|------|--------------|
| `task_manager.py` | Desktop GUI (CustomTkinter) | Main app: multiple projects, add/complete/delete tasks, and a live flagged-emails panel. |
| `morning_briefing.py` | Desktop GUI popup | Daily 9 AM popup: date, motivational quote, top 3 pending tasks, and flagged emails. Has an "Open Task Manager" button. |
| `get_flagged_emails.py` | Library + CLI | Reads flagged emails from local Outlook (the "For Follow Up" search folder) via COM; caches them to JSON. Fetches once a day to keep Outlook responsive, and defers unflags to a nightly batch. Imported by both GUIs. |
| `generate_mobile_view.py` | Library + CLI | Generates a self-contained `tasks.html` into OneDrive for phone viewing. |
| `schedule_briefing.py` | Setup script | Registers Windows scheduled tasks + creates a desktop shortcut. Run once. |
| `todo.py` | CLI (legacy) | The original terminal to-do app. Still works; reads/writes the same `tasks.json`. |

### Data files (created automatically)

- `tasks.json` — the task list (see structure below). Shared by **all** components.
- `projects.json` — ordered list of project names, e.g. `["General", "Mercury", "BM Wave-2"]`. Lets empty projects persist.
- `flagged_emails_cache.json` — cached flagged emails + a "last updated" label and machine-readable timestamp. **Gitignored** (contains real email subjects/senders).
- `pending_unflags.json` — emails you unflagged in the app that haven't been written back to Outlook yet; drained by the nightly 10 PM batch. **Should be gitignored** (contains Outlook entry IDs).

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
        │ reads/writes    └──── morning_briefing.py (also reads tasks.json)
        ▼
   get_flagged_emails.py ──► Outlook (COM) ──► flagged_emails_cache.json
        ▲                                              ▲
        │ both GUIs read the cache instantly,          │
        └── then refresh it in the background          │
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
`todo.py` and `morning_briefing.py`):

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
python morning_briefing.py    # the briefing popup (normally auto-runs at 9 AM)
python todo.py                # legacy terminal version
python generate_mobile_view.py        # regenerate the phone HTML on demand
python get_flagged_emails.py          # print flagged emails (live read)
python get_flagged_emails.py --refresh        # force the daily fetch + cache write (7 AM task)
python get_flagged_emails.py --apply-unflags  # write queued unflags back to Outlook (10 PM task)
python schedule_briefing.py   # one-time setup: scheduled tasks + desktop shortcut
```

### One-time setup (`schedule_briefing.py`) registers:

- **"Morning Briefing"** — runs `morning_briefing.py` daily at 9:00 AM (interactive, so the window shows).
- **"Flagged Email Cache Refresh"** — runs `get_flagged_emails.py --refresh` daily at 7:00 AM (via `pythonw`, no console window).
- **"Apply Unflagged Emails"** — runs `get_flagged_emails.py --apply-unflags` daily at 10:00 PM (writes queued unflags back to Outlook).
- A **"Task Manager"** shortcut on the Desktop.

These tasks run while the screen is **locked**, as long as you're still logged in and the
machine is **awake** (not asleep/hibernated/shut down) at the scheduled time. If the machine
is asleep at 7 AM, the fetch is skipped — but the next time you open an app, `refresh_cache()`
sees the cache is older than today's 7 AM and does a single catch-up read.

To remove: `schtasks /delete /tn "Morning Briefing" /f` (and likewise for `"Flagged Email Cache Refresh"` and `"Apply Unflagged Emails"`).

### Phone access

Open `tasks.html` from the OneDrive app under `TaskManager/`. It's a self-contained,
read-only dashboard (data baked in) so it renders without fetching other files.

## Dependencies

- **`customtkinter`** — modern themed Tkinter UI. `pip install customtkinter`
- **`pywin32`** — Outlook COM access for flagged emails. `pip install pywin32`

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

## Known gaps / planned

- **Task filter default:** the task list currently defaults to the **All** filter; the **Active** list should become the default. *(Planned — not yet done.)*
- **Project ordering:** project order isn't user-editable yet; drag-and-drop reordering is planned. *(Planned — not yet done.)*
- **Phone view is read-only:** editing tasks from the phone isn't possible (no backend, by design — keeps everything local).
- **Phone view freshness:** `tasks.html` only updates when `task_manager.py` saves a change; it isn't regenerated by the CLI or the briefing.
- **Only the primary mailbox's flagged items:** reading the "For Follow Up" search folder covers all folders of the primary mailbox, but **not** separate delegate/archive mailboxes. This was a deliberate trade-off to keep Outlook responsive (the old code swept every store, which made Outlook sluggish). Re-add per-store reads if those mailboxes are ever needed.
- **Cloud / Graph API path is blocked:** reading flagged emails from Office 365 via Microsoft Graph was evaluated (would avoid Outlook entirely), but Accenture's tenant blocks it — Azure app registration returns 401 and any app needs admin consent. Revisit only if IT approves an app. The COM-based approach above is the supported path.
- **Briefing reads the cache:** the morning briefing shows the cached flagged emails; after the 7 AM fetch its `refresh_cache()` call is a no-op (throttled), so it doesn't poke Outlook.

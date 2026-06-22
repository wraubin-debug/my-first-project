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
| `get_flagged_emails.py` | Library + CLI | Reads flagged emails from local Outlook via COM; caches them to JSON. Imported by both GUIs. |
| `generate_mobile_view.py` | Library + CLI | Generates a self-contained `tasks.html` into OneDrive for phone viewing. |
| `schedule_briefing.py` | Setup script | Registers Windows scheduled tasks + creates a desktop shortcut. Run once. |
| `todo.py` | CLI (legacy) | The original terminal to-do app. Still works; reads/writes the same `tasks.json`. |

### Data files (created automatically)

- `tasks.json` — the task list (see structure below). Shared by **all** components.
- `projects.json` — ordered list of project names, e.g. `["General", "Mercury", "BM Wave-2"]`. Lets empty projects persist.
- `flagged_emails_cache.json` — cached flagged emails + a "last updated" label. **Gitignored** (contains real email subjects/senders).

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
- Flagged emails are cached so the GUIs open instantly; the cache is refreshed in the
  background while an app is open and by a scheduled task every 15 minutes.

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
python get_flagged_emails.py --refresh   # refresh the cache (used by the scheduler)
python schedule_briefing.py   # one-time setup: scheduled tasks + desktop shortcut
```

### One-time setup (`schedule_briefing.py`) registers:

- **"Morning Briefing"** — runs `morning_briefing.py` daily at 9:00 AM (interactive, so the window shows).
- **"Flagged Email Cache Refresh"** — runs `get_flagged_emails.py --refresh` every 15 minutes (via `pythonw`, no console window).
- A **"Task Manager"** shortcut on the Desktop.

To remove: `schtasks /delete /tn "Morning Briefing" /f` (and likewise for `"Flagged Email Cache Refresh"`).

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
- Searches **all folders** of every **Exchange** mailbox; skips non-Exchange stores (a conflicted OneDrive `.pst`, SharePoint Lists) and any folder that errors.
- Filtering on `[FlagStatus]` is unreliable in Outlook, so the reader narrows on the `PR_TODO_ITEM_FLAGS` property (DASL), then confirms `FlagStatus == 2`.
- **Only emails the user flagged** are shown: items with `IsMarkedAsTask == True`. Bulk/newsletter mail that arrives pre-flagged by the sender (e.g. the daily "To-Dos and News" digest) has `IsMarkedAsTask == False` and is excluded.
- Outlook calls run on a **background thread** (with `pythoncom.CoInitialize()`); results return to the UI through a `queue.Queue` polled on the main thread (Tkinter is not thread-safe).
- Clicking an email row opens it in Outlook (`GetItemFromID(...).Display()`); **Unflag** clears the flag (`ClearTaskFlag()`), removes it from the cache, and drops just that row — no full re-read.

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
- **Conflicted `.pst`:** flagged emails in the conflicted OneDrive `.pst` are skipped until that file's sync conflict is resolved in Outlook.
- **Briefing reads emails synchronously vs. cached:** the briefing shows cached emails instantly then refreshes once; it does not poll repeatedly (it's a one-shot popup).

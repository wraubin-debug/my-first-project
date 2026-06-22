"""Read flagged emails from the local (classic) Outlook desktop app via COM.

This runs entirely on your machine against the Outlook profile you're already
signed into — no cloud credentials, no third-party service.

To keep Outlook responsive, this module is deliberately gentle with it:

  * It reads only Outlook's built-in **To-Do folder** — the search folder
    Outlook already maintains of every flagged item — instead of walking every
    folder of every mailbox. That's a single, cheap read.
  * The Outlook read happens **once a day** (the scheduled 7 AM fetch).
    refresh_cache() throttles itself: if the cache was already refreshed since
    today's 7 AM mark, it returns the cache and never touches Outlook.
  * Unflagging from the app is **deferred**: queue_unflag() drops the email from
    the cache instantly (so it disappears from the UI) but does not write to
    Outlook. The queued changes are applied in one batch by the nightly 10 PM
    run (apply_pending_unflags()), so Outlook is never poked while you're using it.

The apps therefore read the cache instantly and almost never trigger a live read.
"""

import json
import os
from datetime import datetime, timedelta

_HERE = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(_HERE, "flagged_emails_cache.json")
PENDING_UNFLAGS_FILE = os.path.join(_HERE, "pending_unflags.json")

# We fetch from Outlook once a day at this local hour (the scheduled task runs
# at 07:00). refresh_cache() treats this as the "freshness" boundary: a cache
# written at or after the most recent 7 AM is considered up to date.
DAILY_FETCH_HOUR = 7

# Outlook FlagStatus values (per-item): 0 = none, 1 = complete, 2 = flagged
OL_FLAG_MARKED = 2
# Restrict to mail items only (avoids tasks, flagged contacts, etc.)
OL_MAIL_ITEM = "IPM.Note"
# OlDefaultFolders.olFolderToDo — Outlook's search folder of all flagged items.
OL_FOLDER_TODO = 28


def _read_flagged_from_todo_folder(todo_folder):
    """Yield mail items the user flagged for follow-up from the To-Do folder.

    The To-Do folder already contains only flagged items and tasks, so there's
    no per-folder search to run. We still confirm each item is a mail message
    the user actively flagged (not a task, not completed, not a sender/system
    flag like the daily "To-Dos and News" digest, which has IsMarkedAsTask False).
    """
    for item in todo_folder.Items:
        try:
            if getattr(item, "MessageClass", "") != OL_MAIL_ITEM:
                continue
            if getattr(item, "FlagStatus", 0) != OL_FLAG_MARKED:
                continue  # skip completed or cleared flags
            if not getattr(item, "IsMarkedAsTask", False):
                continue  # skip sender/system flags (newsletters, reminders)
        except Exception:
            continue  # an item we can't inspect — skip it
        yield item


def get_flagged_emails(limit=50):
    """Return (emails, error). emails is a list of dicts with subject, sender,
    received — newest first. error is None on success, or a short string.

    Reads only Outlook's To-Do folder (all flagged items, newest first). This is
    a single light read rather than a full-mailbox folder sweep."""
    try:
        import win32com.client
    except ImportError:
        return [], "pywin32 not installed (pip install pywin32)"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as error:
        return [], f"Couldn't connect to Outlook ({error})"

    try:
        todo_folder = namespace.GetDefaultFolder(OL_FOLDER_TODO)
    except Exception as error:
        return [], f"Couldn't open Outlook's To-Do folder ({error})"

    # All To-Do items live in the same (primary) mailbox store, so one store id
    # serves every row for later open/unflag lookups.
    try:
        store_id = todo_folder.Store.StoreID
    except Exception:
        store_id = ""

    collected = []  # (received_datetime, email_dict)
    try:
        for item in _read_flagged_from_todo_folder(todo_folder):
            received_dt = getattr(item, "ReceivedTime", None)
            collected.append((received_dt, {
                "subject": (item.Subject or "(no subject)").strip(),
                "sender": (getattr(item, "SenderName", "") or "").strip(),
                "received": received_dt.strftime("%b %d") if received_dt else "",
                "entry_id": getattr(item, "EntryID", "") or "",
                "store_id": store_id,
            }))
    except Exception as error:
        return [], f"Couldn't read flagged emails ({error})"

    # Newest first; undated items go last.
    dated = sorted(
        (pair for pair in collected if pair[0] is not None),
        key=lambda pair: pair[0], reverse=True,
    )
    undated = [pair for pair in collected if pair[0] is None]
    ordered = dated + undated
    return [email for _, email in ordered[:limit]], None


# ── Opening / unflagging individual messages ─────────────────────────────────

def _get_item(entry_id, store_id):
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    if store_id:
        return namespace.GetItemFromID(entry_id, store_id)
    return namespace.GetItemFromID(entry_id)


def open_email(entry_id, store_id=""):
    """Open a specific message in Outlook. Returns (ok, error)."""
    if not entry_id:
        return False, "This email can't be opened (no identifier)."
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "pywin32 not installed"
    try:
        _get_item(entry_id, store_id).Display()
        return True, None
    except Exception as error:
        return False, f"Couldn't open email ({error})"


def _clear_flag(entry_id, store_id=""):
    """Clear the follow-up flag on one message in Outlook. Returns (ok, error).
    This is the actual COM write — used by the nightly batch run."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "pywin32 not installed"
    try:
        item = _get_item(entry_id, store_id)
        item.ClearTaskFlag()
        item.Save()
        return True, None
    except Exception as error:
        return False, f"Couldn't unflag email ({error})"


# ── Cache ─────────────────────────────────────────────────────────────────────

def read_cached_emails():
    """Return (emails, updated_label) from the cache file, instantly.
    emails is [] and updated_label is None if there's no cache yet."""
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("emails", []), data.get("updated")
    except Exception:
        return [], None


def _most_recent_fetch_boundary(now):
    """The most recent moment today's daily fetch hour (7 AM) has passed.
    If it's before 7 AM, the boundary is yesterday's 7 AM."""
    boundary = now.replace(hour=DAILY_FETCH_HOUR, minute=0, second=0, microsecond=0)
    if now < boundary:
        boundary -= timedelta(days=1)
    return boundary


def _cache_is_fresh():
    """True if the cache was refreshed at or after the most recent 7 AM mark,
    so there's no need to read Outlook again today."""
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        last_iso = data.get("updated_iso")
        if not last_iso:
            return False
        last_refresh = datetime.fromisoformat(last_iso)
        return last_refresh >= _most_recent_fetch_boundary(datetime.now())
    except Exception:
        return False


def refresh_cache(limit=50, force=False):
    """Do the once-a-day Outlook read and write the cache. Returns (emails, error).

    To keep Outlook responsive, this only reads Outlook when the cache is stale
    (older than today's 7 AM fetch). If the cache is already fresh and force is
    False, it returns the cached list immediately without touching Outlook — so
    the apps calling this during the day never poke Outlook. The scheduled 7 AM
    task passes force=True. On error the existing cache is left untouched."""
    if not force and _cache_is_fresh():
        return read_cached_emails()[0], None

    emails, error = get_flagged_emails(limit=limit)
    if error is None:
        now = datetime.now()
        try:
            with open(CACHE_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "emails": emails,
                    "updated": now.strftime("%b %d, %I:%M %p"),
                    "updated_iso": now.isoformat(),
                }, f, indent=2)
        except Exception:
            pass
    return emails, error


def remove_from_cache(entry_id):
    """Drop a single email from the cache by EntryID, keeping the rest."""
    if not entry_id:
        return
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return
    data["emails"] = [e for e in data.get("emails", []) if e.get("entry_id") != entry_id]
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


# ── Deferred unflagging ────────────────────────────────────────────────────────
# Unflagging writes to Outlook, which can make it sluggish. So the app only
# queues the request (and updates the cache so the row vanishes); the actual
# Outlook write happens once, in the nightly 10 PM batch.

def _load_pending_unflags():
    try:
        with open(PENDING_UNFLAGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_pending_unflags(pending):
    try:
        with open(PENDING_UNFLAGS_FILE, "w", encoding="utf-8") as f:
            json.dump(pending, f, indent=2)
    except Exception:
        pass


def queue_unflag(entry_id, store_id=""):
    """Queue an email to be unflagged in Outlook tonight. Removes it from the
    cache now (so it disappears from the app immediately) but does NOT touch
    Outlook. Returns (ok, error)."""
    if not entry_id:
        return False, "This email can't be unflagged (no identifier)."
    remove_from_cache(entry_id)
    pending = _load_pending_unflags()
    if not any(entry.get("entry_id") == entry_id for entry in pending):
        pending.append({"entry_id": entry_id, "store_id": store_id})
        _save_pending_unflags(pending)
    return True, None


def apply_pending_unflags():
    """Clear the flags for every queued email in Outlook (the nightly 10 PM job).
    Anything that fails is kept in the queue to retry next time.
    Returns (cleared_count, error)."""
    pending = _load_pending_unflags()
    if not pending:
        return 0, None
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return 0, "pywin32 not installed"

    still_pending = []
    cleared = 0
    for entry in pending:
        ok, _ = _clear_flag(entry.get("entry_id", ""), entry.get("store_id", ""))
        if ok:
            cleared += 1
        else:
            still_pending.append(entry)  # keep it to retry next run
    _save_pending_unflags(still_pending)
    return cleared, None


if __name__ == "__main__":
    import sys
    # Email subjects can contain emoji/Unicode the Windows console can't encode.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    # `--refresh`       : force the daily Outlook read + cache write (7 AM task).
    # `--apply-unflags` : write all queued unflags to Outlook (10 PM task).
    # (no args)         : just do a one-off live read and print it.
    if "--refresh" in sys.argv:
        emails, error = refresh_cache(force=True)
        if error:
            print("Refresh error:", error)
        else:
            print(f"Cache refreshed: {len(emails)} flagged email(s).")
    elif "--apply-unflags" in sys.argv:
        cleared, error = apply_pending_unflags()
        if error:
            print("Apply-unflags error:", error)
        else:
            print(f"Applied {cleared} queued unflag(s) to Outlook.")
    else:
        emails, error = get_flagged_emails()
        if error:
            print("Error:", error)
        elif not emails:
            print("No flagged emails.")
        else:
            print(f"{len(emails)} flagged email(s):")
            for email in emails:
                print(f"  [{email['received']}] {email['subject']} — {email['sender']}")

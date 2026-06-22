"""Read flagged emails from the local (classic) Outlook desktop app via COM.

This runs entirely on your machine against the Outlook profile you're already
signed into — no cloud credentials, no third-party service. It searches every
folder of your Exchange mailbox(es) and returns flagged messages, newest first,
so the morning briefing and task manager can summarize them.

Because the full-mailbox search is slow, results are cached to a small JSON file
(read_cached_emails / refresh_cache). The apps show the cache instantly and
refresh it in the background; a scheduled task keeps it warm when nothing's open.
"""

import json
import os
from datetime import datetime

CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "flagged_emails_cache.json")

# Outlook FlagStatus values (per-item): 0 = none, 1 = complete, 2 = flagged
OL_FLAG_MARKED = 2
# Restrict to mail items only (avoids meeting requests, etc.)
OL_MAIL_ITEM = "IPM.Note"

# OlExchangeStoreType: skip Public Folders (2) and non-Exchange stores (3),
# e.g. the conflicted .pst on OneDrive and the SharePoint Lists store.
NON_EXCHANGE_STORE_TYPES = (2, 3)

# Filtering on [FlagStatus] / PR_FLAG_STATUS is unreliable in Outlook and
# silently returns nothing. PR_TODO_ITEM_FLAGS (0x0E2B0003) is the property
# that actually reflects follow-up flags, so we narrow on that, then confirm
# each item is flagged (not completed) by reading FlagStatus directly.
DASL_HAS_TODO_FLAG = '@SQL="http://schemas.microsoft.com/mapi/proptag/0x0E2B0003" > 0'


def _is_exchange_store(store):
    """True for real Exchange mailboxes (primary, delegate, archive); False for
    the OneDrive .pst, SharePoint Lists, and public folders."""
    try:
        return store.ExchangeStoreType not in NON_EXCHANGE_STORE_TYPES
    except Exception:
        return False


def _walk_folders(folder):
    """Yield a folder and all of its subfolders, recursively."""
    yield folder
    try:
        subfolders = folder.Folders
    except Exception:
        return
    for subfolder in subfolders:
        yield from _walk_folders(subfolder)


def _read_flagged_from_folder(folder):
    """Yield mail items the user flagged for follow-up from one folder.

    Only items the user actively flagged (IsMarkedAsTask) are returned. Many
    bulk/newsletter emails (e.g. the daily "To-Dos and News" digest) arrive
    pre-flagged by the sender — those have IsMarkedAsTask False and are skipped.
    """
    flagged = folder.Items.Restrict(DASL_HAS_TODO_FLAG)
    for item in flagged:
        if getattr(item, "MessageClass", "") != OL_MAIL_ITEM:
            continue
        if getattr(item, "FlagStatus", 0) != OL_FLAG_MARKED:
            continue  # skip completed or cleared flags
        if not getattr(item, "IsMarkedAsTask", False):
            continue  # skip sender/system flags (newsletters, reminders)
        yield item


def get_flagged_emails(limit=5):
    """Return (emails, error). emails is a list of dicts with subject, sender,
    received — newest first. error is None on success, or a short string.

    Searches all folders of every Exchange mailbox. Non-Exchange stores (the
    conflicted OneDrive .pst, SharePoint Lists) are skipped, and any individual
    folder that errors is skipped rather than failing the whole read."""
    try:
        import win32com.client
    except ImportError:
        return [], "pywin32 not installed (pip install pywin32)"

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as error:
        return [], f"Couldn't connect to Outlook ({error})"

    collected = []  # (received_datetime, email_dict)
    exchange_stores = 0

    for store in namespace.Stores:
        if not _is_exchange_store(store):
            continue
        try:
            root = store.GetRootFolder()
        except Exception:
            continue
        exchange_stores += 1
        try:
            store_id = store.StoreID
        except Exception:
            store_id = ""

        for folder in _walk_folders(root):
            try:
                items = list(_read_flagged_from_folder(folder))
            except Exception:
                continue  # folder can't be read — skip it
            for item in items:
                received_dt = getattr(item, "ReceivedTime", None)
                collected.append((received_dt, {
                    "subject": (item.Subject or "(no subject)").strip(),
                    "sender": (getattr(item, "SenderName", "") or "").strip(),
                    "received": received_dt.strftime("%b %d") if received_dt else "",
                    "entry_id": getattr(item, "EntryID", "") or "",
                    "store_id": store_id,
                }))

    if exchange_stores == 0:
        return [], "No Exchange mailbox found"

    # Newest first across all folders; undated items go last.
    dated = sorted(
        (pair for pair in collected if pair[0] is not None),
        key=lambda pair: pair[0], reverse=True,
    )
    undated = [pair for pair in collected if pair[0] is None]
    ordered = dated + undated
    return [email for _, email in ordered[:limit]], None


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


def unflag_email(entry_id, store_id=""):
    """Clear the follow-up flag on a message. Returns (ok, error)."""
    if not entry_id:
        return False, "This email can't be unflagged (no identifier)."
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


def read_cached_emails():
    """Return (emails, updated_label) from the cache file, instantly.
    emails is [] and updated_label is None if there's no cache yet."""
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("emails", []), data.get("updated")
    except Exception:
        return [], None


def refresh_cache(limit=50):
    """Do the slow Outlook read and write the cache. Returns (emails, error).
    On error the existing cache is left untouched."""
    emails, error = get_flagged_emails(limit=limit)
    if error is None:
        try:
            with open(CACHE_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "emails": emails,
                    "updated": datetime.now().strftime("%b %d, %I:%M %p"),
                }, f, indent=2)
        except Exception:
            pass
    return emails, error


def remove_from_cache(entry_id):
    """Drop a single email from the cache by EntryID, keeping the rest.
    Used after unflagging so the list updates without a full Outlook re-read."""
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


if __name__ == "__main__":
    import sys
    # Email subjects can contain emoji/Unicode the Windows console can't encode.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    # `--refresh` updates the cache (used by the scheduled task); otherwise
    # just do a one-off live read and print it.
    if "--refresh" in sys.argv:
        emails, error = refresh_cache()
        if error:
            print("Refresh error:", error)
        else:
            print(f"Cache refreshed: {len(emails)} flagged email(s).")
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

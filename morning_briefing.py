import customtkinter as ctk
import json
import os
import queue
import random
import subprocess
import sys
import threading
from datetime import date

try:
    import get_flagged_emails as flagged_email_reader
except ImportError:
    flagged_email_reader = None

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

TASKS_FILE = "tasks.json"
PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}
PRIORITY_COLORS = {
    "High": "#ff6b6b",
    "Medium": "#ffd93d",
    "Low": "#6bcb77",
}

QUOTES = [
    "The secret of getting ahead is getting started. — Mark Twain",
    "Do what you can, with what you have, where you are. — Theodore Roosevelt",
    "Small steps every day lead to big results.",
    "Focus on progress, not perfection.",
    "You don't have to be great to start, but you have to start to be great.",
    "One task at a time. Finish what you begin.",
    "A goal without a plan is just a wish. — Antoine de Saint-Exupéry",
]


def load_tasks():
    if not os.path.exists(TASKS_FILE):
        return []
    with open(TASKS_FILE, "r") as f:
        return json.load(f)


def top_pending_tasks(tasks, limit=3):
    pending = [t for t in tasks if not t.get("done", False)]
    return sorted(pending, key=lambda t: PRIORITY_ORDER.get(t.get("priority", "Low"), 2))[:limit]


class MorningBriefingWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        tasks = load_tasks()
        self._top_tasks = top_pending_tasks(tasks)
        self._today = date.today().strftime("%A, %B %d, %Y")
        self._quote = random.choice(QUOTES)
        self._email_queue = queue.Queue()

        self.title("Morning Briefing")
        self.resizable(False, False)
        self.attributes("-topmost", True)

        self._build_ui()
        self.after(50, self._center_and_raise)
        # Read Outlook off the main thread so the popup appears immediately.
        self._start_email_load()

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, corner_radius=0, fg_color="#111827")
        header.pack(fill="x")

        ctk.CTkLabel(
            header,
            text="Good morning",
            font=ctk.CTkFont(size=13),
            text_color="#6b7280",
        ).pack(pady=(22, 2))

        ctk.CTkLabel(
            header,
            text=self._today,
            font=ctk.CTkFont(size=21, weight="bold"),
        ).pack(pady=(0, 22))

        # ── Body ──────────────────────────────────────────────────────────────
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=26, pady=0)

        # Quote
        ctk.CTkLabel(
            body,
            text="QUOTE OF THE MORNING",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="#6b7280",
            anchor="w",
        ).pack(anchor="w", pady=(20, 6))

        ctk.CTkLabel(
            body,
            text=f'"{self._quote}"',
            font=ctk.CTkFont(size=13),
            wraplength=388,
            justify="left",
            text_color="#9ca3af",
        ).pack(anchor="w")

        ctk.CTkFrame(body, height=1, fg_color="#2d2d3a").pack(fill="x", pady=(18, 0))

        # Tasks
        ctk.CTkLabel(
            body,
            text="TOP TASKS FOR TODAY",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="#6b7280",
            anchor="w",
        ).pack(anchor="w", pady=(18, 10))

        if not self._top_tasks:
            ctk.CTkLabel(
                body,
                text="No pending tasks — enjoy your day!",
                font=ctk.CTkFont(size=13),
                text_color="#6b7280",
            ).pack(anchor="w", pady=(0, 6))
        else:
            for task in self._top_tasks:
                self._render_task_row(body, task)

        ctk.CTkFrame(body, height=1, fg_color="#2d2d3a").pack(fill="x", pady=(18, 0))

        # Flagged emails
        ctk.CTkLabel(
            body,
            text="FLAGGED EMAILS",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color="#6b7280",
            anchor="w",
        ).pack(anchor="w", pady=(18, 10))

        self.email_container = ctk.CTkFrame(body, fg_color="transparent")
        self.email_container.pack(fill="x")
        self._show_email_message("Loading flagged emails…")

        ctk.CTkFrame(body, height=1, fg_color="#2d2d3a").pack(fill="x", pady=(18, 0))

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x", pady=(14, 22))

        ctk.CTkButton(
            btn_row,
            text="Open Task Manager",
            height=38,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._open_task_manager,
        ).pack(side="left")

        ctk.CTkButton(
            btn_row,
            text="Dismiss",
            height=38,
            width=90,
            fg_color="transparent",
            border_width=1,
            border_color="#3d3d50",
            text_color="#9ca3af",
            hover_color="#1f1f2e",
            font=ctk.CTkFont(size=13),
            command=self.destroy,
        ).pack(side="right")

    def _render_task_row(self, parent, task):
        priority = task.get("priority", "Low")
        color = PRIORITY_COLORS.get(priority, "#888")

        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 8))

        ctk.CTkLabel(
            row,
            text="●",
            font=ctk.CTkFont(size=10),
            text_color=color,
            width=20,
        ).pack(side="left")

        ctk.CTkLabel(
            row,
            text=task["title"],
            font=ctk.CTkFont(size=13),
            anchor="w",
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            row,
            text=priority,
            font=ctk.CTkFont(size=11),
            text_color=color,
            width=56,
            anchor="e",
        ).pack(side="right")

        project = task.get("project", "General")
        ctk.CTkLabel(
            row,
            text=project,
            font=ctk.CTkFont(size=11),
            text_color="#6b7280",
            anchor="e",
        ).pack(side="right", padx=(0, 10))

    # ── Flagged emails (loaded in the background) ─────────────────────────────

    def _start_email_load(self):
        if flagged_email_reader is None:
            self._populate_emails([], "Email reader unavailable", None)
            return
        # Show cached emails instantly, then refresh from Outlook in the background.
        emails, updated = flagged_email_reader.read_cached_emails()
        if emails or updated:
            self._populate_emails(emails[:5], None, updated)
        threading.Thread(target=self._load_emails_worker, daemon=True).start()
        self.after(150, self._poll_email_result)

    def _load_emails_worker(self):
        pythoncom = None
        try:
            import pythoncom  # COM must be initialized in this thread
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None
        try:
            emails, error = flagged_email_reader.refresh_cache(limit=50)
        except Exception as worker_error:
            emails, error = [], f"Couldn't read Outlook ({worker_error})"
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        _, updated = flagged_email_reader.read_cached_emails()
        self._email_queue.put((emails, error, updated))

    def _poll_email_result(self):
        try:
            emails, error, updated = self._email_queue.get_nowait()
        except queue.Empty:
            self.after(150, self._poll_email_result)
            return
        # On a transient read error, keep the last good cache.
        if error:
            cached, cached_updated = flagged_email_reader.read_cached_emails()
            if cached:
                emails, updated, error = cached, cached_updated, None
        self._populate_emails(emails[:5], error, updated)

    def _populate_emails(self, emails, error, updated):
        for widget in self.email_container.winfo_children():
            widget.destroy()
        if emails:
            for email in emails:
                self._render_email_row(self.email_container, email)
            if updated:
                ctk.CTkLabel(
                    self.email_container, text=f"Updated {updated}",
                    font=ctk.CTkFont(size=10), text_color="#555555", anchor="w",
                ).pack(anchor="w", pady=(2, 0))
        elif error:
            self._show_email_message(error)
        else:
            self._show_email_message("No flagged emails — inbox is clear!")
        # Content height changed, so resize the popup to fit and re-center.
        self.after(10, self._center_and_raise)

    def _show_email_message(self, text):
        ctk.CTkLabel(
            self.email_container, text=text, font=ctk.CTkFont(size=13),
            text_color="#6b7280", wraplength=388, justify="left",
        ).pack(anchor="w", pady=(0, 6))

    def _render_email_row(self, parent, email):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(0, 8))

        envelope = ctk.CTkLabel(
            row,
            text="✉",
            font=ctk.CTkFont(size=12),
            text_color="#f5a623",
            width=20,
        )
        envelope.pack(side="left")

        text_col = ctk.CTkFrame(row, fg_color="transparent")
        text_col.pack(side="left", fill="x", expand=True)

        subject = ctk.CTkLabel(
            text_col,
            text=email["subject"],
            font=ctk.CTkFont(size=13),
            anchor="w",
            justify="left",
            wraplength=300,
        )
        subject.pack(anchor="w")

        detail = email["sender"]
        if email["received"]:
            detail = f"{detail} · {email['received']}" if detail else email["received"]
        detail_label = None
        if detail:
            detail_label = ctk.CTkLabel(
                text_col,
                text=detail,
                font=ctk.CTkFont(size=11),
                text_color="#6b7280",
                anchor="w",
            )
            detail_label.pack(anchor="w")

        # Click anywhere on the row to open the message in Outlook.
        for widget in [row, envelope, text_col, subject] + ([detail_label] if detail_label else []):
            widget.configure(cursor="hand2")
            widget.bind("<Button-1>", lambda _event, em=email: self._open_email(em))

    def _open_email(self, email):
        if flagged_email_reader is None:
            return
        flagged_email_reader.open_email(email.get("entry_id", ""), email.get("store_id", ""))

    def _open_task_manager(self):
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "task_manager.py")
        subprocess.Popen([sys.executable, script])
        self.destroy()

    def _center_and_raise(self):
        self.update_idletasks()
        # Use the requested (natural) size so the window grows to fit content
        # even after an explicit geometry was set on an earlier call.
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.lift()
        self.focus_force()


if __name__ == "__main__":
    app = MorningBriefingWindow()
    app.mainloop()

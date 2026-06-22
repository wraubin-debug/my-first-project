import customtkinter as ctk
from tkinter import messagebox
import json
import os

try:
    import generate_mobile_view
except ImportError:
    generate_mobile_view = None

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

TASKS_FILE = "tasks.json"
PROJECTS_FILE = "projects.json"
DEFAULT_PROJECT = "General"
PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}
PRIORITY_COLORS = {
    "High": "#ff6b6b",
    "Medium": "#ffd93d",
    "Low": "#6bcb77",
}


def load_tasks():
    if not os.path.exists(TASKS_FILE):
        return []
    with open(TASKS_FILE, "r") as f:
        tasks = json.load(f)
    # Backfill fields for tasks that predate the UI / projects.
    next_id = max((t.get("id", 0) for t in tasks), default=0) + 1
    for task in tasks:
        if "id" not in task:
            task["id"] = next_id
            next_id += 1
        if "project" not in task:
            task["project"] = DEFAULT_PROJECT
    return tasks


def save_tasks(tasks):
    with open(TASKS_FILE, "w") as f:
        json.dump(tasks, f, indent=2)
    regenerate_mobile_view()


def regenerate_mobile_view():
    """Refresh the phone-friendly HTML in OneDrive. Never let a failure here
    (e.g. OneDrive offline) interrupt the desktop app."""
    if generate_mobile_view is None:
        return
    try:
        generate_mobile_view.generate()
    except Exception as error:
        print("Mobile view update skipped:", error)


def load_projects(tasks):
    projects = []
    if os.path.exists(PROJECTS_FILE):
        with open(PROJECTS_FILE, "r") as f:
            projects = json.load(f)
    # Make sure every project referenced by a task exists in the list.
    for task in tasks:
        project = task.get("project", DEFAULT_PROJECT)
        if project not in projects:
            projects.append(project)
    if not projects:
        projects = [DEFAULT_PROJECT]
    return projects


def save_projects(projects):
    with open(PROJECTS_FILE, "w") as f:
        json.dump(projects, f, indent=2)


class TaskManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Task Manager")
        self.geometry("880x600")
        self.minsize(740, 500)

        self.tasks = load_tasks()
        self.projects = load_projects(self.tasks)
        self.current_project = self.projects[0]

        self._build_ui()
        self._refresh_projects()
        self._refresh_task_list()

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── Left sidebar ──────────────────────────────────────────────────────
        sidebar = ctk.CTkFrame(self, width=250, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.grid_rowconfigure(2, weight=1)
        sidebar.grid_columnconfigure(0, weight=1)

        # Projects header + new-project button
        proj_header = ctk.CTkFrame(sidebar, fg_color="transparent")
        proj_header.grid(row=0, column=0, sticky="ew", padx=18, pady=(24, 8))
        proj_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            proj_header,
            text="Projects",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            proj_header,
            text="+ New",
            width=60,
            height=28,
            font=ctk.CTkFont(size=12),
            command=self._new_project,
        ).grid(row=0, column=1, sticky="e")

        # Scrollable project list
        self.project_scroll = ctk.CTkScrollableFrame(sidebar, fg_color="transparent")
        self.project_scroll.grid(row=2, column=0, sticky="nsew", padx=10)
        self.project_scroll.grid_columnconfigure(0, weight=1)

        # Add-task form
        form = ctk.CTkFrame(sidebar, corner_radius=8)
        form.grid(row=3, column=0, sticky="ew", padx=12, pady=12)
        form.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            form, text="Add Task", font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 8))

        self.title_entry = ctk.CTkEntry(form, placeholder_text="Task title...", height=34)
        self.title_entry.grid(row=1, column=0, sticky="ew", padx=12)
        self.title_entry.bind("<Return>", lambda _: self._add_task())

        self.priority_var = ctk.StringVar(value="High")
        ctk.CTkOptionMenu(
            form, values=["High", "Medium", "Low"], variable=self.priority_var, height=34
        ).grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 0))

        ctk.CTkButton(
            form, text="+ Add Task", height=36,
            font=ctk.CTkFont(size=13, weight="bold"), command=self._add_task,
        ).grid(row=3, column=0, sticky="ew", padx=12, pady=12)

        # ── Right content ─────────────────────────────────────────────────────
        right = ctk.CTkFrame(self, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=16, pady=16)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        topbar = ctk.CTkFrame(right, fg_color="transparent")
        topbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        topbar.grid_columnconfigure(1, weight=1)

        self.project_title = ctk.CTkLabel(
            topbar, text="", font=ctk.CTkFont(size=18, weight="bold")
        )
        self.project_title.grid(row=0, column=0, sticky="w")

        filters = ctk.CTkFrame(topbar, fg_color="transparent")
        filters.grid(row=0, column=2, sticky="e")
        self.filter_var = ctk.StringVar(value="All")
        for label in ["All", "Active", "Done"]:
            ctk.CTkRadioButton(
                filters, text=label, variable=self.filter_var, value=label,
                command=self._refresh_task_list,
            ).pack(side="left", padx=(12, 0))

        self.task_scroll = ctk.CTkScrollableFrame(right)
        self.task_scroll.grid(row=1, column=0, sticky="nsew")
        self.task_scroll.grid_columnconfigure(0, weight=1)

        self.stats_label = ctk.CTkLabel(
            right, text="", font=ctk.CTkFont(size=12), text_color="gray", anchor="w"
        )
        self.stats_label.grid(row=2, column=0, sticky="w", pady=(8, 0))

    # ── Project actions ───────────────────────────────────────────────────────

    def _new_project(self):
        dialog = ctk.CTkInputDialog(text="New project name:", title="New Project")
        name = dialog.get_input()
        if name is None:
            return
        name = name.strip()
        if not name:
            return
        if name in self.projects:
            messagebox.showinfo("Project exists", f'"{name}" already exists.')
            self._select_project(name)
            return
        self.projects.append(name)
        save_projects(self.projects)
        self._select_project(name)

    def _delete_project(self, name):
        if len(self.projects) == 1:
            messagebox.showinfo(
                "Can't delete", "You must keep at least one project."
            )
            return
        task_count = sum(1 for t in self.tasks if t.get("project") == name)
        message = f'Delete project "{name}"?'
        if task_count:
            message += f"\nThis will also delete its {task_count} task(s)."
        if not messagebox.askyesno("Delete project", message):
            return
        self.tasks = [t for t in self.tasks if t.get("project") != name]
        self.projects.remove(name)
        save_tasks(self.tasks)
        save_projects(self.projects)
        if self.current_project == name:
            self.current_project = self.projects[0]
        self._refresh_projects()
        self._refresh_task_list()

    def _select_project(self, name):
        self.current_project = name
        self._refresh_projects()
        self._refresh_task_list()

    # ── Task actions ──────────────────────────────────────────────────────────

    def _add_task(self):
        title = self.title_entry.get().strip()
        if not title:
            return
        next_id = max((t.get("id", 0) for t in self.tasks), default=0) + 1
        self.tasks.append({
            "id": next_id,
            "title": title,
            "priority": self.priority_var.get(),
            "done": False,
            "project": self.current_project,
        })
        save_tasks(self.tasks)
        self.title_entry.delete(0, "end")
        self._refresh_task_list()
        self._refresh_projects()  # update task counts

    def _toggle_done(self, task_id):
        for task in self.tasks:
            if task.get("id") == task_id:
                task["done"] = not task["done"]
                break
        save_tasks(self.tasks)
        self._refresh_task_list()

    def _delete_task(self, task_id):
        self.tasks = [t for t in self.tasks if t.get("id") != task_id]
        save_tasks(self.tasks)
        self._refresh_task_list()
        self._refresh_projects()

    # ── Rendering ─────────────────────────────────────────────────────────────

    def _refresh_projects(self):
        for widget in self.project_scroll.winfo_children():
            widget.destroy()

        for name in self.projects:
            count = sum(
                1 for t in self.tasks
                if t.get("project") == name and not t["done"]
            )
            is_current = name == self.current_project

            row = ctk.CTkFrame(self.project_scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            row.grid_columnconfigure(0, weight=1)

            label = f"{name}  ({count})" if count else name
            ctk.CTkButton(
                row,
                text=label,
                anchor="w",
                height=34,
                fg_color="#1f6aa5" if is_current else "transparent",
                hover_color="#144870" if is_current else "#2b2b2b",
                text_color="white" if is_current else "#bbbbbb",
                command=lambda n=name: self._select_project(n),
            ).grid(row=0, column=0, sticky="ew")

            ctk.CTkButton(
                row,
                text="✕",
                width=30,
                height=34,
                fg_color="transparent",
                hover_color="#7b2d2d",
                text_color="#888888",
                command=lambda n=name: self._delete_project(n),
            ).grid(row=0, column=1, padx=(4, 0))

    def _refresh_task_list(self):
        for widget in self.task_scroll.winfo_children():
            widget.destroy()

        self.project_title.configure(text=self.current_project)

        mode = self.filter_var.get()
        visible = [
            t for t in self.tasks
            if t.get("project") == self.current_project
            and (mode == "All"
                 or (mode == "Active" and not t["done"])
                 or (mode == "Done" and t["done"]))
        ]
        visible.sort(
            key=lambda t: (t["done"], PRIORITY_ORDER.get(t.get("priority", "Low"), 3))
        )

        if not visible:
            ctk.CTkLabel(
                self.task_scroll, text="No tasks here.",
                text_color="gray", font=ctk.CTkFont(size=14),
            ).pack(pady=40)
        else:
            for task in visible:
                self._render_task_row(task)

        project_tasks = [t for t in self.tasks if t.get("project") == self.current_project]
        done_count = sum(1 for t in project_tasks if t["done"])
        self.stats_label.configure(
            text=f"{done_count} of {len(project_tasks)} tasks done in this project"
        )

    def _render_task_row(self, task):
        priority = task.get("priority", "Low")
        color = PRIORITY_COLORS.get(priority, "#888")
        is_done = task["done"]
        task_id = task.get("id")

        row = ctk.CTkFrame(self.task_scroll, corner_radius=6, height=42)
        row.pack(fill="x", pady=2)
        row.pack_propagate(False)

        # Priority dot
        ctk.CTkLabel(
            row, text="●", font=ctk.CTkFont(size=12),
            text_color="#555555" if is_done else color, width=24,
        ).pack(side="left", padx=(10, 0))

        # Title
        ctk.CTkLabel(
            row,
            text=task["title"],
            font=ctk.CTkFont(size=13, overstrike=is_done),
            text_color="gray" if is_done else "white",
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(2, 8))

        # Delete (rightmost)
        ctk.CTkButton(
            row, text="✕", width=32, height=28,
            fg_color="transparent", hover_color="#7b2d2d", text_color="#888888",
            command=lambda tid=task_id: self._delete_task(tid),
        ).pack(side="right", padx=(0, 8))

        # Done toggle
        ctk.CTkButton(
            row, text="↺" if is_done else "✓", width=32, height=28,
            fg_color="#555555" if is_done else "#2d6a4f",
            hover_color="#777777" if is_done else "#40916c",
            command=lambda tid=task_id: self._toggle_done(tid),
        ).pack(side="right", padx=(0, 4))

        # Priority text
        ctk.CTkLabel(
            row, text=priority, font=ctk.CTkFont(size=11),
            text_color="gray" if is_done else color, width=58, anchor="e",
        ).pack(side="right", padx=(0, 6))


if __name__ == "__main__":
    app = TaskManagerApp()
    app.mainloop()

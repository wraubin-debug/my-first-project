"""Generate a self-contained HTML dashboard of your tasks for phone viewing.

The page has all task data baked in (no separate files to fetch), so it opens
and renders directly from the OneDrive app on a phone. It is read-only — edit
tasks in the desktop app, which regenerates this file on every change.
"""

import json
import os
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
TASKS_FILE = SCRIPT_DIR / "tasks.json"
PROJECTS_FILE = SCRIPT_DIR / "projects.json"
DEFAULT_PROJECT = "General"
OUTPUT_NAME = "tasks.html"

PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}


def load_tasks():
    if not TASKS_FILE.exists():
        return []
    with open(TASKS_FILE, "r") as f:
        tasks = json.load(f)
    for task in tasks:
        if "project" not in task:
            task["project"] = DEFAULT_PROJECT
    return tasks


def load_projects(tasks):
    projects = []
    if PROJECTS_FILE.exists():
        with open(PROJECTS_FILE, "r") as f:
            projects = json.load(f)
    for task in tasks:
        project = task.get("project", DEFAULT_PROJECT)
        if project not in projects:
            projects.append(project)
    if not projects:
        projects = [DEFAULT_PROJECT]
    return projects


def find_output_dir():
    """Prefer the Accenture (commercial) OneDrive, then any OneDrive, so the
    file syncs to the phone. Falls back to the script folder if none is found."""
    home = Path.home()
    candidates = sorted(home.glob("OneDrive - *")) + [home / "OneDrive"]
    for base in candidates:
        if base.is_dir():
            output_dir = base / "TaskManager"
            output_dir.mkdir(exist_ok=True)
            return output_dir
    return SCRIPT_DIR


def build_html(tasks, projects):
    tasks_sorted = sorted(
        tasks,
        key=lambda t: (t.get("done", False),
                       PRIORITY_ORDER.get(t.get("priority", "Low"), 3)),
    )
    payload = {
        "tasks": tasks_sorted,
        "projects": projects,
        "updated": datetime.now().strftime("%A, %B %d %Y at %I:%M %p"),
    }
    data_json = json.dumps(payload)

    return """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>My Tasks</title>
<style>
  * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  body {
    margin: 0; padding: 0;
    background: #0f1117; color: #e5e7eb;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  }
  header {
    padding: 22px 18px 14px;
    background: #111827;
    position: sticky; top: 0; z-index: 10;
    border-bottom: 1px solid #1f2430;
  }
  h1 { margin: 0; font-size: 22px; font-weight: 700; }
  .updated { margin-top: 4px; font-size: 12px; color: #6b7280; }
  .filters { display: flex; gap: 8px; margin-top: 14px; }
  .chip {
    flex: 1; text-align: center;
    padding: 8px 0; border-radius: 8px;
    background: #1f2430; color: #9ca3af;
    font-size: 13px; font-weight: 600;
    border: none; cursor: pointer;
  }
  .chip.active { background: #1f6aa5; color: #fff; }
  main { padding: 14px 14px 40px; }
  .project { margin-bottom: 22px; }
  .project-head {
    display: flex; align-items: baseline; gap: 8px;
    margin: 0 4px 10px; padding-bottom: 6px;
    border-bottom: 1px solid #1f2430;
  }
  .project-name { font-size: 15px; font-weight: 700; }
  .project-count { font-size: 12px; color: #6b7280; }
  .task {
    display: flex; align-items: center; gap: 10px;
    background: #171b24; border-radius: 8px;
    padding: 12px 14px; margin-bottom: 7px;
  }
  .dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
  .title { flex: 1; font-size: 14px; line-height: 1.3; }
  .task.done .title { text-decoration: line-through; color: #6b7280; }
  .task.done .dot { background: #555 !important; }
  .priority { font-size: 11px; font-weight: 600; flex-shrink: 0; }
  .task.done .priority { color: #6b7280; }
  .empty { text-align: center; color: #6b7280; padding: 40px 0; font-size: 14px; }
</style>
</head>
<body>
<header>
  <h1>My Tasks</h1>
  <div class="updated" id="updated"></div>
  <div class="filters">
    <button class="chip active" data-filter="All">All</button>
    <button class="chip" data-filter="Active">Active</button>
    <button class="chip" data-filter="Done">Done</button>
  </div>
</header>
<main id="content"></main>

<script>
const DATA = __DATA__;
const COLORS = { High: "#ff6b6b", Medium: "#ffd93d", Low: "#6bcb77" };
let filter = "All";

document.getElementById("updated").textContent = "Updated " + DATA.updated;

function visibleTasks() {
  return DATA.tasks.filter(t =>
    filter === "All" ||
    (filter === "Active" && !t.done) ||
    (filter === "Done" && t.done)
  );
}

function render() {
  const content = document.getElementById("content");
  content.innerHTML = "";
  const tasks = visibleTasks();

  if (tasks.length === 0) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "No tasks to show.";
    content.appendChild(empty);
    return;
  }

  DATA.projects.forEach(projectName => {
    const projectTasks = tasks.filter(t => t.project === projectName);
    if (projectTasks.length === 0) return;

    const section = document.createElement("div");
    section.className = "project";

    const head = document.createElement("div");
    head.className = "project-head";
    const name = document.createElement("span");
    name.className = "project-name";
    name.textContent = projectName;
    const count = document.createElement("span");
    count.className = "project-count";
    const active = projectTasks.filter(t => !t.done).length;
    count.textContent = active + " active";
    head.appendChild(name);
    head.appendChild(count);
    section.appendChild(head);

    projectTasks.forEach(t => {
      const row = document.createElement("div");
      row.className = "task" + (t.done ? " done" : "");

      const dot = document.createElement("span");
      dot.className = "dot";
      dot.style.background = COLORS[t.priority] || "#888";

      const title = document.createElement("span");
      title.className = "title";
      title.textContent = t.title;

      const pri = document.createElement("span");
      pri.className = "priority";
      pri.style.color = COLORS[t.priority] || "#888";
      pri.textContent = t.priority || "Low";

      row.appendChild(dot);
      row.appendChild(title);
      row.appendChild(pri);
      section.appendChild(row);
    });

    content.appendChild(section);
  });
}

document.querySelectorAll(".chip").forEach(chip => {
  chip.addEventListener("click", () => {
    document.querySelectorAll(".chip").forEach(c => c.classList.remove("active"));
    chip.classList.add("active");
    filter = chip.dataset.filter;
    render();
  });
});

render();
</script>
</body>
</html>
""".replace("__DATA__", data_json)


def generate():
    tasks = load_tasks()
    projects = load_projects(tasks)
    html = build_html(tasks, projects)
    output_dir = find_output_dir()
    output_path = output_dir / OUTPUT_NAME
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return output_path


if __name__ == "__main__":
    path = generate()
    print("Mobile view generated:")
    print(f"  {path}")
    print()
    print("Open this file from the OneDrive app on your phone to view your tasks.")

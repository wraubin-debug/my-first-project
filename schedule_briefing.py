import os
import subprocess
import sys
from pathlib import Path


def find_pythonw():
    """Use pythonw.exe so no console window flashes when tasks run."""
    pythonw = Path(sys.executable).parent / "pythonw.exe"
    return str(pythonw) if pythonw.exists() else sys.executable


def get_desktop_path():
    """Get the real Desktop path — handles OneDrive-redirected Desktops."""
    result = subprocess.run(
        ["powershell", "-Command", "[Environment]::GetFolderPath('Desktop')"],
        capture_output=True, text=True,
    )
    desktop = result.stdout.strip()
    if desktop and os.path.isdir(desktop):
        return desktop
    return str(Path.home() / "Desktop")


def register_email_cache_task(pythonw_path, reader_script):
    # Fetch flagged emails from Outlook just once a day, at 7 AM, so Outlook
    # isn't repeatedly poked during the day. (This overwrites the old
    # every-15-minutes task of the same name, if it exists.)
    task_name = "Flagged Email Cache Refresh"
    command = f'"{pythonw_path}" "{reader_script}" --refresh'

    print("Registering daily flagged-email fetch (7 AM)...")

    result = subprocess.run(
        [
            "schtasks", "/create",
            "/tn", task_name,
            "/tr", command,
            "/sc", "daily",
            "/st", "07:00",
            "/f",            # overwrite if it already exists
        ],
        capture_output=True, text=True,
    )

    if result.returncode == 0:
        print(f"  '{task_name}' registered — fetches flagged emails daily at 7:00 AM.")
    else:
        print("  ERROR: Could not register daily fetch task.")
        print(f"  {result.stderr.strip()}")


def register_assistant_briefing_task(pythonw_path, assistant_script):
    # Generate the Claude-powered daily briefing at 7:05 AM — just after the
    # 7:00 AM flagged-email fetch, so the briefing sees a fresh flagged cache.
    task_name = "Personal Assistant Briefing"
    command = f'"{pythonw_path}" "{assistant_script}" --refresh'

    print("Registering daily assistant briefing (7:05 AM)...")

    result = subprocess.run(
        [
            "schtasks", "/create",
            "/tn", task_name,
            "/tr", command,
            "/sc", "daily",
            "/st", "07:05",
            "/f",            # overwrite if it already exists
        ],
        capture_output=True, text=True,
    )

    if result.returncode == 0:
        print(f"  '{task_name}' registered — builds your briefing daily at 7:05 AM.")
    else:
        print("  ERROR: Could not register assistant briefing task.")
        print(f"  {result.stderr.strip()}")


def register_memory_update_task(pythonw_path, assistant_script):
    # Fold the day's email into the rolling topic memory at 10:05 PM — just after
    # the 10 PM apply-unflags batch, so the two nightly jobs don't overlap.
    task_name = "Assistant Memory Update"
    command = f'"{pythonw_path}" "{assistant_script}" --update-memory'

    print("Registering nightly memory update (10:05 PM)...")

    result = subprocess.run(
        [
            "schtasks", "/create",
            "/tn", task_name,
            "/tr", command,
            "/sc", "daily",
            "/st", "22:05",
            "/f",            # overwrite if it already exists
        ],
        capture_output=True, text=True,
    )

    if result.returncode == 0:
        print(f"  '{task_name}' registered — updates topic memory daily at 10:05 PM.")
    else:
        print("  ERROR: Could not register memory update task.")
        print(f"  {result.stderr.strip()}")


def register_apply_unflags_task(pythonw_path, reader_script):
    # Write any unflags you made in the app back to Outlook in one nightly
    # batch at 10 PM, so Outlook isn't poked while you're using it.
    task_name = "Apply Unflagged Emails"
    command = f'"{pythonw_path}" "{reader_script}" --apply-unflags'

    print("Registering nightly apply-unflags task (10 PM)...")

    result = subprocess.run(
        [
            "schtasks", "/create",
            "/tn", task_name,
            "/tr", command,
            "/sc", "daily",
            "/st", "22:00",
            "/f",            # overwrite if it already exists
        ],
        capture_output=True, text=True,
    )

    if result.returncode == 0:
        print(f"  '{task_name}' registered — applies queued unflags daily at 10:00 PM.")
    else:
        print("  ERROR: Could not register apply-unflags task.")
        print(f"  {result.stderr.strip()}")


def create_desktop_shortcut(pythonw_path, task_manager_script, script_dir, desktop):
    shortcut_path = os.path.join(desktop, "Task Manager.lnk")

    print("Creating desktop shortcut...")

    # Write a temp .ps1 file so paths with spaces are handled cleanly.
    temp_ps1 = os.path.join(script_dir, "_temp_shortcut.ps1")
    ps_content = f'''
$WshShell = New-Object -comObject WScript.Shell
$s = $WshShell.CreateShortcut("{shortcut_path}")
$s.TargetPath     = "{pythonw_path}"
$s.Arguments      = '"{task_manager_script}"'
$s.WorkingDirectory = "{script_dir}"
$s.IconLocation   = "{pythonw_path},0"
$s.Description    = "Open Morning Task Manager"
$s.Save()
'''

    try:
        with open(temp_ps1, "w", encoding="utf-8") as f:
            f.write(ps_content)

        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", temp_ps1],
            capture_output=True, text=True,
        )

        if result.returncode == 0:
            print(f"  Shortcut created: {shortcut_path}")
        else:
            print("  ERROR: Could not create desktop shortcut.")
            print(f"  {result.stderr.strip()}")
    finally:
        if os.path.exists(temp_ps1):
            os.remove(temp_ps1)


def main():
    script_dir = str(Path(__file__).parent.resolve())
    task_manager_script = os.path.join(script_dir, "task_manager.py")
    reader_script       = os.path.join(script_dir, "get_flagged_emails.py")
    assistant_script    = os.path.join(script_dir, "personal_assistant.py")

    for path, name in [
        (task_manager_script, "task_manager.py"),
        (reader_script,       "get_flagged_emails.py"),
        (assistant_script,    "personal_assistant.py"),
    ]:
        if not os.path.exists(path):
            print(f"Error: {name} not found at {path}")
            sys.exit(1)

    pythonw_path = find_pythonw()
    desktop      = get_desktop_path()

    print("Morning Agent Setup")
    print("=" * 44)
    print(f"Python:  {pythonw_path}")
    print(f"Scripts: {script_dir}")
    print(f"Desktop: {desktop}")
    print()

    register_email_cache_task(pythonw_path, reader_script)
    register_assistant_briefing_task(pythonw_path, assistant_script)
    register_apply_unflags_task(pythonw_path, reader_script)
    register_memory_update_task(pythonw_path, assistant_script)
    create_desktop_shortcut(pythonw_path, task_manager_script, script_dir, desktop)

    print()
    print("All done!")
    print("  Email fetch:      reads flagged emails from Outlook daily at 7:00 AM")
    print("  Assistant brief:  Claude builds your daily plan at 7:05 AM")
    print("  Apply unflags:    writes your unflags back to Outlook daily at 10:00 PM")
    print("  Memory update:    folds the day's email into topic memory at 10:05 PM")
    print("  Task Manager:     shortcut on your Desktop")
    print()
    print("These run while the screen is locked, as long as you're still logged in")
    print("and the machine is awake (not asleep/shut down) at those times.")
    print()
    print("To change a time: Task Scheduler > Task Scheduler Library > pick the task")
    print("To remove a task: schtasks /delete /tn \"Flagged Email Cache Refresh\" /f")
    print("  (likewise for \"Personal Assistant Briefing\",")
    print("   \"Apply Unflagged Emails\", and \"Assistant Memory Update\")")


if __name__ == "__main__":
    main()

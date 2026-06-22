import os
import subprocess
import sys
from pathlib import Path


def find_pythonw():
    """Use pythonw.exe so no console window flashes at 9 AM."""
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


def register_scheduled_task(pythonw_path, briefing_script):
    task_name = "Morning Briefing"
    command = f'"{pythonw_path}" "{briefing_script}"'

    print("Registering Windows scheduled task...")

    result = subprocess.run(
        [
            "schtasks", "/create",
            "/tn", task_name,
            "/tr", command,
            "/sc", "daily",
            "/st", "09:00",
            "/f",           # overwrite if the task already exists
        ],
        capture_output=True, text=True,
    )

    if result.returncode == 0:
        print(f"  '{task_name}' registered — runs every day at 9:00 AM.")
    else:
        print(f"  ERROR: Could not register task.")
        print(f"  {result.stderr.strip()}")
        sys.exit(1)


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
    briefing_script     = os.path.join(script_dir, "morning_briefing.py")
    task_manager_script = os.path.join(script_dir, "task_manager.py")

    for path, name in [
        (briefing_script,     "morning_briefing.py"),
        (task_manager_script, "task_manager.py"),
    ]:
        if not os.path.exists(path):
            print(f"Error: {name} not found at {path}")
            sys.exit(1)

    pythonw_path = find_pythonw()
    desktop      = get_desktop_path()

    print("Morning Briefing Setup")
    print("=" * 44)
    print(f"Python:  {pythonw_path}")
    print(f"Scripts: {script_dir}")
    print(f"Desktop: {desktop}")
    print()

    register_scheduled_task(pythonw_path, briefing_script)
    create_desktop_shortcut(pythonw_path, task_manager_script, script_dir, desktop)

    print()
    print("All done!")
    print("  Morning Briefing: launches automatically every day at 9:00 AM")
    print("  Task Manager:     shortcut on your Desktop")
    print()
    print("To change the time: Task Scheduler > Task Scheduler Library > 'Morning Briefing'")
    print("To remove the task: schtasks /delete /tn \"Morning Briefing\" /f")


if __name__ == "__main__":
    main()

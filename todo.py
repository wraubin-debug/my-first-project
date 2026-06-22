import json
import os

TASKS_FILE = "tasks.json"

PRIORITIES = {"H": "High", "M": "Medium", "L": "Low"}
PRIORITY_ORDER = {"High": 0, "Medium": 1, "Low": 2}


def load_tasks():
    if not os.path.exists(TASKS_FILE):
        return []
    with open(TASKS_FILE, "r") as f:
        return json.load(f)


def save_tasks(tasks):
    with open(TASKS_FILE, "w") as f:
        json.dump(tasks, f, indent=2)


def prompt_priority():
    while True:
        raw = input("Priority ([H]igh / [M]edium / [L]ow): ").strip().upper()
        if raw in PRIORITIES:
            return PRIORITIES[raw]
        print("Enter H, M, or L.")


def add_task(tasks, title, priority):
    tasks.append({"title": title, "done": False, "priority": priority})
    save_tasks(tasks)
    print(f"Added: [{priority}] {title}")


def view_tasks(tasks):
    if not tasks:
        print("No tasks yet.")
        return
    sorted_tasks = sorted(tasks, key=lambda t: PRIORITY_ORDER.get(t.get("priority", "Low"), 2))
    print("\nYour tasks:")
    for i, task in enumerate(sorted_tasks, 1):
        status = "[x]" if task["done"] else "[ ]"
        priority = task.get("priority", "Low")
        print(f"  {i}. {status} [{priority}] {task['title']}")
    print()
    return sorted_tasks


def mark_done(tasks, sorted_tasks, index):
    if index < 1 or index > len(sorted_tasks):
        print("Invalid task number.")
        return
    task = sorted_tasks[index - 1]
    if task["done"]:
        print(f"Task {index} is already done.")
    else:
        task["done"] = True
        save_tasks(tasks)
        print(f"Marked done: {task['title']}")


def print_menu():
    print("\n--- To-Do List ---")
    print("1. View tasks")
    print("2. Add task")
    print("3. Mark task as done")
    print("4. Quit")


def main():
    tasks = load_tasks()
    while True:
        print_menu()
        choice = input("Choose an option: ").strip()

        if choice == "1":
            view_tasks(tasks)
        elif choice == "2":
            title = input("Task title: ").strip()
            if title:
                priority = prompt_priority()
                add_task(tasks, title, priority)
            else:
                print("Task title cannot be empty.")
        elif choice == "3":
            sorted_tasks = view_tasks(tasks)
            if tasks:
                try:
                    num = int(input("Enter task number to mark done: "))
                    mark_done(tasks, sorted_tasks, num)
                except ValueError:
                    print("Please enter a valid number.")
        elif choice == "4":
            print("Goodbye!")
            break
        else:
            print("Invalid choice. Please enter 1, 2, 3, or 4.")


if __name__ == "__main__":
    main()

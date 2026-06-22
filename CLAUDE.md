# To-Do List App

A command-line to-do list app written in Python. Tasks are saved to a JSON file so they persist between sessions.

## How to run

```
python todo.py
```

## Project files

- `todo.py` — the main app
- `tasks.json` — where tasks are saved (created automatically on first run)

## Code style

- **Simple and readable over clever.** Use straightforward logic that a beginner can follow.
- **Clear variable and function names.** Avoid abbreviations or one-letter names.
- **No unnecessary abstractions.** Don't add classes, decorators, or patterns unless they genuinely simplify things.
- **Short comments only when the reason isn't obvious.** Don't comment what the code already says clearly.

## What the app does

- Add tasks with a title and priority (High, Medium, Low)
- View all tasks sorted by priority (High first)
- Mark tasks as done
- Tasks are saved to `tasks.json` in JSON format

## Dependencies

None — uses only Python standard library (`json`, `os`).

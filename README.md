# green-electoral-sheets — local notes

This repository uses a virtual environment that lives next to the repo at `../electoral-venv`.

Quick notes

- Activate the venv on macOS / Linux:

```bash
source ../electoral-venv/bin/activate
```

- Deactivate when done:

```bash
deactivate
```

- VS Code: the workspace contains `.vscode/settings.json` which sets the workspace Python interpreter to `../electoral-venv/bin/python`. To change or re-select the interpreter:

1. Open the Command Palette (⇧⌘P) and run `Python: Select Interpreter`.
2. Choose the interpreter located at `../electoral-venv/bin/python` (or select another one).

If VS Code doesn’t pick up the interpreter automatically, the `ms-python.python` and `ms-python.vscode-pylance` extensions are recommended in `.vscode/extensions.json`.

Why we use a sibling venv

- Keeps the project separate from other environments but avoids committing the venv into the repo.

Notes about gitignore

If you want to ignore the parent venv locally, add it to `.git/info/exclude` or add a global gitignore entry — do not add parent-relative paths to this repo's `.gitignore` because that affects other clones.

# green-electoral-sheets — local notes

This repository uses a virtual environment that lives next to the repo at `../electoral-venv`.

Quick notes

## Setting up the venv:

You only need to do this once.

```bash
python -m venv ../electoral-venv
source ../electoral-venv/bin/activate
pip install -r requirements.txt
```

## Activate the venv on macOS / Linux:

```bash
source ../electoral-venv/bin/activate
```

## Deactivate when done:

```bash
deactivate
```

## Google setup

### App credentials

Can't remember how we got this file, but it's somewhere in the app settings in the google cloud console. 

### Client secrets file

This is copied from a help file and the steps aren't quite right, but you get the general idea. If I recall correctly I previously had to turn on OAuth for the project in one of the menus.

The client_secrets.json file is typically required for authenticating with Google APIs (such as Google Sheets, Drive, etc.). You can obtain it by following these steps:

1. Go to the Google Cloud Console: https://console.cloud.google.com/
2. Create a new project (or select an existing one).
3. Enable the relevant API (e.g., Google Sheets API, Google Drive API).
4. Go to "APIs & Services" > "Credentials".
5. Click "Create Credentials" > "OAuth client ID".
6. Configure the consent screen if prompted.
7. Choose "Desktop app" or "Other" as the application type.
8. Download the generated client_secrets.json file.
9. Place the file in your project directory as required by your code.


- VS Code: the workspace contains `.vscode/settings.json` which sets the workspace Python interpreter to `../electoral-venv/bin/python`. To change or re-select the interpreter:

1. Open the Command Palette (⇧⌘P) and run `Python: Select Interpreter`.
2. Choose the interpreter located at `../electoral-venv/bin/python` (or select another one).

If VS Code doesn’t pick up the interpreter automatically, the `ms-python.python` and `ms-python.vscode-pylance` extensions are recommended in `.vscode/extensions.json`.

Why we use a sibling venv

- Keeps the project separate from other environments but avoids committing the venv into the repo.

Notes about gitignore

If you want to ignore the parent venv locally, add it to `.git/info/exclude` or add a global gitignore entry — do not add parent-relative paths to this repo's `.gitignore` because that affects other clones.

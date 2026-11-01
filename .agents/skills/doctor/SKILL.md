---
name: Doctor
description: Health-check this repo on a new machine or fresh clone. Verifies the agent skills, Python version, dependencies, data directory, and git remote, then prints a pass/fail report with remediation commands. Run after cloning, on a new machine, or whenever the setup seems off.
---

# Repo Doctor

Run every check below, then print a single report table. Do not fix anything
without asking — propose the remediation command and let the user approve it.

## 1. Agent skills

These 15 skill directories must exist under `.agents/skills/`, each containing
a `SKILL.md`:

```
ask-matt  code-review  codebase-design  diagnosing-bugs  doctor
domain-modeling  grill-with-docs  grilling  handoff  implement
improve-codebase-architecture  setup-matt-pocock-skills  tdd
to-spec  to-tickets
```

Also verify `skills-lock.json` exists at the repo root.

- ✅ All present → pass
- ❌ Missing → remediation:
  ```bash
  # Restore everything from the lock file:
  DISABLE_TELEMETRY=1 npx -y skills experimental_install
  # Or install specific missing skills:
  DISABLE_TELEMETRY=1 npx -y skills add mattpocock/skills -a opencode -s <name> -y --copy
  ```

## 2. Python

```bash
python3 --version   # must be >= 3.12
```

- ❌ Missing/old → remediation: `pyenv install 3.12` or `brew install python@3.12`

## 3. Dependencies

```bash
python3 -c "import flask, openpyxl, gunicorn" 2>&1
```

- ❌ Import error → remediation:
  ```bash
  python3 -m venv .venv && source .venv/bin/activate
  pip install -r requirements.txt
  ```

## 4. Data directory

```bash
ls data/*.xlsx 2>/dev/null
```

- Files present → this machine has the **real business data** (owner's
  machine or the server). Normal.
- Missing on a fresh clone → **also normal**: `data/` is gitignored and the
  app recreates empty xlsx files on first run.
- ⚠️ Remind the user: `data/` contains real customer PII. Never commit it,
  never print its contents, never upload it.

## 5. Git

```bash
git remote get-url origin   # expect: https://github.com/SwarnadeepC0/inventory-app.git
git status --short
```

- ❌ Wrong/missing remote → remediation:
  `git remote add origin https://github.com/SwarnadeepC0/inventory-app.git`

## Report format

```
| Check              | Status | Notes            |
|--------------------|--------|------------------|
| Agent skills (15)  | ✅/❌  | …                |
| Python >= 3.12     | ✅/❌  | …                |
| Dependencies       | ✅/❌  | …                |
| Data directory     | ✅/⚠️  | …                |
| Git remote         | ✅/❌  | …                |
```

End with **"All checks passed ✅"**, or a numbered list of the remediation
commands the user should approve.

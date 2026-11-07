# AGENTS.md

Guidance for AI coding agents working in this repository.

## Project overview

Single-file **Flask** inventory & billing app for a saree/clothing retail shop.
No database — all data lives in Excel files via `openpyxl`.

- **Language/runtime:** Python 3.12, Flask 3.1
- **Storage:** `data/*.xlsx` (openpyxl) — inventory, orders, customers, order_items
- **Frontend:** server-rendered Jinja templates (`templates/`) + static assets (`static/`)
- **Server:** `gunicorn` in production, `flask run` / `python app.py` in dev
- **Hosting:** PythonAnywhere (`/home/swarnadeep/inventory-app`)

## Setup & run

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt   # flask, openpyxl, gunicorn
python app.py                     # dev server
```

The `data/` directory and empty xlsx files are auto-created on first run
(`init_inventory()`, `init_orders()`, `init_customers()` in `app.py`).

## Architecture

Everything is in **`app.py`** (~1k lines):

| Layer | Location |
|---|---|
| Excel helpers (read/save inventory, orders, customers) | top of `app.py` (`read_inventory`, `save_order`, `update_customer`, …) |
| Page routes | `/` (index), `/bill/<order_id>` (printable bill) |
| JSON API | `/api/dashboard`, `/api/inventory`, `/api/orders`, `/api/customers`, … |

Key conventions:

- **IDs:** items `ITM0001…`, orders sequential via `next_order_id()`.
- **Excel rows:** header in row 1, data from row 2; always `load_workbook` →
  mutate → `save` in one function call. No ORM, no migrations.
- **`amount_to_words()`** converts totals to words for printed bills.

## ⚠️ Critical rules

1. **`data/` contains REAL customer PII** (names, phones, emails, GST numbers).
   It is gitignored — **never commit it, never print its contents into logs or
   prompts, never upload it anywhere.**
2. Do not replace the Excel storage with a database without being asked.
3. Keep the app single-file unless a change genuinely requires a module split.

## Skills

Curated agent skills (from [mattpocock/skills](https://github.com/mattpocock/skills))
live in `.agents/skills/` — discovered automatically by OpenCode and other agents.

**New machine or fresh clone? Run the `doctor` skill first** — it verifies the
skills, Python, dependencies, data layout, and git remote, then reports
pass/fail with fixes.

Workflow for changes: **grill-with-docs** (align) → **to-spec** → **to-tickets** →
**implement** (drives **tdd**) → **code-review**. Use **diagnosing-bugs** for
defects, **improve-codebase-architecture** for design surveys, **handoff** to
transfer work between agents, and **ask-matt** if unsure which skill fits.

## Testing

No test suite exists yet. Verify changes manually: start the dev server and
exercise the index page, the `/api/*` endpoints, and bill rendering.

## Deployment

Push-to-deploy via GitHub Actions (`.github/workflows/deploy.yml`): every push
to `main` uploads `app.py`, `requirements.txt`, `templates/`, and `static/` to
`/home/swarnadeep/inventory-app` over the PythonAnywhere Files API, then
reloads `swarnadeep.pythonanywhere.com`. The `PA_API_TOKEN` is stored in
GitHub Actions secrets.

- `data/` is NEVER deployed — live customer data exists only on the server.
- If `requirements.txt` changes, packages must still be installed manually in
  a PythonAnywhere console (`pip install -r inventory-app/requirements.txt`),
  then reload.

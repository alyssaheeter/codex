# Debt/Credit Repair Agent

## Quick start (Python CLI)
```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python agent.py make-calendar
python agent.py filings
python agent.py letters
```

Outputs land in `out/`. Edit `data/plan.yaml` to change dates, addresses, and amounts.

## Google Apps Script workflow
A companion Google Apps Script implementation lives under `apps_script/` for users who prefer to generate Google Docs, Drive files, and an `.ics` calendar artifact directly inside Google Workspace.

1. In Google Drive, create a new **Apps Script** project (or open an existing one).
2. Replace the default `Code.gs` with the contents of [`apps_script/Code.gs`](apps_script/Code.gs). Update any plan values in the `PLAN` object as needed.
3. Replace the default `appsscript.json` in the project’s **Project Settings → Show "appsscript.json" manifest file** with [`apps_script/appsscript.json`](apps_script/appsscript.json) (time zone is already set to America/Chicago).
4. Save the project. Run the desired functions from the Apps Script editor:
   - `makeFilings()` – builds the opposition, motion to strike, and hardship declaration Google Docs (kept in a Drive folder named “Debt Agent Output”).
   - `makeLetters()` – creates all dispute and settlement letters, reusing ladder math for each approval-ready offer variant and bureau.
   - `makeCalendar()` – writes/updates `deadlines.ics` in the same Drive folder so you can import due dates into any calendar tool.
   - `scaffold()` – pre-creates placeholder Docs with the required filenames if you need to reserve them before drafting content.
5. Review the generated Docs in the “Debt Agent Output” folder (created automatically the first time you run any function) and export/share as needed.

The Apps Script renderer supports the same data points as the CLI agent, including automatic uppercase conversion, percent-to-dollar math, and deadline management, so outputs stay consistent across environments.

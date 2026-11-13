# Streamlit Community Cloud — deployment guide

Do **not** deploy until the project owner confirms.

This app supports an **Excel-backed** Community Cloud deployment when the approved `history.xlsx` is included in the GitHub repository. Simulated Demo Scenarios remain optional and are never forced solely because the host is Streamlit Cloud.

## What to upload (code)

Include:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`
- Optional docs: `DEMO_GUIDE.md`, `QA_CHECKLIST.md`, `DEPLOYMENT_GUIDE.md`

Still exclude unless separately approved:

- `demo_events.json`, `demo_scenario_events.json`
- `app_backup*.py`, `__pycache__/`, local IDE folders
- secrets / `.env`

## Including the approved workbook (owner approval required)

`.gitignore` currently excludes:

```text
history.xlsx
*.xlsx
```

**Do not commit the workbook until the owner authorises it.** When authorised, the exact steps are:

1. Confirm the file to publish is the approved `history.xlsx` (correct sheet `Trend History`, stations, and dates).
2. Edit `.gitignore` and remove the lines that exclude `history.xlsx` / `*.xlsx` (or replace `*.xlsx` with a narrower rule that still ignores other Excel files).
3. Stage and commit only that approved workbook:

```bash
git add -f history.xlsx
git add .gitignore
git commit -m "Add approved history.xlsx for Excel-backed Cloud demo"
git push
```

4. Redeploy / reboot the Streamlit Community Cloud app so it receives the workbook.

Until those steps are done, a Cloud app without the workbook will show a clear Excel-mode error unless the visitor enables Demo Scenarios.

## Recommended Streamlit Cloud settings

1. Deploy from the GitHub repo after the approved workbook policy is decided.
2. Main file: `app.py`
3. Python version: 3.11 if offered.
4. Optional secret / environment variable for session-only operator notes:

```toml
CP_DEMO_SESSION_EVENTS = "1"
```

On Community Cloud, session-only events are also auto-selected when the app path starts with `/mount/src`.

Do **not** set a flag that forces simulated data. Excel mode is the default whenever `history.xlsx` is present.

## Runtime behaviour on Cloud

| Topic | Behaviour |
| --- | --- |
| Excel present | Default source is the deployed `history.xlsx` copy |
| Demo Scenarios | Optional; clearly labelled simulated |
| Live SCADA | Not connected; workbook dates are historical |
| Operator events | Session-scoped per visitor; labelled session-only |
| Auto-refresh | Rechecks the deployed workbook file signature about every 5 seconds; does not pull new Excel from SCADA or from a local PC |

## Local Excel mode

```bash
streamlit run app.py
```

With `history.xlsx` next to `app.py`, Excel mode starts by default. Demo Scenarios stay optional.

## Security checklist before public upload

- [ ] Owner has authorised whether `history.xlsx` may be public on GitHub
- [ ] If authorised: `.gitignore` updated and only the approved workbook committed
- [ ] Local `demo_*.json` event files are **not** committed
- [ ] Backup copies of `app.py` are **not** committed
- [ ] README states Cloud notes are temporary demonstration records when session storage is used

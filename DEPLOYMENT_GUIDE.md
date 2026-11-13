# Streamlit Community Cloud — deployment guide

Do **not** deploy until the project owner confirms. This guide prepares a public **synthetic-only** demonstration.

## What to upload

Include:

- `app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `README.md`
- Optional docs: `DEMO_GUIDE.md`, `QA_CHECKLIST.md`, `DEPLOYMENT_GUIDE.md`

Exclude (already listed in `.gitignore`):

- `history.xlsx` (operational / historical workbook)
- `demo_events.json`, `demo_scenario_events.json`
- `app_backup*.py`, `__pycache__/`, local IDE folders

## Recommended Streamlit Cloud settings

1. Create a GitHub repository from the cleaned project (after owner approval).
2. In [share.streamlit.io](https://share.streamlit.io), deploy from that repo.
3. Main file: `app.py`
4. Python version: 3.11 if offered.
5. Under **Advanced settings → Secrets** or environment variables, set:

```toml
CP_DEMO_CLOUD = "1"
```

Or set environment variable `CP_DEMO_CLOUD=1`.

Omitting `history.xlsx` alone also enables cloud demo mode. Setting `CP_DEMO_CLOUD=1` makes the intent explicit and prevents accidental Excel mode if a workbook is added later.

## Runtime behaviour on Cloud

| Topic | Behaviour |
| --- | --- |
| Data | Synthetic Demo Scenarios only |
| Excel | Not loaded |
| Operator events | Session-scoped per visitor; not shared; not durable |
| Auto-refresh | Fragment still refreshes the simulated frame about every 5 seconds |
| Persistent disk | Not relied on for audits or workbook state |

## Local Excel mode after a cloud-oriented clone

Keep `history.xlsx` only on trusted local machines. With the workbook present and `CP_DEMO_CLOUD` unset/off, local Excel mode continues to work as before.

```bash
# Local Excel mode
unset CP_DEMO_CLOUD   # or set CP_DEMO_CLOUD=0
streamlit run app.py
```

## Security checklist before public upload

- [ ] `history.xlsx` is **not** in the GitHub tree
- [ ] Local `demo_*.json` event files are **not** committed
- [ ] Backup copies of `app.py` are **not** committed
- [ ] `CP_DEMO_CLOUD=1` is configured for the Cloud app
- [ ] README states that Cloud notes are temporary demonstration records

## After deploy (owner only)

Confirm the Cloud banner shows simulated/synthetic data, scenario switching works, and Maintenance events stay session-only with clear captions.

# CP Demo

Local and Streamlit Community Cloud demonstration of Spectrum’s cathodic-protection monitoring ideas.  
This is a **demonstration**, not a live SCADA system and not a production audit trail.

## What it does

- Four pages: **Overview**, **Historical Trends**, **Alerts**, **Maintenance**
- Five in-memory **Demo Scenarios** (normal / warning / critical / data quality / sudden drop)
- Illustrative threshold classification and operator inspection notes
- Read-only Excel mode for offline local demos when `history.xlsx` is present

The app never writes to `history.xlsx`.

## Local offline mode (Excel + optional scenarios)

**Needs:** Python 3.11+ (e.g. Anaconda env `cpmonitor`), this folder, and `history.xlsx` next to `app.py`.

```bash
conda activate cpmonitor
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501.

| Setting | Behaviour |
| --- | --- |
| Demo Scenarios **off** (default) | Reads `history.xlsx` about every 5 seconds |
| Demo Scenarios **on** | Synthetic readings in memory only |
| Operator events (Excel mode) | `demo_events.json` |
| Operator events (simulated mode) | `demo_scenario_events.json` |

Those JSON files are local demo trails only. They are separate from each other and from the workbook.

Presentation notes: `DEMO_GUIDE.md`  
QA steps: `QA_CHECKLIST.md`

## Cloud demonstration mode (synthetic only)

Public Streamlit Community Cloud deploys should **not** include `history.xlsx` or local event JSON files.

Cloud demo behaviour:

- Synthetic Demo Scenarios only (no Excel load)
- Operator notes stored in the **visitor’s browser session only**
- Clear on-screen labelling that records are temporary and not operational
- Visitors do not share each other’s demonstration notes

Force cloud behaviour even if a workbook is present:

```text
CP_DEMO_CLOUD=1
```

If `CP_DEMO_CLOUD` is unset and `history.xlsx` is missing, the app also enters cloud demo mode automatically.

See `DEPLOYMENT_GUIDE.md` before uploading to GitHub or Streamlit Community Cloud.

## Important limits

- Voltage/current tags do **not** prove pipeline protection or anode life.
- Thresholds are illustrative demo assumptions, not approved operating limits.
- Session or local JSON notes are **not** a production-grade audit system.

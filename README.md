# CP Demo

Cathodic-protection monitoring demonstration for Spectrum.  
This is a **demonstration dashboard**, not live SCADA and not a production audit system.

## What it does

- Four pages: **Overview**, **Historical Trends**, **Alerts**, **Maintenance**
- **Excel mode** (default when `history.xlsx` is present): historical SCADA tag readings from the workbook
- **Simulated Demo Scenarios** (optional): in-memory examples that are clearly labelled as not Excel observations
- Illustrative thresholds and operator inspection notes

The app never writes to `history.xlsx`.

## Data source behaviour

| Condition | Behaviour |
| --- | --- |
| `history.xlsx` present, Demo Scenarios **off** (default) | Excel mode |
| Demo Scenarios **on** | Simulated scenario data only (banner shown) |
| `history.xlsx` missing, Demo Scenarios **off** | Clear error — Excel mode does not start |
| `history.xlsx` missing, Demo Scenarios **on** | Simulated mode only |

Streamlit Community Cloud does **not** force simulated mode. If the approved workbook is deployed with the app, Excel mode is the default there too.

## Local launch

```bash
conda activate cpmonitor
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501.

| Setting | Behaviour |
| --- | --- |
| Demo Scenarios **off** | Reads local `history.xlsx` (saved-file changes rechecked about every 5 seconds) |
| Demo Scenarios **on** | Synthetic readings in memory only |
| Operator events (local Excel mode) | `demo_events.json` |
| Operator events (local simulated mode) | `demo_scenario_events.json` |

Presentation notes: `DEMO_GUIDE.md`  
QA steps: `QA_CHECKLIST.md`  
Cloud deploy notes: `DEPLOYMENT_GUIDE.md`

## Streamlit Community Cloud (Excel-backed)

When the approved `history.xlsx` is included in the GitHub repo used by Community Cloud:

- Default data source is the **deployed workbook copy** (historical, not live)
- Demo Scenarios remain available and are labelled simulated
- Operator notes use **session-only** storage on Cloud so visitors do not share records

Optional environment / secret:

```text
CP_DEMO_SESSION_EVENTS=1
```

Forces session-only event storage. On Community Cloud this is also detected automatically via the `/mount/src` app path.

`history.xlsx` is still listed in `.gitignore` until you explicitly approve committing it. See `DEPLOYMENT_GUIDE.md`.

## Important limits

- Voltage/current tags do **not** prove pipeline protection or anode life.
- Workbook dates are historical measurement times from Excel, not a live field clock.
- The app does **not** claim real-time monitoring or automatic SCADA→Excel synchronization.
- Thresholds are illustrative demo assumptions, not approved operating limits.
- Session or local JSON notes are **not** a production-grade audit system.

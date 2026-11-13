CP DEMO

See README.md for Excel vs Simulated data-source behaviour and Cloud notes.

Quick local launch (Anaconda env cpmonitor):

    conda activate cpmonitor
    cd "$env:USERPROFILE\Documents\Spectrum\cp_demo"
    python -m pip install -r requirements.txt
    python -m streamlit run app.py

http://localhost:8501

Default: Excel mode when history.xlsx is present (historical readings, not live).
Demo Scenarios: optional simulated data with a clear banner.
history.xlsx is read-only. Streamlit Cloud does not force simulated mode.
Workbook remains gitignored until explicitly approved — see DEPLOYMENT_GUIDE.md.

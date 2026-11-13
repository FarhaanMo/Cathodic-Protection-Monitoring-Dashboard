import json
import os
import uuid
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
from pathlib import Path

st.set_page_config(page_title='CP Demo', layout='wide', page_icon='🛡️')
st.markdown(
    '''
    <style>
    html, body, [data-testid="stAppViewContainer"] {
        color: #0b1f33;
        font-size: 15px;
    }
    .block-container {
        padding-top: 1.1rem;
        padding-bottom: 1.5rem;
        max-width: 1200px;
    }
    h1 {
        font-size: 1.85rem !important;
        color: #0b1f33 !important;
        margin-bottom: 0.35rem !important;
    }
    h2, h3 {
        color: #0b1f33 !important;
        letter-spacing: 0.01em;
    }
    [data-testid="stSidebar"] {
        background: #f8fafc;
        border-right: 1px solid #d9e2ec;
    }
    [data-testid="stSidebar"] * {
        color: #0b1f33 !important;
        font-size: 0.95rem !important;
    }
    [data-testid="stSidebar"] h2 {
        font-size: 1.2rem !important;
        font-weight: 700 !important;
    }
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p {
        font-size: 0.92rem !important;
        color: #243b53 !important;
    }
    [data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #d9e2ec;
        border-radius: 10px;
        padding: 14px 16px;
    }
    [data-testid="stMetricLabel"],
    [data-testid="stMetricLabel"] * {
        color: #1d4e89 !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
    }
    [data-testid="stMetricValue"],
    [data-testid="stMetricValue"] * {
        color: #0b1f33 !important;
        font-size: 1.7rem !important;
        font-weight: 650 !important;
    }
    [data-testid="stCaption"] {
        color: #334e68 !important;
        font-size: 0.86rem !important;
        line-height: 1.45 !important;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid #d9e2ec;
        border-radius: 8px;
    }
    .cp-source-box {
        background: #ffffff;
        border: 1px solid #d9e2ec;
        border-radius: 8px;
        padding: 0.75rem 1rem;
        margin: 0.4rem 0 0.9rem 0;
        color: #0b1f33;
        font-size: 0.92rem;
        line-height: 1.45;
    }
    .cp-source-box strong { color: #1d4e89; }
    .cp-sim-banner {
        background: #fff4e5;
        border: 2px solid #d9480f;
        border-radius: 8px;
        padding: 0.85rem 1rem;
        margin: 0.35rem 0 0.9rem 0;
        color: #0b1f33;
        font-size: 1rem;
        font-weight: 600;
        line-height: 1.4;
    }
    .stDeployButton { display: none; }
    </style>
    ''',
    unsafe_allow_html=True,
)

WORKBOOK = Path(__file__).resolve().parent / 'history.xlsx'
EVENTS_FILE = Path(__file__).resolve().parent / 'demo_events.json'
SCENARIO_EVENTS_FILE = Path(__file__).resolve().parent / 'demo_scenario_events.json'
STATIONS = [
    ('MUSCAT-OCC-OQGN-304', 'TR', 2, 3),
    ('MUSCAT-MANAL-OQGN-345', 'TR', 6, 7),
    ('FAHUD-OQGN-74', 'Solar', 10, 11),
    ('FAHUD-OQGN-114', 'Solar', 14, 15),
    ('FAHUD-OQGN-194', 'Solar', 18, 19),
]
PAGES = ['Overview', 'Historical Trends', 'Alerts', 'Maintenance']
SCENARIO_NAMES = [
    'Normal operation',
    'Sudden current drop',
    'Gradual output deterioration',
    'Missing measurement / possible communication issue',
    'Critical alarm with suggested inspection',
]
# Fewer than this many current readings is too thin for a median comparison.
MIN_VALID_CURRENTS = 3
# Each half of the history needs this many current readings before a trend label.
MIN_TREND_HALF = 5
# Recent median within 5% of the earlier median is called Stable.
TREND_BAND = 0.05
# Source samples are hourly. A larger gap weakens a sudden-drop claim.
EXPECTED_SAMPLE_HOURS = 1.0
MAX_SUDDEN_GAP_HOURS = 1.5
NORMAL = 'Normal'
WARNING = 'Warning'
CRITICAL = 'Critical'
DATA_QUALITY = 'Data quality'
UNKNOWN = 'Unknown'
REVIEW_STATUSES = (WARNING, CRITICAL, DATA_QUALITY, UNKNOWN)
OPEN_STATUSES = ('Open', 'Acknowledged')
NAVY = '#0b1f33'
BLUE = '#1d4e89'
STATUS_FILL = {
    NORMAL: '#d8f3dc',
    WARNING: '#fff3cd',
    CRITICAL: '#f8d7da',
    DATA_QUALITY: '#e9ecef',
    UNKNOWN: '#e9ecef',
    'Possible sudden drop': '#ffe8d6',
    'No sudden drop': '#d8f3dc',
    'Gap too large to confirm': '#fff3cd',
    'Insufficient data': '#e9ecef',
}


def events_file_for(demo_mode: bool) -> Path:
    """Excel mode and simulated mode keep separate local event files."""
    return SCENARIO_EVENTS_FILE if demo_mode else EVENTS_FILE


def is_cloud_demo() -> bool:
    """True for Streamlit Community Cloud / synthetic-only deployments."""
    flag = os.environ.get('CP_DEMO_CLOUD', '').strip().lower()
    if flag in {'1', 'true', 'yes', 'on'}:
        return True
    if flag in {'0', 'false', 'no', 'off'}:
        return False
    # Public deploys may omit history.xlsx on purpose.
    return not WORKBOOK.is_file()


def session_store_key(demo_mode: bool) -> str:
    """Per-visitor session keys so cloud users do not share event records."""
    return 'cp_session_events_scenario' if demo_mode else 'cp_session_events_excel'


def workbook_signature(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return (stat.st_mtime_ns, stat.st_size)


@st.cache_data(show_spinner='Reading history.xlsx...')
def load_history(path_str: str, mtime_ns: int, size: int):
    # mtime_ns and size are arguments so either change misses the cache and rereads the file.
    del mtime_ns, size
    raw = pd.read_excel(path_str, sheet_name='Trend History', header=None, skiprows=3)
    raw['timestamp'] = pd.to_datetime(raw[0], errors='coerce') + pd.to_timedelta(raw[1].astype(str), errors='coerce')
    frames = []
    for name, kind, vc, ic in STATIONS:
        frames.append(pd.DataFrame({
            'timestamp': raw['timestamp'],
            'station': name,
            'system': kind,
            'voltage': pd.to_numeric(raw[vc], errors='coerce'),
            'current': pd.to_numeric(raw[ic], errors='coerce'),
        }))
    frame = pd.concat(frames, ignore_index=True).dropna(subset=['timestamp']).sort_values('timestamp')
    if frame.empty:
        raise ValueError('history.xlsx loaded but contained no usable timestamps.')
    loaded_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return frame, loaded_at


def load_with_fallback():
    try:
        if not WORKBOOK.is_file():
            raise FileNotFoundError(f'{WORKBOOK.name} was not found next to app.py.')
        mtime_ns, size = workbook_signature(WORKBOOK)
        frame, loaded_at = load_history(str(WORKBOOK), mtime_ns, size)
    except Exception as exc:
        saved = st.session_state.get('cp_dataset')
        loaded_at = st.session_state.get('cp_loaded_at')
        if saved is None:
            return None, None, exc
        return saved.copy(), loaded_at, exc
    st.session_state['cp_dataset'] = frame.copy()
    st.session_state['cp_loaded_at'] = loaded_at
    return frame, loaded_at, None


def describe_read_error(exc: Exception) -> str:
    """Turn a file-read failure into a short message for the operator."""
    if isinstance(exc, FileNotFoundError):
        return 'history.xlsx was not found next to app.py.'
    if isinstance(exc, PermissionError):
        return 'history.xlsx is locked or still being saved.'
    name = type(exc).__name__
    text = str(exc).lower()
    if name in {'BadZipFile', 'BadZipfile'} or 'excel file format' in text or 'not a zip' in text:
        return 'history.xlsx could not be read. It may still be saving.'
    return f'history.xlsx could not be read ({name}).'


def format_timestamp(value) -> str:
    if value is None or pd.isna(value):
        return 'not available'
    return pd.Timestamp(value).strftime('%Y-%m-%d %H:%M')


def format_reading(value) -> str:
    if value is None or pd.isna(value):
        return 'Missing'
    return f'{float(value):.2f}'


def measurement_timestamps(frame: pd.DataFrame):
    newest = frame['timestamp'].max()
    valid = frame['voltage'].notna() | frame['current'].notna()
    latest_valid = frame.loc[valid, 'timestamp'].max() if valid.any() else pd.NaT
    return newest, latest_valid


def ordered_thresholds(warning_pct: int, critical_pct: int) -> tuple[int, int, bool]:
    """Keep Critical as the lower percentage so it is the stricter class."""
    warning_pct = int(warning_pct)
    critical_pct = int(critical_pct)
    if critical_pct < warning_pct:
        return warning_pct, critical_pct, False
    return warning_pct, max(5, warning_pct - 5), True


def classify_output(voltage, current, median, n_currents, warning_pct, critical_pct) -> str:
    """Classify the latest reading. Critical is tested before Warning."""
    if pd.isna(voltage) or pd.isna(current):
        return DATA_QUALITY
    if n_currents < MIN_VALID_CURRENTS or pd.isna(median) or median <= 0:
        return UNKNOWN
    # Critical is the lower share of the median, so it is checked first.
    if float(current) < float(median) * critical_pct / 100:
        return CRITICAL
    if float(current) < float(median) * warning_pct / 100:
        return WARNING
    return NORMAL


def explain_status(status, current, median, warning_pct, critical_pct) -> str:
    if status == DATA_QUALITY:
        return 'The latest voltage or current is missing, so this is not classified as Normal.'
    if status == UNKNOWN:
        return 'The baseline is missing, too small, or not positive, so this is not classified as Normal.'
    warning_a = float(median) * warning_pct / 100
    critical_a = float(median) * critical_pct / 100
    base = (
        f'Latest {float(current):.2f} A, baseline {float(median):.2f} A. '
        f'Warning is below {warning_a:.2f} A ({warning_pct}%). '
        f'Critical is below {critical_a:.2f} A ({critical_pct}%). '
    )
    if status == CRITICAL:
        return base + 'Critical applies because the latest current is below the Critical level. This is not a corrosion-risk category.'
    if status == WARNING:
        return base + 'Warning applies because the latest current is below the Warning level and not below Critical. This is not a corrosion-risk category.'
    return base + 'Normal means the latest current is not below the Warning level. It is not verified CP protection.'


def assess_station(station: pd.DataFrame, name: str, kind: str, warning_pct: int, critical_pct: int) -> dict:
    """Label the latest reading. Labels are demo output classes, not CP compliance."""
    measured = station[station['voltage'].notna() | station['current'].notna()]
    currents = station.loc[station['current'].notna(), 'current']
    median = float(currents.median()) if len(currents) else float('nan')
    missing_voltage = int(station['voltage'].isna().sum())
    missing_current = int(station['current'].isna().sum())
    missing_hours = int((station['voltage'].isna() | station['current'].isna()).sum())
    row = {
        'Station': name,
        'System': kind,
        'Voltage (V)': float('nan'),
        'Current (A)': float('nan'),
        'Reading time': pd.NaT,
        'Status': DATA_QUALITY,
        'Median current (A)': median,
        'Warning threshold (A)': float('nan'),
        'Critical threshold (A)': float('nan'),
        'Explanation': explain_status(DATA_QUALITY, None, None, warning_pct, critical_pct),
        'Missing voltage': missing_voltage,
        'Missing current': missing_current,
        'Missing hours': missing_hours,
        'Hours': int(len(station)),
    }
    if pd.notna(median) and median > 0:
        row['Warning threshold (A)'] = median * warning_pct / 100
        row['Critical threshold (A)'] = median * critical_pct / 100
    if measured.empty:
        return row
    latest = measured.iloc[-1]
    voltage = latest['voltage']
    current = latest['current']
    row['Voltage (V)'] = float(voltage) if pd.notna(voltage) else float('nan')
    row['Current (A)'] = float(current) if pd.notna(current) else float('nan')
    row['Reading time'] = latest['timestamp']
    status = classify_output(voltage, current, median, len(currents), warning_pct, critical_pct)
    row['Status'] = status
    row['Explanation'] = explain_status(status, current, median, warning_pct, critical_pct)
    return row


def assess_stations(df: pd.DataFrame, warning_pct: int, critical_pct: int) -> pd.DataFrame:
    rows = [
        assess_station(df[df['station'] == name], name, kind, warning_pct, critical_pct)
        for name, kind, _, _ in STATIONS
    ]
    return pd.DataFrame(rows)


def output_trend(station: pd.DataFrame) -> tuple[str, str]:
    """Compare the later half of valid current readings with the earlier half."""
    plain = (
        'Valid current readings are split by time into an earlier half and a later half. '
        'Each half needs at least 5 readings. '
        'Increasing means the recent median is more than 5% above the earlier median. '
        'Decreasing means it is more than 5% below. Otherwise the label is Stable. '
        'This is an output trend, not proof of electrode deterioration.'
    )
    if 'timestamp' not in station.columns or 'current' not in station.columns:
        return 'Insufficient Data', plain
    valid = station.dropna(subset=['current']).sort_values('timestamp')
    midpoint = len(valid) // 2
    earlier = valid.iloc[:midpoint]
    recent = valid.iloc[midpoint:]
    if len(earlier) < MIN_TREND_HALF or len(recent) < MIN_TREND_HALF:
        return 'Insufficient Data', plain + ' This station does not have enough readings in both periods.'
    earlier_median = float(earlier['current'].median())
    recent_median = float(recent['current'].median())
    numbers = (
        f' Earlier median {earlier_median:.2f} A from {len(earlier)} readings. '
        f'Recent median {recent_median:.2f} A from {len(recent)} readings.'
    )
    if earlier_median <= 0:
        return 'Insufficient Data', plain + ' The earlier median is not positive, so the change was not labelled.'
    if recent_median > earlier_median * (1 + TREND_BAND):
        label = 'Increasing'
    elif recent_median < earlier_median * (1 - TREND_BAND):
        label = 'Decreasing'
    else:
        label = 'Stable'
    return label, plain + numbers


def sudden_drop(station: pd.DataFrame, drop_pct: int) -> dict:
    """Compare the latest valid current with the previous valid current only."""
    result = {
        'Result': 'Insufficient data',
        'Previous current (A)': float('nan'),
        'Latest current (A)': float('nan'),
        'Change (%)': float('nan'),
        'Previous time': pd.NaT,
        'Latest time': pd.NaT,
        'Gap (hours)': float('nan'),
        'Note': 'At least two valid current readings are needed for a sudden-drop check.',
    }
    if 'timestamp' not in station.columns or 'current' not in station.columns:
        return result
    valid = station.dropna(subset=['current']).sort_values('timestamp')
    if len(valid) < 2:
        return result
    previous = valid.iloc[-2]
    latest = valid.iloc[-1]
    previous_a = float(previous['current'])
    latest_a = float(latest['current'])
    gap_hours = (pd.Timestamp(latest['timestamp']) - pd.Timestamp(previous['timestamp'])).total_seconds() / 3600
    result.update({
        'Previous current (A)': previous_a,
        'Latest current (A)': latest_a,
        'Previous time': previous['timestamp'],
        'Latest time': latest['timestamp'],
        'Gap (hours)': gap_hours,
    })
    if previous_a <= 0:
        result['Note'] = 'The previous current is not positive, so the percentage change was not labelled.'
        return result
    change_pct = (latest_a - previous_a) / previous_a * 100
    result['Change (%)'] = change_pct
    if gap_hours > MAX_SUDDEN_GAP_HOURS:
        result['Result'] = 'Gap too large to confirm'
        result['Note'] = (
            f'The readings are {gap_hours:.1f} hours apart. The expected sample interval is about '
            f'{EXPECTED_SAMPLE_HOURS:.0f} hour. This is not presented as a confirmed sudden failure.'
        )
        return result
    if change_pct <= -float(drop_pct):
        result['Result'] = 'Possible sudden drop'
        result['Note'] = (
            f'Latest current fell by {abs(change_pct):.1f}% from the previous valid reading '
            f'(threshold {drop_pct}%). This is a possible output anomaly, not confirmed equipment failure. '
            'It is separate from the median-based Warning and Critical classes.'
        )
    else:
        result['Result'] = 'No sudden drop'
        result['Note'] = (
            f'The change from the previous valid reading is {change_pct:.1f}%. '
            f'A possible sudden drop is below -{drop_pct}%.'
        )
    return result


def sudden_drop_table(df: pd.DataFrame, drop_pct: int) -> pd.DataFrame:
    rows = []
    for name, kind, _, _ in STATIONS:
        item = sudden_drop(df[df['station'] == name], drop_pct)
        rows.append({
            'Station': name,
            'Result': item['Result'],
            'System': kind,
            'Previous current (A)': format_reading(item['Previous current (A)']),
            'Latest current (A)': format_reading(item['Latest current (A)']),
            'Change (%)': 'n/a' if pd.isna(item['Change (%)']) else f"{item['Change (%)']:.1f}",
            'Previous time': format_timestamp(item['Previous time']),
            'Latest time': format_timestamp(item['Latest time']),
            'Gap (hours)': 'n/a' if pd.isna(item['Gap (hours)']) else f"{item['Gap (hours)']:.1f}",
            'Note': item['Note'],
        })
    return pd.DataFrame(rows)


def reading_freshness(summary: pd.DataFrame, newest_timestamp) -> pd.DataFrame:
    """Show whether the displayed reading matches the newest workbook hour."""
    rows = []
    for _, station in summary.iterrows():
        reading_time = station['Reading time']
        if pd.isna(reading_time):
            age = 'No measured reading'
        elif pd.isna(newest_timestamp):
            age = 'Unknown'
        elif pd.Timestamp(reading_time) < pd.Timestamp(newest_timestamp):
            age = 'Older than newest workbook hour'
        else:
            age = 'Matches newest workbook hour'
        rows.append({
            'Station': station['Station'],
            'Reading time': format_timestamp(reading_time),
            'Newest workbook hour': format_timestamp(newest_timestamp),
            'Reading age': age,
        })
    return pd.DataFrame(rows)


def _scenario_station(name: str, kind: str, timestamps, voltages, currents) -> pd.DataFrame:
    return pd.DataFrame({
        'timestamp': pd.to_datetime(timestamps),
        'station': name,
        'system': kind,
        'voltage': voltages,
        'current': currents,
    })


def build_scenario_frame(scenario: str) -> pd.DataFrame:
    """Build in-memory synthetic readings. Never written to history.xlsx."""
    hours = pd.date_range('2024-12-20 00:00', periods=24, freq='h')
    frames = []
    if scenario == 'Normal operation':
        profiles = {
            'MUSCAT-OCC-OQGN-304': (17.2, 36.2),
            'MUSCAT-MANAL-OQGN-345': (21.4, 16.9),
            'FAHUD-OQGN-74': (9.0, 18.0),
            'FAHUD-OQGN-114': (14.4, 5.3),
            'FAHUD-OQGN-194': (15.4, 4.6),
        }
        for name, kind, _, _ in STATIONS:
            voltage, current = profiles[name]
            frames.append(_scenario_station(
                name,
                kind,
                hours,
                [voltage] * len(hours),
                [current] * len(hours),
            ))
    elif scenario == 'Sudden current drop':
        for name, kind, _, _ in STATIONS:
            if name == 'MUSCAT-OCC-OQGN-304':
                currents = [36.2] * 23 + [18.0]
                voltages = [17.2] * 24
            else:
                currents = [16.9 if 'MANAL' in name else 5.0] * 24
                voltages = [21.0 if 'MANAL' in name else 14.0] * 24
            frames.append(_scenario_station(name, kind, hours, voltages, currents))
    elif scenario == 'Gradual output deterioration':
        for name, kind, _, _ in STATIONS:
            if name == 'MUSCAT-OCC-OQGN-304':
                currents = [36.0 - i * 0.7 for i in range(24)]
                voltages = [17.2] * 24
            else:
                currents = [16.0] * 24
                voltages = [21.0] * 24
            frames.append(_scenario_station(name, kind, hours, voltages, currents))
    elif scenario == 'Missing measurement / possible communication issue':
        for name, kind, _, _ in STATIONS:
            if name == 'MUSCAT-OCC-OQGN-304':
                # Latest hour keeps a voltage but loses current so status is Data quality,
                # not an automatic communication-failure label.
                currents = [36.2] * 23 + [float('nan')]
                voltages = [17.2] * 24
            else:
                currents = [16.9] * 24
                voltages = [21.0] * 24
            frames.append(_scenario_station(name, kind, hours, voltages, currents))
    else:
        # Critical alarm with suggested inspection
        for name, kind, _, _ in STATIONS:
            if name == 'MUSCAT-OCC-OQGN-304':
                currents = [36.2] * 20 + [10.0] * 4
                voltages = [17.2] * 24
            else:
                currents = [16.9] * 24
                voltages = [21.0] * 24
            frames.append(_scenario_station(name, kind, hours, voltages, currents))
    return pd.concat(frames, ignore_index=True).sort_values('timestamp')


def show_electrode_life_example():
    """Educational remaining-life example using fixed hypothetical assumptions only."""
    with st.expander('Illustrative electrode remaining-life example'):
        st.warning(
            'This example uses explicitly hypothetical assumptions. '
            'It does not use the Excel current readings, and it is not an estimate of actual electrode life.'
        )
        st.caption(
            'The supplied workbook does not contain pipe-to-soil potentials, electrode condition, '
            'consumption data, or the engineering parameters needed to estimate actual electrode remaining life.'
        )
        start_mass = 100.0
        rate = 2.0
        end_mass = 20.0
        remaining = (start_mass - end_mass) / rate
        st.markdown(
            f"""
**Assumed starting condition:** {start_mass:.0f} kg of usable electrode material  
**Assumed deterioration rate:** {rate:.1f} kg per year  
**Assumed end-of-life threshold:** {end_mass:.0f} kg remaining  

**Formula used for this illustration only:**  
remaining years = (starting mass − end-of-life mass) ÷ yearly consumption rate  

**Illustrative result:** ({start_mass:.0f} − {end_mass:.0f}) ÷ {rate:.1f} = **{remaining:.0f} years**
"""
        )
        timeline = pd.DataFrame({
            'Year from now': list(range(0, int(remaining) + 1)),
            'Illustrative remaining mass (kg)': [start_mass - rate * year for year in range(0, int(remaining) + 1)],
        })
        fig = px.line(
            timeline,
            x='Year from now',
            y='Illustrative remaining mass (kg)',
            title='Hypothetical remaining-mass timeline',
        )
        fig.update_traces(line_color=BLUE)
        fig.add_hline(y=end_mass, line_dash='dash', line_color='#9b2226', annotation_text='Assumed end-of-life')
        st.plotly_chart(style_chart(fig, height=300), config={'displayModeBar': False})
        st.caption(
            'Separate from the optional daily-average current line above. '
            'That line is an output trend only and is not electrode-life prediction.'
        )


def style_chart(fig, height=420):
    fig.update_layout(
        template='plotly_white',
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='#ffffff',
        font={'color': NAVY, 'size': 15},
        title={'font': {'size': 16, 'color': NAVY}},
        margin={'l': 16, 'r': 16, 't': 48, 'b': 16},
        height=height,
        legend={'font': {'size': 13}},
    )
    fig.update_xaxes(showgrid=False, color=NAVY, title_font={'size': 14}, tickfont={'size': 13})
    fig.update_yaxes(gridcolor='#e2e8f0', color=NAVY, title_font={'size': 14}, tickfont={'size': 13})
    return fig


def style_status_frame(frame: pd.DataFrame, column: str = 'Status'):
    """Presentation-only colouring. Does not change underlying values."""
    if column not in frame.columns or frame.empty:
        return frame

    def paint(value):
        fill = STATUS_FILL.get(str(value), '#ffffff')
        return f'background-color: {fill}; color: {NAVY}; font-weight: 600'

    return frame.style.map(paint, subset=[column])


def show_table(frame, status_column: str | None = None, height=None):
    """Readable dataframe defaults for presentation screens."""
    view = style_status_frame(frame, status_column) if status_column else frame
    config = {}
    if 'Status' in frame.columns:
        config['Status'] = st.column_config.TextColumn('Status', width='medium')
    if 'Result' in frame.columns:
        config['Result'] = st.column_config.TextColumn('Result', width='medium')
    if 'Explanation' in frame.columns:
        config['Explanation'] = st.column_config.TextColumn('Explanation', width='large')
    if 'Note' in frame.columns:
        config['Note'] = st.column_config.TextColumn('Note', width='large')
    if 'Reason' in frame.columns:
        config['Reason'] = st.column_config.TextColumn('Reason', width='large')
    if 'Follow-up' in frame.columns:
        config['Follow-up'] = st.column_config.TextColumn('Follow-up', width='medium')
    if 'Suggested inspection' in frame.columns:
        config['Suggested inspection'] = st.column_config.TextColumn('Suggested inspection', width='large')
    kwargs = {'hide_index': True, 'column_config': config or None, 'width': 'stretch'}
    if height is not None:
        kwargs['height'] = height
    st.dataframe(view, **kwargs)


def overview_table(summary: pd.DataFrame) -> pd.DataFrame:
    table = pd.DataFrame({
        'Station': summary['Station'],
        'Status': summary['Status'],
        'System': summary['System'],
        'Voltage (V)': summary['Voltage (V)'].map(format_reading),
        'Current (A)': summary['Current (A)'].map(format_reading),
        'Reading time': summary['Reading time'].map(format_timestamp),
    })
    return table


def page_overview(summary: pd.DataFrame, df: pd.DataFrame, latest_valid, newest, simulated: bool):
    valid_count = int((df['voltage'].notna() | df['current'].notna()).sum())
    review_count = int(summary['Status'].isin(REVIEW_STATUSES).sum())
    latest_date = format_timestamp(latest_valid).split(' ')[0]
    a, b, c, d = st.columns(4)
    a.metric('Stations', len(STATIONS))
    b.metric('Requiring review', review_count)
    c.metric('Latest valid date', latest_date)
    d.metric('Valid readings', valid_count)

    st.subheader('Station snapshot')
    show_table(overview_table(summary), status_column='Status')
    st.caption('Normal means the latest current is not below the Warning level. It is not verified pipeline protection.')

    fresh_col, chart_col = st.columns((1, 1.05))
    with fresh_col:
        st.subheader('Reading time and freshness')
        show_table(reading_freshness(summary, newest))
        with st.expander('Help: reading freshness'):
            st.write(
                'Each station shows the timestamp of the reading used for status. '
                'If that time is older than the newest workbook hour, the displayed reading is stale '
                'relative to the latest hour in the file. '
                + ('These times come from simulated data.' if simulated else 'These times come from history.xlsx.')
            )
    with chart_col:
        st.subheader('Latest current by station')
        chart = summary.dropna(subset=['Current (A)'])
        fig = px.bar(chart, y='Station', x='Current (A)', orientation='h', title='Latest current (A)')
        fig.update_traces(marker_color=BLUE, texttemplate='%{x:.2f}', textposition='outside')
        style_chart(fig, height=320)
        fig.update_yaxes(categoryorder='array', categoryarray=list(chart['Station'])[::-1])
        st.plotly_chart(fig, config={'displayModeBar': False})


def page_trends(df: pd.DataFrame, simulated: bool):
    name = st.selectbox('Station', [station[0] for station in STATIONS], key='trend_station')
    station = df[df['station'] == name]
    if simulated:
        st.info('Charts below use simulated demonstration data, not history.xlsx.')
    voltage, current = st.columns(2)
    with voltage:
        fig = px.line(station, x='timestamp', y='voltage', title='Voltage (V)')
        fig.update_traces(line_color=BLUE, connectgaps=False)
        st.plotly_chart(style_chart(fig, height=360), config={'displayModeBar': False})
    with current:
        fig = px.line(station, x='timestamp', y='current', title='Current (A)')
        fig.update_traces(line_color=BLUE, connectgaps=False)
        st.plotly_chart(style_chart(fig, height=360), config={'displayModeBar': False})
    st.caption(
        f'Missing voltage samples: {int(station.voltage.isna().sum())}'
        f' | Missing current samples: {int(station.current.isna().sum())}'
    )
    label, detail = output_trend(station)
    st.subheader('Output change over time')
    trend_col, note_col = st.columns((0.35, 0.65))
    with trend_col:
        st.metric('Output trend', label)
    with note_col:
        st.caption('Earlier-versus-later valid current comparison. Operational output trend only.')
        with st.expander('Help: how the output trend is calculated'):
            st.write(detail)
    show_projection = st.checkbox('Show illustrative current trend', key='show_projection')
    st.caption('Optional daily-average line. Not protection, corrosion, or anode-life prediction.')
    if show_projection:
        show_current_projection(station)
    show_electrode_life_example()


def show_current_projection(station: pd.DataFrame):
    daily = station.dropna(subset=['current']).set_index('timestamp').resample('D')['current'].mean().dropna()
    st.subheader('Illustrative current trend')
    if len(daily) < 3:
        st.info('Not enough days of current readings to draw a trend.')
        return
    elapsed_days = (daily.index - daily.index.min()).days.astype(float)
    slope, intercept = np.polyfit(elapsed_days, daily.to_numpy(dtype=float), 1)
    future_days = np.arange(0, int(elapsed_days.max()) + 8)
    projected = pd.DataFrame({
        'Date': daily.index.min() + pd.to_timedelta(future_days, unit='D'),
        'Illustrative trend (A)': intercept + slope * future_days,
    })
    fig = px.line(projected, x='Date', y='Illustrative trend (A)', title='Daily average current, extended 7 days')
    fig.update_traces(line_color=BLUE, name='Illustrative trend')
    fig.add_scatter(x=daily.index, y=daily.values, mode='markers', name='Daily average (A)', marker_color=NAVY)
    st.plotly_chart(style_chart(fig, height=360), config={'displayModeBar': False})
    st.metric('Illustrative slope', f'{slope:.3f} A/day')
    st.caption(
        'This line only extends the daily average current by 7 days. '
        'It is not a prediction of CP compliance, corrosion rate, or anode remaining useful life.'
    )


def output_rows(summary: pd.DataFrame) -> pd.DataFrame:
    """Show the latest current beside the demo baseline and both thresholds."""
    return pd.DataFrame({
        'Station': summary['Station'],
        'Status': summary['Status'],
        'Reading time': summary['Reading time'].map(format_timestamp),
        'Latest current (A)': summary['Current (A)'].map(format_reading),
        'Baseline median (A)': summary['Median current (A)'].map(format_reading),
        'Warning threshold (A)': summary['Warning threshold (A)'].map(format_reading),
        'Critical threshold (A)': summary['Critical threshold (A)'].map(format_reading),
        'Explanation': summary['Explanation'],
    })


def page_alerts(summary: pd.DataFrame, df: pd.DataFrame, warning_pct: int, critical_pct: int, drop_pct: int, simulated: bool):
    st.subheader('Current output classification')
    st.caption(
        f'Median-based classes: Warning below {warning_pct}%, Critical below {critical_pct}% (checked first). '
        + ('Source: simulated data.' if simulated else 'Source: history.xlsx.')
    )
    with st.expander('Help: median-based Warning and Critical'):
        st.write(
            'Uses each station\'s latest valid current only. '
            'These are demo output classes, not validated CP compliance or corrosion risk.'
        )
    alarms = summary[summary['Status'].isin([CRITICAL, WARNING])]
    if alarms.empty:
        st.info('No Warning or Critical output. No station\'s latest current is below the Warning level.')
    else:
        show_table(output_rows(alarms), status_column='Status')
        st.caption(
            'Cause unknown. Unverified possibilities: equipment output, measurement error, or communications.'
        )
    st.markdown('**Latest reading check**')
    show_table(output_rows(summary), status_column='Status')

    st.subheader('Sudden output drop check')
    st.caption(
        f'Previous-reading check: possible sudden drop is a fall of {drop_pct}% or more. '
        'Separate from the median-based classes above.'
    )
    with st.expander('Help: sudden-drop method'):
        st.write(
            'Compares the latest valid current with the immediately previous valid current for the same station. '
            'A large time gap is not treated as a confirmed sudden failure. '
            'This is a possible output anomaly, not confirmed equipment failure.'
        )
    drops = sudden_drop_table(df, drop_pct)
    flagged = drops[drops['Result'] == 'Possible sudden drop']
    if flagged.empty:
        st.info('No possible sudden drop at the current threshold.')
    else:
        show_table(flagged, status_column='Result')
        st.caption('Possible output anomaly, not confirmed equipment failure.')
    st.markdown('**Previous-to-latest comparison**')
    show_table(drops, status_column='Result')

    st.subheader('Historical data-quality observation')
    st.caption(
        'Missing hours are Data quality / Unknown concerns, not automatic communication failure or low output.'
    )
    gaps = summary[['Station', 'Missing voltage', 'Missing current', 'Missing hours']].rename(columns={
        'Missing voltage': 'Missing voltage samples',
        'Missing current': 'Missing current samples',
        'Missing hours': 'Hours with a gap',
    })
    show_table(gaps)


def empty_store() -> dict:
    return {'events': [], 'audit': []}


def load_store(path=None, session_key=None) -> dict:
    """Load demo events from a local file or from this visitor's Streamlit session."""
    if session_key is not None:
        if session_key not in st.session_state or not isinstance(st.session_state[session_key], dict):
            st.session_state[session_key] = empty_store()
        store = st.session_state[session_key]
        events = store.get('events') if isinstance(store.get('events'), list) else []
        audit = store.get('audit') if isinstance(store.get('audit'), list) else []
        store['events'] = events
        store['audit'] = audit
        return store
    path = Path(path) if path else EVENTS_FILE
    if not path.is_file():
        return empty_store()
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return empty_store()
    if not isinstance(data, dict):
        return empty_store()
    events = data.get('events') if isinstance(data.get('events'), list) else []
    audit = data.get('audit') if isinstance(data.get('audit'), list) else []
    return {'events': events, 'audit': audit}


def save_store(store: dict, path=None, session_key=None) -> None:
    """Persist demo events to a local JSON file or this visitor's session only."""
    if session_key is not None:
        st.session_state[session_key] = store
        return
    path = Path(path) if path else EVENTS_FILE
    path.write_text(json.dumps(store, indent=2), encoding='utf-8')


def stamp() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def add_audit(store: dict, event: dict, action: str, operator: str, detail: str) -> None:
    store['audit'].append({
        'time': stamp(),
        'station': event['station'],
        'event_id': event['id'],
        'action': action,
        'operator': operator,
        'detail': detail,
    })


def create_event(station: str, level: str, operator: str, note: str, path=None, session_key=None) -> tuple[bool, str]:
    """Create one event when the operator submits the form. Refresh does not call this."""
    operator = operator.strip()
    note = note.strip()
    if not operator:
        return False, 'Enter an operator name.'
    store = load_store(path, session_key=session_key)
    for event in store['events']:
        if event['station'] == station and event['level'] == level and event['status'] in OPEN_STATUSES:
            return False, 'An open event already exists for this station and level. It was not created again.'
    now = stamp()
    event = {
        'id': uuid.uuid4().hex[:8],
        'station': station,
        'level': level,
        'status': 'Open',
        'operator': operator,
        'note': note,
        'created_at': now,
        'updated_at': now,
    }
    store['events'].append(event)
    add_audit(store, event, 'Created', operator, note or 'Event opened.')
    save_store(store, path, session_key=session_key)
    return True, 'Inspection event created.'


def update_event(event_id: str, new_status: str, operator: str, note: str, path=None, session_key=None) -> tuple[bool, str]:
    operator = operator.strip()
    note = note.strip()
    if not operator:
        return False, 'Enter an operator name.'
    if new_status == 'Resolved' and not note:
        return False, 'Enter a corrective-action note before resolving.'
    store = load_store(path, session_key=session_key)
    for event in store['events']:
        if event['id'] != event_id:
            continue
        event['status'] = new_status
        event['operator'] = operator
        event['updated_at'] = stamp()
        if note:
            event['note'] = note
        action = 'Acknowledged' if new_status == 'Acknowledged' else 'Resolved'
        add_audit(store, event, action, operator, note or 'Status updated.')
        save_store(store, path, session_key=session_key)
        return True, f'Event marked {new_status}.'
    return False, 'That event was not found.'


def event_choices(summary: pd.DataFrame) -> list[tuple[str, str]]:
    """Stations an operator may open an event for. Normal output with no gaps is excluded."""
    choices = []
    for _, station in summary.iterrows():
        if station['Status'] != NORMAL:
            choices.append((station['Station'], station['Status']))
        elif int(station['Missing hours']) > 0:
            choices.append((station['Station'], DATA_QUALITY))
    return choices


def page_events(summary: pd.DataFrame, demo_mode: bool):
    """Operator-created events. Recommendations above this section are generated and not saved."""
    cloud = is_cloud_demo()
    store_path = None if cloud else events_file_for(demo_mode)
    session_key = session_store_key(demo_mode) if cloud else None
    store_label = 'this browser session only' if cloud else store_path.name
    st.subheader('Operator inspection events')
    if st.session_state.get('event_message'):
        st.success(st.session_state.pop('event_message'))
    if cloud:
        st.caption(
            f'Demonstration notes for {store_label}. Not permanent and not operational records. '
            'Other visitors cannot see these notes. history.xlsx is not used in cloud demo mode.'
        )
    elif demo_mode:
        st.caption(
            f'Simulated-mode notes only → {store_label}. Not operational records. '
            'Separate from Excel-mode demo_events.json. history.xlsx is unchanged.'
        )
    else:
        st.caption(
            f'Created only when you press Create event → {store_label}. '
            'Separate from the generated suggestions above. history.xlsx is unchanged.'
        )
    choices = event_choices(summary)
    form_col, list_col = st.columns((1, 1.15))
    with form_col:
        st.markdown('**Create event**')
        if not choices:
            st.info('No flagged station is available for a new event.')
        else:
            labels = [f'{station} — {level}' for station, level in choices]
            with st.form('create_event_form', clear_on_submit=True):
                picked = st.selectbox('Flagged station', labels)
                operator = st.text_input('Operator name')
                note = st.text_input('Note')
                submitted = st.form_submit_button('Create event')
            if submitted:
                station, level = choices[labels.index(picked)]
                ok, message = create_event(
                    station, level, operator, note, path=store_path, session_key=session_key
                )
                if ok:
                    st.session_state['event_message'] = message
                    st.rerun()
                else:
                    st.warning(message)

    store = load_store(store_path, session_key=session_key)
    events = store['events']
    with list_col:
        st.markdown('**Event status**')
        if not events:
            st.info('No operator events yet.')
        else:
            shown = pd.DataFrame([{
                'Event': event['id'],
                'Station': event['station'],
                'Level': event['level'],
                'Status': event['status'],
                'Operator': event['operator'],
                'Note': event['note'],
                'Created': event['created_at'],
                'Updated': event['updated_at'],
                'Store': store_label,
            } for event in events])
            show_table(shown)

    if events:
        open_events = [event for event in events if event['status'] != 'Resolved']
        if open_events:
            st.markdown('**Update open events**')
        for event in open_events:
            st.markdown(f"**{event['station']}** · {event['id']} · {event['status']}")
            with st.form(f"update_{event['id']}"):
                operator = st.text_input('Operator name')
                note = st.text_input('Corrective-action note')
                acknowledge, resolve = st.columns(2)
                ack = acknowledge.form_submit_button('Acknowledge') if event['status'] == 'Open' else False
                done = resolve.form_submit_button('Resolve')
            if ack:
                ok, message = update_event(
                    event['id'], 'Acknowledged', operator, note, path=store_path, session_key=session_key
                )
                if ok:
                    st.session_state['event_message'] = message
                    st.rerun()
                st.warning(message)
            if done:
                ok, message = update_event(
                    event['id'], 'Resolved', operator, note, path=store_path, session_key=session_key
                )
                if ok:
                    st.session_state['event_message'] = message
                    st.rerun()
                st.warning(message)

    with st.expander('Audit history'):
        if cloud:
            st.caption(
                'Temporary demonstration audit trail for this browser session only. '
                'It is not a production-grade immutable audit system and is cleared when the session ends.'
            )
        else:
            st.caption(
                f'Local demo audit trail in {store_label}. '
                'Simulated-mode records stay separate from Excel-mode records. '
                'Not a production-grade immutable audit system.'
            )
        audit = load_store(store_path, session_key=session_key)['audit']
        if not audit:
            st.info('No audit records yet.')
        else:
            audit_table = pd.DataFrame(audit).rename(columns={
                'time': 'Time',
                'station': 'Station',
                'event_id': 'Event',
                'action': 'Action',
                'operator': 'Operator',
                'detail': 'Detail',
            })
            show_table(audit_table.iloc[::-1])


def page_maintenance(summary: pd.DataFrame, demo_mode: bool):
    inspect_col, quality_col = st.columns(2)
    with inspect_col:
        st.subheader('Recommended physical inspection')
        st.caption('Generated suggestion only. Not an operator event. No work order is created.')
        inspections = []
        for _, station in summary.iterrows():
            if station['Status'] in (CRITICAL, WARNING):
                inspections.append({
                    'Station': station['Station'],
                    'Suggested inspection': 'Inspect equipment output on site and confirm pipe-to-soil readings.',
                    'Reason': f"Latest current is {station['Status']}. The cause is not known.",
                })
        if inspections:
            show_table(pd.DataFrame(inspections))
            st.caption('Unverified possibilities: equipment output, measurement error, or communications.')
        else:
            st.info('No physical inspection is suggested.')

    with quality_col:
        st.subheader('Data Quality Follow-up')
        st.caption('Data quality / Unknown only. Not automatic communication failure. No work order.')
        follow_up = []
        for _, station in summary.iterrows():
            if station['Status'] == DATA_QUALITY:
                follow_up.append({
                    'Station': station['Station'],
                    'Follow-up': 'Check why the latest voltage or current is incomplete.',
                    'Reason': (
                        'The latest voltage or current reading is missing. '
                        'Communication problems are one unverified possibility, not a diagnosed cause.'
                    ),
                })
            if int(station['Missing hours']) > 0:
                follow_up.append({
                    'Station': station['Station'],
                    'Follow-up': 'Confirm why some stored hours have no voltage or current.',
                    'Reason': f"{int(station['Missing hours'])} of {int(station['Hours'])} hours are missing.",
                })
            if station['Status'] == UNKNOWN:
                follow_up.append({
                    'Station': station['Station'],
                    'Follow-up': 'Collect more valid current readings before judging output.',
                    'Reason': 'The current baseline is missing, too small, or not positive.',
                })
        if follow_up:
            show_table(pd.DataFrame(follow_up))
        else:
            st.info('No data-quality follow-up is suggested.')
    page_events(summary, demo_mode)


@st.fragment(run_every='5s')
def render_dashboard():
    cloud = is_cloud_demo()
    with st.sidebar:
        st.header('CP Demo')
        if cloud:
            st.caption('Cloud demonstration mode · simulated data only')
        refresh_now = st.button('Refresh Now', key='refresh_now', use_container_width=True)
        page = st.radio('Page', PAGES, key='page')
        st.markdown('**Thresholds**')
        warning_pct = st.slider(
            'Warning below (% of station median)',
            50,
            95,
            80,
            5,
            key='warning_pct',
        )
        critical_pct = st.slider(
            'Critical below (% of station median)',
            20,
            90,
            60,
            5,
            key='critical_pct',
        )
        drop_pct = st.slider(
            'Sudden drop from previous reading (%)',
            10,
            80,
            30,
            5,
            key='drop_pct',
        )
        warning_pct, critical_pct, adjusted = ordered_thresholds(warning_pct, critical_pct)
        st.caption('Illustrative demo assumptions only.')
        if adjusted:
            st.caption(f'Critical must stay below Warning. Using Critical {critical_pct}%.')
        with st.expander('Help: threshold meaning'):
            st.write(
                'Warning and Critical compare the latest current with each station\'s historical median. '
                'Sudden drop compares the latest reading with the previous valid reading. '
                'These are not approved operating limits or corrosion-risk categories.'
            )
        st.markdown('**Demo Scenarios**')
        if cloud:
            demo_mode = True
            scenario = st.radio('Selected scenario', SCENARIO_NAMES, key='demo_scenario')
            st.caption('Public cloud demo uses synthetic readings only. No Excel workbook is loaded.')
        else:
            demo_mode = st.checkbox('Enable Demo Scenarios', value=False, key='demo_mode')
            scenario = None
            if demo_mode:
                scenario = st.radio('Selected scenario', SCENARIO_NAMES, key='demo_scenario')
                st.caption('Simulated in memory only. Never written to history.xlsx.')
            else:
                st.caption('Excel mode: saved edits refresh about every 5 seconds.')

    if refresh_now and not cloud:
        load_history.clear()

    st.title('CP Demo')
    if demo_mode:
        event_note = (
            'Events stay in this browser session only and are not permanent.'
            if cloud else
            'Events created here go to demo_scenario_events.json only.'
        )
        st.markdown(
            f'<div class="cp-sim-banner">SIMULATED DATA — not from history.xlsx<br>'
            f'Scenario: {scenario}<br>'
            f'{event_note}</div>',
            unsafe_allow_html=True,
        )
        df = build_scenario_frame(scenario)
        loaded_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        read_error = None
    else:
        df, loaded_at, read_error = load_with_fallback()

    st.warning(
        'Voltage and current come from rectifier and solar tags. '
        'They do not show whether the pipeline is protected, and they do not estimate anode life.'
    )

    if not demo_mode and read_error is not None and df is None:
        st.warning(f'{describe_read_error(read_error)} No earlier data is available.')
        st.stop()
    if not demo_mode and read_error is not None:
        st.warning(
            f'{describe_read_error(read_error)} '
            f'Showing the last successfully loaded data from {loaded_at}. '
            'The file will be tried again on the next refresh.'
        )

    newest, latest_valid = measurement_timestamps(df)
    if demo_mode:
        source_line = (
            'synthetic demonstration readings for Streamlit Community Cloud'
            if cloud else
            'simulated demonstration readings (not live)'
        )
        st.markdown(
            f'<div class="cp-source-box">'
            f'<strong>Data source:</strong> {source_line}<br>'
            f'<strong>Scenario build time (app clock):</strong> {loaded_at}<br>'
            f'<strong>Newest simulated measurement time:</strong> {format_timestamp(newest)}<br>'
            f'<strong>Latest simulated valid measurement:</strong> {format_timestamp(latest_valid)}'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f'<div class="cp-source-box">'
            f'<strong>Data source:</strong> history.xlsx historical SCADA tags (not live telemetry)<br>'
            f'<strong>App refresh / Excel load time:</strong> {loaded_at}<br>'
            f'<strong>Newest workbook hour (historical):</strong> {format_timestamp(newest)}<br>'
            f'<strong>Latest valid station measurement (historical):</strong> {format_timestamp(latest_valid)}'
            f'</div>',
            unsafe_allow_html=True,
        )
    summary = assess_stations(df, warning_pct, critical_pct)
    if page == 'Overview':
        page_overview(summary, df, latest_valid, newest, demo_mode)
    elif page == 'Historical Trends':
        page_trends(df, demo_mode)
    elif page == 'Alerts':
        page_alerts(summary, df, warning_pct, critical_pct, drop_pct, demo_mode)
    else:
        page_maintenance(summary, demo_mode)
    st.divider()
    if cloud:
        st.caption('CP Demo · public cloud demonstration · synthetic data · session-only notes')
    elif demo_mode:
        st.caption('CP Demo · simulated data in memory only · workbook not modified')
    else:
        st.caption('CP Demo · local offline demo · workbook is read only')


if __name__ == '__main__':
    render_dashboard()

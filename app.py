import html
from calendar import monthrange
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from zoneinfo import ZoneInfo


st.set_page_config(
    page_title="Dashboard CTOs - Cusezar",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

MONTHS_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

WEEKDAYS_ES = ["Dom", "Lun", "Mar", "Mie", "Jue", "Vie", "Sab"]
EXCEL_PATH = Path("CTOs.xlsx")
TABLE_NAME = "CTO"
BOGOTA_TZ = ZoneInfo("America/Bogota")


def inject_base_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
        #MainMenu, footer, header { visibility: hidden; }
        .block-container {
            padding-top: 2.6rem !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
            max-width: 100% !important;
        }
        .page-head {
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            gap: 16px;
            margin: 0 0 18px 0;
            padding-bottom: 10px;
            border-bottom: 1px solid #E5E9F0;
        }
        .page-title {
            font-size: 24px;
            font-weight: 800;
            color: #111827;
            line-height: 1.1;
            margin: 0;
        }
        .page-sub {
            font-size: 12px;
            color: #9CA3AF;
            margin: 4px 0 0 0;
        }
        .section-head {
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            gap: 16px;
            margin: 0 0 10px 0;
            padding-bottom: 8px;
            border-bottom: 1px solid #E5E9F0;
        }
        .section-title {
            font-size: 15px;
            font-weight: 700;
            color: #111827;
            line-height: 1.2;
            margin: 0;
        }
        .section-sub {
            font-size: 11px;
            color: #9CA3AF;
            line-height: 1.4;
            text-align: right;
            max-width: 60%;
            margin: 0;
        }
        .kpi-card {
            background: #FFFFFF;
            border: 1px solid #E5E9F0;
            border-radius: 14px;
            padding: 18px 20px;
            position: relative;
            overflow: hidden;
            box-shadow: 0 1px 4px rgba(15, 23, 42, .05);
            transition: transform .15s, box-shadow .15s;
            min-height: 116px;
        }
        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(15, 23, 42, .08);
        }
        .kpi-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            border-radius: 14px 14px 0 0;
        }
        .kp-blue::before { background: #7BA7D4; }
        .kp-green::before { background: #6BBF9E; }
        .kp-red::before { background: #D98B8B; }
        .kpi-label {
            font-size: 10px;
            font-weight: 700;
            color: #9CA3AF;
            text-transform: uppercase;
            letter-spacing: .06em;
            margin-bottom: 6px;
        }
        .kpi-value {
            font-size: 28px;
            font-weight: 800;
            line-height: 1;
            font-family: 'DM Mono', monospace;
            margin-bottom: 6px;
        }
        .kpi-sub {
            font-size: 11px;
            color: #9CA3AF;
        }
        .kp-blue .kpi-value { color: #4A7BA8; }
        .kp-green .kpi-value { color: #3D8B6E; }
        .kp-red .kpi-value { color: #B05B5B; }
        .info-note {
            padding: 10px 14px;
            background: #EEF3FA;
            border: 1px solid #C8DCF0;
            border-radius: 10px;
            font-size: 12px;
            color: #4A7BA8;
            font-weight: 500;
            margin-bottom: 18px;
        }
        div[data-testid="stSelectbox"] > label {
            font-size: 11px !important;
            font-weight: 600 !important;
            color: #6B7280 !important;
            text-transform: uppercase !important;
            letter-spacing: .05em !important;
            margin-bottom: 4px !important;
        }
        div[data-testid="stSelectbox"] > div > div {
            border-radius: 10px !important;
            border: 1.5px solid #E5E9F0 !important;
            background: #FAFBFC !important;
            font-size: 13px !important;
            color: #374151 !important;
        }
        div[data-testid="stSelectbox"] > div > div:focus-within {
            border-color: #7BA7D4 !important;
            box-shadow: 0 0 0 3px rgba(123, 167, 212, .12) !important;
        }
        .legend-wrap {
            display: flex;
            gap: 16px;
            justify-content: flex-end;
            align-items: center;
            margin-top: 26px;
            flex-wrap: wrap;
        }
        .legend-item {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            color: #6B7280;
            font-size: 12px;
            font-weight: 600;
        }
        .legend-dot {
            width: 10px;
            height: 10px;
            border-radius: 3px;
        }
        .legend-signed { background: #6BBF9E; }
        .legend-unsigned { background: #D98B8B; }
        .month-chip {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 100%;
            min-height: 38px;
            background: #F8FAFC;
            border: 1px solid #E5E9F0;
            border-radius: 10px;
            color: #111827;
            font-size: 14px;
            font-weight: 700;
        }
        div[data-testid="stButton"] > button {
            width: 100%;
            border-radius: 10px;
            border: 1.5px solid #E5E9F0;
            background: #FFFFFF;
            color: #4A7BA8;
            font-weight: 700;
        }
        div[data-testid="stButton"] > button:hover {
            border-color: #7BA7D4;
            color: #355D8A;
        }
        .table-shell {
            overflow-x: auto;
            border-radius: 10px;
            border: 1px solid #E5E9F0;
            max-height: 460px;
            overflow-y: auto;
        }
        .rt {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        .rt th {
            background: #B5545C;
            padding: 10px 14px;
            text-align: left;
            font-size: 11px;
            font-weight: 700;
            color: #FFF7F7;
            text-transform: uppercase;
            letter-spacing: .05em;
            border-bottom: 1px solid #9F434B;
            white-space: nowrap;
        }
        .rt td {
            padding: 10px 14px;
            border-bottom: 1px solid #F1D9DB;
            color: #6B7280;
        }
        .rt td:first-child {
            background: #F8E8E8;
            color: #8F3942;
            font-weight: 700;
        }
        .rt tr:last-child td { border-bottom: none; }
        .rt tr:hover td { background: #FCF4F5; }
        .rt tr:hover td:first-child { background: #F8E8E8; }
        .pill-count {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 4px 10px;
            border-radius: 999px;
            background: #F8E8E8;
            color: #8F3942;
            font-size: 12px;
            font-weight: 700;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_page_heading(title: str, subtitle: str = "") -> None:
    subtitle_html = f'<div class="page-sub">{html.escape(subtitle)}</div>' if subtitle else ""
    st.markdown(
        f'<div class="page-head"><div><div class="page-title">{html.escape(title)}</div>{subtitle_html}</div></div>',
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str = "") -> str:
    subtitle_html = f'<div class="section-sub">{html.escape(subtitle)}</div>' if subtitle else ""
    return (
        f'<div class="section-head">'
        f'<div class="section-title">{html.escape(title)}</div>'
        f"{subtitle_html}"
        f"</div>"
    )


def kpi_card(label: str, value: str, subtitle: str, tone: str) -> str:
    return (
        f'<div class="kpi-card {tone}">'
        f'<div class="kpi-label">{html.escape(label)}</div>'
        f'<div class="kpi-value">{html.escape(value)}</div>'
        f'<div class="kpi-sub">{html.escape(subtitle)}</div>'
        f"</div>"
    )


def read_excel_table(path: Path, table_name: str) -> pd.DataFrame:
    wb = load_workbook(path, data_only=True, read_only=False)
    try:
        for ws in wb.worksheets:
            if table_name in ws.tables:
                ref = ws.tables[table_name].ref
                rows = list(ws[ref])
                data = [[cell.value for cell in row] for row in rows]
                headers = data[0]
                values = data[1:]
                return pd.DataFrame(values, columns=headers)
    finally:
        wb.close()
    raise ValueError(f"No se encontro la tabla '{table_name}' en {path.name}.")


def get_excel_signature(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_mtime_ns, stat.st_size


@st.cache_data(show_spinner=False)
def load_cto_data(_signature: tuple[int, int]) -> pd.DataFrame:
    df = read_excel_table(EXCEL_PATH, TABLE_NAME).copy()
    df.columns = [str(col).strip() for col in df.columns]
    rename_map = {
        "Torre/ZC": "Torre",
        "Fecha anterior": "Fecha anterior",
        "Fecha actual": "Fecha actual",
        "Firmado": "Firmado",
    }
    df = df.rename(columns=rename_map)

    for col in ["Proyecto", "Torre"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    for col in ["Apartamento", "Local"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in ["Fecha anterior", "Fecha actual"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df["Firmado"] = pd.to_numeric(df["Firmado"], errors="coerce").fillna(0).astype(int)
    df["Unidad"] = df.apply(build_unit_label, axis=1)
    df["Torre_display"] = df["Torre"].replace("", "-")

    df = df.sort_values(["Fecha actual", "Proyecto", "Torre_display", "Unidad"], na_position="last").reset_index(drop=True)
    return df


def build_unit_label(row: pd.Series) -> str:
    apartamento = row.get("Apartamento")
    local = row.get("Local")

    if pd.notna(apartamento):
        return f"Apto {format_number(apartamento)}"
    if pd.notna(local):
        return f"Local {format_number(local)}"
    return ""


def format_number(value) -> str:
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value)
    return str(int(num)) if num.is_integer() else str(num)


def format_short_date(value) -> str:
    if pd.isna(value):
        return "-"
    return value.strftime("%d/%m/%Y")


def format_long_date(value) -> str:
    if pd.isna(value):
        return "-"
    return value.strftime("%d %b %Y")


def get_label(row: pd.Series) -> str:
    pieces = [row["Proyecto"]]
    if row["Unidad"]:
        pieces.append(row["Unidad"])
    elif row["Torre_display"] != "-":
        pieces.append(f"T{row['Torre_display']}")
    return " - ".join(pieces)


def get_tooltip(row: pd.Series) -> str:
    parts = [
        row["Proyecto"],
        f"Torre/ZC: {row['Torre_display']}",
        f"Unidad: {row['Unidad'] or '-'}",
        f"Fecha actual: {format_short_date(row['Fecha actual'])}",
        f"Fecha anterior: {format_short_date(row['Fecha anterior'])}",
        "Firmado" if row["Firmado"] == 1 else "Sin firma",
    ]
    return " | ".join(parts)


def init_month_state(df: pd.DataFrame) -> None:
    if "ctos_view_year" in st.session_state and "ctos_view_month" in st.session_state:
        return

    today = datetime.now(BOGOTA_TZ).date()
    min_date = df["Fecha actual"].dropna().min()

    if pd.isna(min_date):
        st.session_state["ctos_view_year"] = today.year
        st.session_state["ctos_view_month"] = today.month
        return

    st.session_state["ctos_view_year"] = today.year
    st.session_state["ctos_view_month"] = today.month


def shift_month(delta: int) -> None:
    year = st.session_state["ctos_view_year"]
    month = st.session_state["ctos_view_month"] + delta

    while month > 12:
        month -= 12
        year += 1
    while month < 1:
        month += 12
        year -= 1

    st.session_state["ctos_view_year"] = year
    st.session_state["ctos_view_month"] = month


def build_calendar_html(data: pd.DataFrame, year: int, month: int) -> str:
    grouped = {}
    month_rows = data[data["Fecha actual"].notna()].copy()
    month_rows = month_rows[
        (month_rows["Fecha actual"].dt.year == year) & (month_rows["Fecha actual"].dt.month == month)
    ]

    for _, row in month_rows.iterrows():
        day = int(row["Fecha actual"].day)
        grouped.setdefault(day, []).append(row)

    for day in grouped:
        grouped[day] = sorted(grouped[day], key=lambda item: (item["Firmado"], item["Proyecto"], item["Torre_display"]))

    first_weekday, days_in_month = monthrange(year, month)
    first_weekday = (first_weekday + 1) % 7
    prev_month = 12 if month == 1 else month - 1
    prev_year = year - 1 if month == 1 else year
    prev_days = monthrange(prev_year, prev_month)[1]
    today = datetime.now(BOGOTA_TZ).date()

    cells = []

    for offset in range(first_weekday):
        day_num = prev_days - first_weekday + offset + 1
        cells.append(make_day_cell(day_num, other_month=True, is_today=False, events=[]))

    for day in range(1, days_in_month + 1):
        is_today = today == date(year, month, day)
        cells.append(make_day_cell(day, other_month=False, is_today=is_today, events=grouped.get(day, [])))

    while len(cells) % 7 != 0:
        cells.append(make_day_cell(len(cells) % 7 + 1, other_month=True, is_today=False, events=[]))

    return f"""
    <style>
    .cto-calendar {{
        background: #FFFFFF;
        border: 1px solid #E5E9F0;
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 1px 4px rgba(15, 23, 42, .05);
    }}
    .cto-cal-head {{
        display: grid;
        grid-template-columns: repeat(7, 1fr);
        border-bottom: 1px solid #E5E9F0;
        background: #F8FAFC;
    }}
    .cto-cal-day-name {{
        padding: 12px 6px;
        text-align: center;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: .06em;
        text-transform: uppercase;
        color: #9CA3AF;
    }}
    .cto-cal-grid {{
        display: grid;
        grid-template-columns: repeat(7, 1fr);
    }}
    .cto-cal-cell {{
        min-height: 128px;
        border-right: 1px solid #E5E9F0;
        border-bottom: 1px solid #E5E9F0;
        padding: 10px 8px 8px 8px;
        background: #FFFFFF;
    }}
    .cto-cal-cell:nth-child(7n) {{ border-right: none; }}
    .cto-cal-cell.other-month {{ background: #FAFBFC; }}
    .cto-cal-cell.today {{ background: #F4F8FD; }}
    .cto-day-num {{
        width: 24px;
        height: 24px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 12px;
        font-weight: 700;
        color: #6B7280;
        font-family: 'DM Mono', monospace;
        margin-bottom: 8px;
        border-radius: 999px;
    }}
    .cto-cal-cell.other-month .cto-day-num {{
        color: #C4CAD4;
    }}
    .cto-cal-cell.today .cto-day-num {{
        background: #7BA7D4;
        color: #FFFFFF;
    }}
    .cto-events {{
        display: flex;
        flex-direction: column;
        gap: 4px;
    }}
    .cto-event {{
        border-radius: 6px;
        padding: 4px 6px;
        font-size: 10.5px;
        font-weight: 600;
        line-height: 1.35;
        overflow: hidden;
        white-space: nowrap;
        text-overflow: ellipsis;
    }}
    .cto-event.signed {{
        background: #E4F4EE;
        border-left: 3px solid #6BBF9E;
        color: #3D8B6E;
    }}
    .cto-event.unsigned {{
        background: #F8E8E8;
        border-left: 3px solid #D98B8B;
        color: #A54D57;
    }}
    .cto-more {{
        font-size: 10.5px;
        color: #9CA3AF;
        padding: 1px 4px;
    }}
    </style>
    <div class="cto-calendar">
        <div class="cto-cal-head">
            {''.join(f'<div class="cto-cal-day-name">{day}</div>' for day in WEEKDAYS_ES)}
        </div>
        <div class="cto-cal-grid">
            {''.join(cells)}
        </div>
    </div>
    """


def make_day_cell(day: int, other_month: bool, is_today: bool, events: list[pd.Series]) -> str:
    classes = ["cto-cal-cell"]
    if other_month:
        classes.append("other-month")
    if is_today:
        classes.append("today")

    events_html = ""
    if not other_month and events:
        blocks = []
        for row in events[:3]:
            status_class = "signed" if int(row["Firmado"]) == 1 else "unsigned"
            blocks.append(
                f'<div class="cto-event {status_class}" title="{html.escape(get_tooltip(row))}">'
                f"{html.escape(get_label(row))}"
                f"</div>"
            )
        if len(events) > 3:
            blocks.append(f'<div class="cto-more">+{len(events) - 3} mas</div>')
        events_html = f'<div class="cto-events">{"".join(blocks)}</div>'

    return (
        f'<div class="{" ".join(classes)}">'
        f'<div class="cto-day-num">{day}</div>'
        f"{events_html}"
        f"</div>"
    )


def build_pending_table_html(data: pd.DataFrame) -> str:
    if data.empty:
        return (
            '<div class="info-note" style="margin-bottom:0;">'
            "No hay CTOs pendientes para los filtros seleccionados."
            "</div>"
        )

    rows_html = "".join(
        "<tr>"
        f"<td>{html.escape(row['Proyecto'])}</td>"
        f"<td>{html.escape(row['Torre_display'])}</td>"
        f"<td>{html.escape(row['Unidad'] or '-')}</td>"
        f"<td>{html.escape(format_short_date(row['Fecha actual']))}</td>"
        f"<td>{html.escape(format_short_date(row['Fecha anterior']))}</td>"
        "</tr>"
        for _, row in data.iterrows()
    )

    return (
        '<div class="table-shell">'
        '<table class="rt">'
        "<thead><tr>"
        "<th>Proyecto</th>"
        "<th>Torre / ZC</th>"
        "<th>Unidad</th>"
        "<th>Fecha actual</th>"
        "<th>Fecha anterior</th>"
        "</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        "</table>"
        "</div>"
    )


def render_dashboard(df: pd.DataFrame) -> None:
    inject_base_styles()
    render_page_heading(
        "Dashboard CTOs",
        "Seguimiento de firmas programadas por proyecto, con calendario mensual y detalle de pendientes.",
    )

    updated_at = datetime.fromtimestamp(EXCEL_PATH.stat().st_mtime, BOGOTA_TZ)
    st.markdown(
        f'<div class="info-note">Fuente: <strong>{EXCEL_PATH.name}</strong> · Tabla <strong>{TABLE_NAME}</strong> · '
        f'Actualizado: {updated_at.strftime("%d/%m/%Y %H:%M")}</div>',
        unsafe_allow_html=True,
    )

    projects = ["Todos"] + sorted(df["Proyecto"].dropna().unique().tolist())
    init_month_state(df)

    filter_col, month_col, legend_col = st.columns([2.2, 2.3, 1.9])

    with filter_col:
        selected_project = st.selectbox("Proyecto", projects, index=0)

    with month_col:
        nav_prev, nav_label, nav_next = st.columns([0.8, 2.2, 0.8])
        with nav_prev:
            st.write("")
            if st.button("<", key="ctos_prev_month"):
                shift_month(-1)
        with nav_label:
            st.markdown(
                f'<div class="month-chip">{MONTHS_ES[st.session_state["ctos_view_month"]]} {st.session_state["ctos_view_year"]}</div>',
                unsafe_allow_html=True,
            )
        with nav_next:
            st.write("")
            if st.button(">", key="ctos_next_month"):
                shift_month(1)

    with legend_col:
        st.markdown(
            """
            <div class="legend-wrap">
                <div class="legend-item"><span class="legend-dot legend-signed"></span>Firmado</div>
                <div class="legend-item"><span class="legend-dot legend-unsigned"></span>Sin firma</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    filtered = df.copy()
    if selected_project != "Todos":
        filtered = filtered[filtered["Proyecto"] == selected_project].copy()

    total = len(filtered)
    signed = int((filtered["Firmado"] == 1).sum())
    pending = int((filtered["Firmado"] == 0).sum())

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(kpi_card("Total CTOs", f"{total:,}", "Registros visibles con el filtro actual", "kp-blue"), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card("Firmados", f"{signed:,}", "CTOs con firma registrada", "kp-green"), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi_card("Pendientes", f"{pending:,}", "CTOs aun sin firma", "kp-red"), unsafe_allow_html=True)

    st.markdown(
        section_header(
            "Calendario de CTOs",
            "Cada evento se ubica en la fecha actual programada. Pase el cursor sobre un registro para ver el detalle.",
        ),
        unsafe_allow_html=True,
    )
    calendar_html = build_calendar_html(
        filtered,
        st.session_state["ctos_view_year"],
        st.session_state["ctos_view_month"],
    )
    st.components.v1.html(calendar_html, height=840, scrolling=False)

    pending_df = filtered[filtered["Firmado"] == 0].copy()
    pending_df = pending_df.sort_values(
        ["Fecha actual", "Proyecto", "Torre_display", "Unidad"],
        na_position="last",
    )

    st.markdown(
        section_header("CTOs sin firma", ""),
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div style="margin:0 0 12px 0;"><span class="pill-count">{len(pending_df):,} pendientes</span></div>',
        unsafe_allow_html=True,
    )
    st.markdown(build_pending_table_html(pending_df), unsafe_allow_html=True)


def main() -> None:
    if not EXCEL_PATH.exists():
        st.error(f"No se encontro el archivo {EXCEL_PATH.name} en el repositorio.")
        return

    try:
        signature = get_excel_signature(EXCEL_PATH)
        df = load_cto_data(signature)
    except Exception as exc:
        st.error(f"No fue posible cargar la informacion de CTOs: {exc}")
        return

    render_dashboard(df)


if __name__ == "__main__":
    main()

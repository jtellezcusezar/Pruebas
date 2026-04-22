import html
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from zoneinfo import ZoneInfo

try:
    from streamlit_calendar import calendar as render_calendar
except ImportError:
    render_calendar = None


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

EXCEL_PATH = Path("CTOs.xlsx")
TABLE_NAME = "CTO"
BOGOTA_TZ = ZoneInfo("America/Bogota")


def inject_base_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
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
            font-weight: 600;
            margin-bottom: 18px;
        }
        div[data-testid="stSelectbox"] > label {
            font-size: 11px !important;
            font-weight: 700 !important;
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
            margin-top: 28px;
            flex-wrap: wrap;
        }
        .legend-item {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            color: #6B7280;
            font-size: 12px;
            font-weight: 700;
        }
        .legend-dot {
            width: 10px;
            height: 10px;
            border-radius: 3px;
        }
        .legend-signed { background: #6BBF9E; }
        .legend-unsigned { background: #D98B8B; }
        .detail-card {
            background: #FFFFFF;
            border: 1px solid #E5E9F0;
            border-radius: 14px;
            padding: 16px 18px;
            box-shadow: 0 1px 4px rgba(15, 23, 42, .05);
            margin-top: 12px;
        }
        .detail-title {
            font-size: 16px;
            font-weight: 800;
            color: #111827;
            margin-bottom: 6px;
        }
        .detail-sub {
            font-size: 12px;
            color: #6B7280;
            margin-bottom: 14px;
        }
        .detail-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 14px;
        }
        .detail-label {
            font-size: 10px;
            text-transform: uppercase;
            letter-spacing: .06em;
            color: #9CA3AF;
            font-weight: 800;
            margin-bottom: 6px;
        }
        .detail-value {
            font-size: 13px;
            color: #111827;
            line-height: 1.5;
            white-space: pre-wrap;
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
            font-weight: 800;
        }
        @media (max-width: 900px) {
            .detail-grid { grid-template-columns: 1fr; }
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
                return pd.DataFrame(data[1:], columns=data[0])
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
    df = df.rename(columns={"Torre/ZC": "Torre"})

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
    df["Proyecto_key"] = df["Proyecto"].str.strip()
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
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    return str(int(number)) if number.is_integer() else str(number)


def format_short_date(value) -> str:
    if pd.isna(value):
        return "-"
    return value.strftime("%d/%m/%Y")


def as_text_list(values: list[str], fallback: str = "-") -> str:
    clean = [value for value in values if value]
    return ", ".join(clean) if clean else fallback


def get_nearest_date(df: pd.DataFrame) -> pd.Timestamp | None:
    valid_dates = df["Fecha actual"].dropna()
    if valid_dates.empty:
        return None

    today = pd.Timestamp(datetime.now(BOGOTA_TZ).date())
    return min(valid_dates, key=lambda item: abs(item.normalize() - today))


def ensure_view_date(df: pd.DataFrame, selected_project: str) -> None:
    prev_project = st.session_state.get("ctos_selected_project_prev")
    current_view = st.session_state.get("ctos_view_date")

    if prev_project == selected_project and current_view:
        return

    target_df = df if selected_project == "Todos" else df[df["Proyecto"] == selected_project]
    nearest_date = get_nearest_date(target_df)

    if nearest_date is not None:
        st.session_state["ctos_view_date"] = nearest_date.strftime("%Y-%m-%d")
    elif current_view is None:
        st.session_state["ctos_view_date"] = datetime.now(BOGOTA_TZ).strftime("%Y-%m-%d")

    st.session_state["ctos_selected_project_prev"] = selected_project


def aggregate_calendar_events(df: pd.DataFrame) -> list[dict]:
    rows = df[df["Fecha actual"].notna()].copy()
    if rows.empty:
        return []

    grouped = (
        rows.groupby(["Proyecto", rows["Fecha actual"].dt.normalize()], dropna=False)
        .apply(build_event_payload)
        .tolist()
    )
    return grouped


def build_event_payload(group: pd.DataFrame) -> dict:
    project = str(group.iloc[0]["Proyecto"])
    event_date = group.iloc[0]["Fecha actual"].normalize()
    signed_count = int((group["Firmado"] == 1).sum())
    pending_count = int((group["Firmado"] == 0).sum())
    towers = sorted({str(value).strip() for value in group["Torre_display"].tolist() if str(value).strip() and str(value).strip() != "-"})
    units = sorted({str(value).strip() for value in group["Unidad"].tolist() if str(value).strip()})

    all_signed = pending_count == 0
    detail_lines = []
    if towers:
        detail_lines.append(f"Torres/ZC: {', '.join(towers)}")
    if units:
        detail_lines.append(f"Unidades: {', '.join(units)}")

    return {
        "id": f"{project}|{event_date.strftime('%Y-%m-%d')}",
        "title": project,
        "start": event_date.strftime("%Y-%m-%d"),
        "allDay": True,
        "backgroundColor": "#E4F4EE" if all_signed else "#F8E8E8",
        "borderColor": "#6BBF9E" if all_signed else "#D98B8B",
        "textColor": "#2F6F58" if all_signed else "#8F3942",
        "extendedProps": {
            "proyecto": project,
            "fecha_actual": format_short_date(group.iloc[0]["Fecha actual"]),
            "fecha_anterior_min": format_short_date(group["Fecha anterior"].min()),
            "fecha_anterior_max": format_short_date(group["Fecha anterior"].max()),
            "firmados": signed_count,
            "pendientes": pending_count,
            "torres": towers,
            "unidades": units,
            "detalle": "\n".join(detail_lines) if detail_lines else "Sin torres o unidades detalladas",
            "estado": "Firmado" if all_signed else "Con pendientes",
            "total_registros": len(group),
        },
    }


def calendar_custom_css() -> str:
    return """
    .fc {
        font-family: 'Manrope', sans-serif;
        color: #111827;
    }
    .fc-toolbar.fc-header-toolbar {
        margin-bottom: 1rem;
        gap: 8px;
        flex-wrap: wrap;
    }
    .fc-toolbar-title {
        font-size: 1.1rem;
        font-weight: 800;
        color: #111827;
    }
    .fc-button {
        background: #FFFFFF !important;
        color: #4A7BA8 !important;
        border: 1px solid #D5E2F0 !important;
        border-radius: 10px !important;
        box-shadow: none !important;
        font-weight: 700 !important;
        text-transform: capitalize !important;
    }
    .fc-button-primary:not(:disabled).fc-button-active,
    .fc-button-primary:not(:disabled):active,
    .fc-button:hover {
        background: #EEF3FA !important;
        border-color: #7BA7D4 !important;
        color: #355D8A !important;
    }
    .fc-theme-standard td,
    .fc-theme-standard th,
    .fc-theme-standard .fc-scrollgrid {
        border-color: #E5E9F0;
    }
    .fc-col-header-cell-cushion {
        color: #9CA3AF;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: .06em;
        text-transform: uppercase;
        padding: 10px 4px;
    }
    .fc-daygrid-day-number {
        color: #6B7280;
        font-family: 'DM Mono', monospace;
        font-size: 12px;
        font-weight: 700;
    }
    .fc-day-today {
        background: #F4F8FD !important;
    }
    .fc-daygrid-event {
        border-radius: 8px;
        padding: 2px 4px;
        font-size: 11px;
        font-weight: 700;
    }
    .fc-event-title {
        font-weight: 800;
    }
    .fc-daygrid-more-link {
        color: #4A7BA8;
        font-weight: 700;
    }
    """


def build_pending_table_html(data: pd.DataFrame) -> str:
    if data.empty:
        return '<div class="info-note" style="margin-bottom:0;">No hay CTOs pendientes para los filtros seleccionados.</div>'

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
        '<div class="table-shell"><table class="rt"><thead><tr>'
        "<th>Proyecto</th><th>Torre / ZC</th><th>Unidad</th><th>Fecha actual</th><th>Fecha anterior</th>"
        f"</tr></thead><tbody>{rows_html}</tbody></table></div>"
    )


def render_selected_event_detail(calendar_response: dict | None) -> None:
    if not calendar_response or calendar_response.get("callback") != "eventClick":
        st.markdown(
            '<div class="info-note">Haz clic sobre un proyecto en el calendario para ver el detalle de torres, unidades y firmas de esa fecha.</div>',
            unsafe_allow_html=True,
        )
        return

    event = calendar_response.get("eventClick", {}).get("event", {})
    details = event.get("extendedProps", {})
    title = details.get("proyecto", event.get("title", "Detalle"))
    subtitle = f"Fecha actual: {details.get('fecha_actual', '-')} · Registros: {details.get('total_registros', 0)}"

    st.markdown(
        f"""
        <div class="detail-card">
            <div class="detail-title">{html.escape(title)}</div>
            <div class="detail-sub">{html.escape(subtitle)}</div>
            <div class="detail-grid">
                <div>
                    <div class="detail-label">Torres / ZC</div>
                    <div class="detail-value">{html.escape(as_text_list(details.get('torres', [])))}</div>
                </div>
                <div>
                    <div class="detail-label">Unidades</div>
                    <div class="detail-value">{html.escape(as_text_list(details.get('unidades', [])))}</div>
                </div>
                <div>
                    <div class="detail-label">Firmados</div>
                    <div class="detail-value">{details.get('firmados', 0)}</div>
                </div>
                <div>
                    <div class="detail-label">Pendientes</div>
                    <div class="detail-value">{details.get('pendientes', 0)}</div>
                </div>
                <div>
                    <div class="detail-label">Fecha anterior mas temprana</div>
                    <div class="detail-value">{html.escape(details.get('fecha_anterior_min', '-'))}</div>
                </div>
                <div>
                    <div class="detail-label">Fecha anterior mas tardia</div>
                    <div class="detail-value">{html.escape(details.get('fecha_anterior_max', '-'))}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_calendar_section(filtered: pd.DataFrame) -> None:
    if render_calendar is None:
        st.error("Falta instalar `streamlit-calendar`. Agrega esta dependencia al entorno antes de probar esta vista.")
        st.code("pip install streamlit-calendar")
        return

    ensure_view_date(filtered, st.session_state["ctos_selected_project"])
    events = aggregate_calendar_events(filtered)

    calendar_options = {
        "initialView": "dayGridMonth",
        "initialDate": st.session_state.get("ctos_view_date"),
        "locale": "es",
        "height": 760,
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,listMonth",
        },
        "buttonText": {
            "today": "Hoy",
            "month": "Mes",
            "list": "Lista",
        },
        "dayMaxEventRows": 3,
        "eventDisplay": "block",
        "displayEventTime": False,
        "fixedWeekCount": False,
        "showNonCurrentDates": True,
    }

    calendar_response = render_calendar(
        events=events,
        options=calendar_options,
        custom_css=calendar_custom_css(),
        callbacks=["eventClick"],
        key=f"ctos_calendar_{st.session_state.get('ctos_view_date', 'default')}_{st.session_state['ctos_selected_project']}",
    )
    render_selected_event_detail(calendar_response)


def render_dashboard(df: pd.DataFrame) -> None:
    inject_base_styles()
    render_page_heading(
        "Dashboard CTOs",
        "Seguimiento de firmas programadas por proyecto, agrupando eventos por fecha para una lectura mas limpia.",
    )

    updated_at = datetime.fromtimestamp(EXCEL_PATH.stat().st_mtime, BOGOTA_TZ)
    st.markdown(
        f'<div class="info-note">Fuente: <strong>{EXCEL_PATH.name}</strong> · Tabla <strong>{TABLE_NAME}</strong> · '
        f'Actualizado: {updated_at.strftime("%d/%m/%Y %H:%M")}</div>',
        unsafe_allow_html=True,
    )

    projects = ["Todos"] + sorted(df["Proyecto"].dropna().unique().tolist())
    default_project = st.session_state.get("ctos_selected_project", "Todos")
    if default_project not in projects:
        default_project = "Todos"

    filter_col, legend_col = st.columns([2.2, 1.8])
    with filter_col:
        selected_project = st.selectbox(
            "Proyecto",
            projects,
            index=projects.index(default_project),
            key="ctos_selected_project",
        )
    with legend_col:
        st.markdown(
            """
            <div class="legend-wrap">
                <div class="legend-item"><span class="legend-dot legend-signed"></span>Todo firmado</div>
                <div class="legend-item"><span class="legend-dot legend-unsigned"></span>Con pendientes</div>
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
            "Cada evento representa un proyecto en una fecha. Haz clic sobre el evento para ver torres y unidades asociadas.",
        ),
        unsafe_allow_html=True,
    )
    render_calendar_section(filtered)

    pending_df = filtered[filtered["Firmado"] == 0].copy()
    pending_df = pending_df.sort_values(
        ["Fecha actual", "Proyecto", "Torre_display", "Unidad"],
        na_position="last",
    )

    st.markdown(section_header("CTOs sin firma"), unsafe_allow_html=True)
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

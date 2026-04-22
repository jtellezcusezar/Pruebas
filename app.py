import html
import json
from calendar import monthrange
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
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

EXCEL_PATH = Path("CTOs.xlsx")
TABLE_NAME = "CTO"
BOGOTA_TZ = ZoneInfo("America/Bogota")
WEEKDAY_LABELS = ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"]


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


def compact_numeric_tokens(values: list[str], label: str, max_items: int = 12) -> str:
    clean = [value for value in values if value]
    if not clean:
        return "-"

    numeric = []
    text = []
    for value in clean:
        stripped = str(value).strip()
        try:
            number = float(stripped)
            numeric.append(str(int(number)) if number.is_integer() else stripped)
        except ValueError:
            text.append(stripped)

    ordered = numeric + sorted(set(text))
    unique_values = list(dict.fromkeys(ordered))
    shown = unique_values[:max_items]
    suffix = f" +{len(unique_values) - max_items}" if len(unique_values) > max_items else ""
    return f"{label}: {', '.join(shown)}{suffix}"


def compact_unit_labels(values: list[str], max_items: int = 12) -> str:
    clean = [value.strip() for value in values if value and value.strip()]
    if not clean:
        return "-"

    groups: dict[str, list[str]] = {}
    for value in clean:
        parts = value.split(" ", 1)
        prefix = parts[0]
        suffix = parts[1] if len(parts) > 1 else value
        groups.setdefault(prefix, []).append(suffix)

    chunks = []
    for prefix, suffixes in groups.items():
        label = "Locales" if prefix.lower() == "local" else "Aptos" if prefix.lower() == "apto" else prefix
        chunks.append(compact_numeric_tokens(suffixes, label, max_items=max_items))
    return " | ".join(chunks)


def summarize_project_group(project: str, group: pd.DataFrame) -> dict:
    towers = sorted(
        {
            str(value).strip()
            for value in group["Torre_display"].tolist()
            if str(value).strip() and str(value).strip() != "-"
        }
    )
    apartments = sorted(
        {
            format_number(value)
            for value in group["Apartamento"].dropna().tolist()
        }
    )
    locals_list = sorted(
        {
            format_number(value)
            for value in group["Local"].dropna().tolist()
        }
    )

    has_apartments = bool(apartments)
    project_label = project
    if has_apartments and towers:
        project_label = f"{project} · Torre {', '.join(towers[:8])}"
        if len(towers) > 8:
            project_label += f" +{len(towers) - 8}"

    return {
        "proyecto": project_label,
        "torres": compact_numeric_tokens(towers, "Torres", max_items=10) if towers and not has_apartments else "",
        "apartamentos": compact_numeric_tokens(apartments, "Apartamentos", max_items=10) if apartments else "",
        "locales": compact_numeric_tokens(locals_list, "Locales", max_items=10) if locals_list else "",
        "firmados": int((group["Firmado"] == 1).sum()),
        "total": int(len(group)),
    }


def get_nearest_date(df: pd.DataFrame) -> pd.Timestamp | None:
    valid_dates = df["Fecha actual"].dropna()
    if valid_dates.empty:
        return None

    today = pd.Timestamp(datetime.now(BOGOTA_TZ).date())
    return min(valid_dates, key=lambda item: abs(item.normalize() - today))


def aggregate_daily_data(df: pd.DataFrame) -> pd.DataFrame:
    rows = df[df["Fecha actual"].notna()].copy()
    if rows.empty:
        return pd.DataFrame()

    rows["Fecha_dia"] = rows["Fecha actual"].dt.normalize()
    grouped_rows = []

    for event_date, group in rows.groupby("Fecha_dia", sort=True):
        project_summary = []
        for project_name, project_group in group.groupby("Proyecto", sort=True):
            project_summary.append(summarize_project_group(str(project_name), project_group))

        towers = sorted(
            {
                str(value).strip()
                for value in group["Torre_display"].tolist()
                if str(value).strip() and str(value).strip() != "-"
            }
        )
        units = sorted({str(value).strip() for value in group["Unidad"].tolist() if str(value).strip()})
        total = int(len(group))
        signed = int((group["Firmado"] == 1).sum())
        pending = int((group["Firmado"] == 0).sum())
        pending_ratio = pending / total if total else 0

        grouped_rows.append(
            {
                "Fecha_dia": event_date,
                "fecha_iso": event_date.strftime("%Y-%m-%d"),
                "fecha_label": format_short_date(event_date),
                "total": total,
                "firmados": signed,
                "pendientes": pending,
                "pending_ratio": pending_ratio,
                "intensity": pending * 100 + total,
                "project_blocks": project_summary,
                "torres": towers,
                "unidades": units,
                "torres_resumen": compact_numeric_tokens(towers, "Torres/ZC"),
                "unidades_resumen": compact_unit_labels(units),
                "fecha_anterior_min": format_short_date(group["Fecha anterior"].min()),
                "fecha_anterior_max": format_short_date(group["Fecha anterior"].max()),
                "all_signed": pending == 0,
            }
        )

    return pd.DataFrame(grouped_rows)


def event_status_label(row: pd.Series) -> str:
    if row["pendientes"] == 0:
        return "Todo firmado"
    if row["firmados"] == 0:
        return "Todo pendiente"
    return "Mixto"


def build_heatmap_series(agg_df: pd.DataFrame) -> list[dict]:
    if agg_df.empty:
        return []

    items = []
    for _, row in agg_df.iterrows():
        item = {
            "value": [row["fecha_iso"], int(row["intensity"])],
            "fecha": row["fecha_label"],
            "total": int(row["total"]),
            "firmados": int(row["firmados"]),
            "pendientes": int(row["pendientes"]),
            "estado": event_status_label(row),
            "projectBlocks": row["project_blocks"],
            "torres": row["torres"],
            "unidades": row["unidades"],
            "torresResumen": row["torres_resumen"],
            "unidadesResumen": row["unidades_resumen"],
            "fechaAnteriorMin": row["fecha_anterior_min"],
            "fechaAnteriorMax": row["fecha_anterior_max"],
            "allSigned": bool(row["all_signed"]),
        }
        if bool(row["all_signed"]):
            item["itemStyle"] = {
                "color": "#BFE3CF",
                "borderColor": "#6BBF9E",
                "borderWidth": 1.2,
            }
        items.append(item)
    return items


def get_visual_bounds(agg_df: pd.DataFrame) -> tuple[int, int]:
    if agg_df.empty:
        return 0, 1

    min_value = int(agg_df["intensity"].min())
    max_value = int(agg_df["intensity"].max())

    if min_value == max_value:
        return min_value, min_value + 1

    return min_value, max_value


def build_month_matrix_option(agg_df: pd.DataFrame, year: int, month: int) -> tuple[dict, int]:
    visual_min, visual_max = get_visual_bounds(agg_df)
    first_weekday, days_in_month = monthrange(year, month)
    week_count = (first_weekday + days_in_month + 6) // 7
    week_labels = [f"Semana {index}" for index in range(1, week_count + 1)]
    row_positions = {label: idx for idx, label in enumerate(week_labels)}

    data = []
    valid_positions = {}
    for day in range(1, days_in_month + 1):
        weekday_idx = (first_weekday + day - 1) % 7
        week_idx = (first_weekday + day - 1) // 7
        valid_positions[(week_idx, weekday_idx)] = day
        data.append(
            {
                "value": [weekday_idx, row_positions[week_labels[week_idx]], 0],
                "fecha": f"{day:02d}/{month:02d}/{year}",
                "emptyCell": True,
                "itemStyle": {
                    "color": "#F8FAFC",
                    "borderColor": "#D1DAE6",
                    "borderWidth": 1.2,
                },
                "label": {
                    "show": True,
                    "formatter": str(day),
                    "color": "#9CA3AF",
                    "fontFamily": "DM Mono, monospace",
                    "fontSize": 10,
                    "fontWeight": 600,
                    "position": "insideTopLeft",
                    "padding": [4, 0, 0, 4],
                },
            }
        )

    for _, row in agg_df.iterrows():
        current_date = row["Fecha_dia"]
        weekday_idx = current_date.weekday()
        week_idx = (first_weekday + current_date.day - 1) // 7
        item = {
            "value": [weekday_idx, row_positions[week_labels[week_idx]], int(row["intensity"])],
            "fecha": row["fecha_label"],
            "total": int(row["total"]),
            "firmados": int(row["firmados"]),
            "pendientes": int(row["pendientes"]),
            "estado": event_status_label(row),
            "projectBlocks": row["project_blocks"],
            "fechaAnteriorMin": row["fecha_anterior_min"],
            "fechaAnteriorMax": row["fecha_anterior_max"],
            "allSigned": bool(row["all_signed"]),
            "emptyCell": False,
        }
        if bool(row["all_signed"]):
            item["itemStyle"] = {
                "color": "#BFE3CF",
                "borderColor": "#6BBF9E",
                "borderWidth": 1.2,
            }
        data.append(item)

    option = {
        "backgroundColor": "rgba(0,0,0,0)",
        "title": {
            "text": f"{MONTHS_ES[month]} {year}",
            "left": "center",
            "top": 8,
            "textStyle": {
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 16,
                "fontWeight": 800,
                "color": "#111827",
            },
        },
        "grid": {
            "top": 58,
            "left": 70,
            "right": 16,
            "bottom": 18,
            "containLabel": False,
        },
        "tooltip": {
            "confine": True,
            "extraCssText": "max-width:380px; white-space:normal; box-shadow:0 10px 24px rgba(15,23,42,.28); border-radius:12px;",
            "position": """__JS__function (point, params, dom, rect, size) {
                const boxWidth = size.contentSize[0];
                const boxHeight = size.contentSize[1];
                const viewWidth = size.viewSize[0];
                const viewHeight = size.viewSize[1];
                let x = point[0] + 14;
                let y = point[1] + 14;
                if (x + boxWidth > viewWidth - 12) x = point[0] - boxWidth - 14;
                if (x < 12) x = 12;
                if (y + boxHeight > viewHeight - 12) y = point[1] - boxHeight - 14;
                if (y < 12) y = 12;
                return [x, y];
            }""",
            "backgroundColor": "#111827",
            "borderColor": "#334155",
            "textStyle": {
                "color": "#F8FAFC",
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 12,
            },
            "formatter": """__JS__function (params) {
                const d = params.data || {};
                if (d.emptyCell) return '';
                const blocks = (d.projectBlocks || []).map((block, idx) => {
                    const lines = [
                        '<div style="font-weight:800;color:#F8FAFC;">Proyecto: ' + (block.proyecto || '-') + '</div>'
                    ];
                    if (block.torres) lines.push('<div>' + block.torres + '</div>');
                    if (block.apartamentos) lines.push('<div>' + block.apartamentos + '</div>');
                    if (block.locales) lines.push('<div>' + block.locales + '</div>');
                    lines.push('<div>Fecha anterior: <b>' + (d.fechaAnteriorMin || '-') + '</b>'
                        + ((d.fechaAnteriorMin !== d.fechaAnteriorMax) ? ' a <b>' + (d.fechaAnteriorMax || '-') + '</b>' : '')
                        + '</div>');
                    const separator = idx === 0 ? '' : 'margin-top:8px;padding-top:8px;border-top:1px solid rgba(148,163,184,.28);';
                    return '<div style="' + separator + '">' + lines.join('') + '</div>';
                }).join('');
                return '<div style="font-family:Manrope,sans-serif;line-height:1.5;">'
                    + '<div style="font-weight:800;margin-bottom:6px;">' + (d.fecha || '-') + '</div>'
                    + blocks
                    + '</div>';
            }""",
        },
        "xAxis": {
            "type": "category",
            "data": WEEKDAY_LABELS,
            "position": "top",
            "axisLine": {"show": False},
            "axisTick": {"show": False},
            "axisLabel": {
                "color": "#6B7280",
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 12,
                "fontWeight": 700,
                "margin": 12,
            },
            "splitArea": {"show": False},
        },
        "yAxis": {
            "type": "category",
            "data": week_labels,
            "inverse": True,
            "axisLine": {"show": False},
            "axisTick": {"show": False},
            "axisLabel": {
                "show": False,
            },
            "splitArea": {"show": False},
        },
        "visualMap": {
            "min": visual_min,
            "max": visual_max,
            "show": False,
            "inRange": {
                "color": ["#B5C8E3", "#9AB3D7", "#7F9ECB", "#6489BF", "#4A74B3", "#2F5FA7"],
            },
        },
        "graphic": [
            {
                "type": "text",
                "left": 16,
                "top": "middle",
                "style": {
                    "text": MONTHS_ES[month][:3],
                    "fill": "#6B7280",
                    "font": "700 14px 'Manrope'",
                },
            }
        ],
        "series": [
            {
                "type": "heatmap",
                "data": data,
                "label": {
                    "show": True,
                    "formatter": """__JS__function (params) {
                        const d = params.data || {};
                        if (d.emptyCell) {
                            return params.data.label ? params.data.label.formatter : '';
                        }
                        return '';
                    }""",
                },
                "itemStyle": {
                    "borderColor": "#D1DAE6",
                    "borderWidth": 1.2,
                    "borderRadius": 0,
                },
                "emphasis": {
                    "itemStyle": {
                        "shadowBlur": 10,
                        "shadowColor": "rgba(74,123,168,0.25)",
                        "borderColor": "#4A7BA8",
                    }
                },
            }
        ],
    }
    height = 84 + (week_count * 44)
    return option, height


def render_echarts(option: dict, height: int = 420) -> None:
    chart_id = f"echarts-{uuid4().hex}"
    option_json = json.dumps(option, ensure_ascii=False)
    html_code = f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
      html, body {{
        margin: 0;
        padding: 0;
        font-family: 'Manrope', sans-serif;
        background: transparent;
      }}
    </style>
    <div id="{chart_id}" style="width:100%;height:{height}px;"></div>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
    <script>
      function reviveJsFns(obj) {{
        if (Array.isArray(obj)) {{
          return obj.map(reviveJsFns);
        }}
        if (obj && typeof obj === 'object') {{
          for (const key of Object.keys(obj)) {{
            const value = obj[key];
            if (typeof value === 'string' && value.startsWith('__JS__')) {{
              obj[key] = eval('(' + value.slice(6) + ')');
            }} else {{
              obj[key] = reviveJsFns(value);
            }}
          }}
        }}
        return obj;
      }}
      const chart = echarts.init(document.getElementById('{chart_id}'));
      const option = reviveJsFns({option_json});
      chart.setOption(option);
      window.addEventListener('resize', () => chart.resize());
    </script>
    """
    components.html(html_code, height=height, scrolling=False)


def build_heatmap_option(agg_df: pd.DataFrame, year: int, selected_month: str) -> tuple[dict, int]:
    month_selected = selected_month != "Todos"
    if month_selected:
        return build_month_matrix_option(agg_df, year, int(selected_month))

    range_value = f"{year}-{selected_month}" if month_selected else str(year)
    visual_min, visual_max = get_visual_bounds(agg_df)
    height = 280

    option = {
        "backgroundColor": "rgba(0,0,0,0)",
        "title": {
            "text": f"{MONTHS_ES[int(selected_month)]} {year}" if month_selected else str(year),
            "left": "center",
            "top": 8,
            "textStyle": {
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 16,
                "fontWeight": 800,
                "color": "#111827",
            },
        },
        "tooltip": {
            "confine": True,
            "extraCssText": "max-width:380px; white-space:normal; box-shadow:0 10px 24px rgba(15,23,42,.28); border-radius:12px;",
            "position": """__JS__function (point, params, dom, rect, size) {
                const boxWidth = size.contentSize[0];
                const boxHeight = size.contentSize[1];
                const viewWidth = size.viewSize[0];
                const viewHeight = size.viewSize[1];
                let x = point[0] + 14;
                let y = point[1] + 14;
                if (x + boxWidth > viewWidth - 12) x = point[0] - boxWidth - 14;
                if (x < 12) x = 12;
                if (y + boxHeight > viewHeight - 12) y = point[1] - boxHeight - 14;
                if (y < 12) y = 12;
                return [x, y];
            }""",
            "backgroundColor": "#111827",
            "borderColor": "#334155",
            "textStyle": {
                "color": "#F8FAFC",
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 12,
            },
            "formatter": """__JS__function (params) {
                const d = params.data || {};
                const blocks = (d.projectBlocks || []).map((block, idx) => {
                    const lines = [
                        '<div style="font-weight:800;color:#F8FAFC;">Proyecto: ' + (block.proyecto || '-') + '</div>'
                    ];
                    if (block.torres) lines.push('<div>' + block.torres + '</div>');
                    if (block.apartamentos) lines.push('<div>' + block.apartamentos + '</div>');
                    if (block.locales) lines.push('<div>' + block.locales + '</div>');
                    lines.push('<div>Fecha anterior: <b>' + (d.fechaAnteriorMin || '-') + '</b>'
                        + ((d.fechaAnteriorMin !== d.fechaAnteriorMax) ? ' a <b>' + (d.fechaAnteriorMax || '-') + '</b>' : '')
                        + '</div>');
                    const separator = idx === 0 ? '' : 'margin-top:8px;padding-top:8px;border-top:1px solid rgba(148,163,184,.28);';
                    return '<div style="' + separator + '">'
                        + lines.join('')
                        + '</div>';
                }).join('');
                return '<div style="font-family:Manrope,sans-serif;line-height:1.5;">'
                    + '<div style="font-weight:800;margin-bottom:6px;">' + (d.fecha || params.value[0]) + '</div>'
                    + blocks
                    + '</div>';
            }""",
        },
        "visualMap": {
            "min": visual_min,
            "max": visual_max,
            "show": False,
            "inRange": {
                "color": ["#B5C8E3", "#9AB3D7", "#7F9ECB", "#6489BF", "#4A74B3", "#2F5FA7"],
            },
        },
        "calendar": {
            "top": 54,
            "left": 30,
            "right": 20,
            "range": range_value,
            "cellSize": ["auto", 17],
            "splitLine": {
                "show": True,
                "lineStyle": {
                    "color": "#B8C7D8",
                    "width": 1.2,
                },
            },
            "itemStyle": {
                "borderWidth": 1,
                "borderColor": "#E5E9F0",
                "color": "#FAFBFC",
            },
            "yearLabel": {"show": False},
            "monthLabel": {
                "nameMap": "es",
                "color": "#6B7280",
                "fontFamily": "Manrope, sans-serif",
                "fontWeight": 700,
                "margin": 16,
            },
            "dayLabel": {
                "firstDay": 0,
                "nameMap": ["Dom", "Lun", "Mar", "Mie", "Jue", "Vie", "Sab"],
                "color": "#9CA3AF",
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 11,
                "fontWeight": 700,
            },
        },
        "series": [
            {
                "type": "heatmap",
                "coordinateSystem": "calendar",
                "data": build_heatmap_series(agg_df),
                "label": {"show": False},
                "emphasis": {
                    "itemStyle": {
                        "shadowBlur": 10,
                        "shadowColor": "rgba(74,123,168,0.25)",
                        "borderColor": "#4A7BA8",
                    }
                },
            }
        ],
    }
    return option, height


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


def render_heatmap_section(filtered: pd.DataFrame, selected_year: int, selected_month: str) -> None:
    agg_df = aggregate_daily_data(filtered)
    if agg_df.empty:
        st.markdown(
            '<div class="info-note">No hay fechas programadas para los filtros seleccionados.</div>',
            unsafe_allow_html=True,
        )
        return

    if selected_month == "Todos":
        current_year = datetime.now(BOGOTA_TZ).year
        years_to_render = sorted(year for year in agg_df["Fecha_dia"].dt.year.unique().tolist() if year >= current_year)
        if not years_to_render:
            years_to_render = sorted(agg_df["Fecha_dia"].dt.year.unique().tolist())

        for year in years_to_render:
            year_df = agg_df[agg_df["Fecha_dia"].dt.year == year].copy()
            option, height = build_heatmap_option(year_df, year, selected_month)
            render_echarts(option, height=height)
    else:
        agg_df = agg_df[agg_df["Fecha_dia"].dt.year == selected_year].copy()
        agg_df = agg_df[agg_df["Fecha_dia"].dt.month == int(selected_month)].copy()

        if agg_df.empty:
            st.markdown(
                '<div class="info-note">No hay fechas programadas en el periodo seleccionado.</div>',
                unsafe_allow_html=True,
            )
            return

        left, center, right = st.columns([0.3, 0.4, 0.3])
        with center:
            option, height = build_heatmap_option(agg_df, selected_year, selected_month)
            render_echarts(option, height=height)


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

    valid_dates = df["Fecha actual"].dropna()
    nearest_date = get_nearest_date(df)
    nearest_year = nearest_date.year if nearest_date is not None else datetime.now(BOGOTA_TZ).year
    available_years = sorted(valid_dates.dt.year.unique().tolist()) if not valid_dates.empty else [nearest_year]
    month_options = ["Todos"] + [f"{month:02d}" for month in range(1, 13)]
    month_labels = {"Todos": "Todos"} | {f"{month:02d}": MONTHS_ES[month] for month in range(1, 13)}

    projects = ["Todos"] + sorted(df["Proyecto"].dropna().unique().tolist())
    default_project = st.session_state.get("ctos_selected_project", "Todos")
    if default_project not in projects:
        default_project = "Todos"
    default_year = st.session_state.get("ctos_selected_year", nearest_year)
    if default_year not in available_years:
        default_year = nearest_year
    default_month = st.session_state.get("ctos_selected_month", "Todos")
    if default_month not in month_options:
        default_month = "Todos"

    filter_col, year_col, month_col, legend_col = st.columns([2.0, 1.0, 1.2, 1.8])
    with filter_col:
        selected_project = st.selectbox(
            "Proyecto",
            projects,
            index=projects.index(default_project),
            key="ctos_selected_project",
        )
    with year_col:
        selected_year = st.selectbox(
            "Año",
            available_years,
            index=available_years.index(default_year),
            key="ctos_selected_year",
        )
    with month_col:
        selected_month = st.selectbox(
            "Mes",
            month_options,
            index=month_options.index(default_month),
            format_func=lambda value: month_labels[value],
            key="ctos_selected_month",
        )
    with legend_col:
        st.markdown(
            """
            <div class="legend-wrap">
                <div class="legend-item"><span class="legend-dot legend-signed"></span>Todo firmado</div>
                <div class="legend-item"><span class="legend-dot legend-unsigned"></span>Mayor presión</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    filtered = df.copy()
    if selected_project != "Todos":
        filtered = filtered[filtered["Proyecto"] == selected_project].copy()

    kpi_df = filtered.copy()
    if selected_month != "Todos":
        kpi_df = kpi_df[kpi_df["Fecha actual"].dt.year == selected_year].copy()
        kpi_df = kpi_df[kpi_df["Fecha actual"].dt.month == int(selected_month)].copy()

    total = len(kpi_df)
    signed = int((kpi_df["Firmado"] == 1).sum())
    pending = int((kpi_df["Firmado"] == 0).sum())

    if selected_month != "Todos":
        filtered = filtered[filtered["Fecha actual"].dt.month == int(selected_month)].copy()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(kpi_card("Total CTOs", f"{total:,}", "Registros visibles con el filtro actual", "kp-blue"), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card("Firmados", f"{signed:,}", "CTOs con firma registrada", "kp-green"), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi_card("Pendientes", f"{pending:,}", "CTOs aun sin firma", "kp-red"), unsafe_allow_html=True)

    st.markdown(
        section_header(
            "Heatmap de CTOs",
            "Vista anual por defecto. Al elegir un mes, el heatmap cambia a una lectura mensual con detalle agregado por fecha.",
        ),
        unsafe_allow_html=True,
    )
    render_heatmap_section(df if selected_project == "Todos" else df[df["Proyecto"] == selected_project].copy(), selected_year, selected_month)

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

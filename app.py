import json
from calendar import monthrange
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from zoneinfo import ZoneInfo


st.set_page_config(
    page_title="Notificaciones - Cusezar",
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

EXCEL_PATH = Path("Incidencias.xlsx")
BOGOTA_TZ = ZoneInfo("America/Bogota")
TARGET_PHASES = {"Cimentación", "Estructura", "Obra gris", "Obra blanca"}
STRUCTURE_PHASES = {"Cimentación", "Estructura"}
FINISH_PHASES = {"Obra gris", "Obra blanca"}
POSITIVE_STATES = {"CERRADA", "ANULADA"}
HEATMAP_CATEGORIES = ("Estructura", "Acabados")


def inject_base_styles() -> None:
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');
        html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
        #MainMenu, footer, header { visibility: hidden; }
        .block-container {
            padding-top: 2.7rem !important;
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
            font-size: 28px;
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
        .kpi-card {
            background: #FFFFFF;
            border: 1px solid #E5E9F0;
            border-radius: 14px;
            padding: 18px 20px;
            position: relative;
            overflow: hidden;
            box-shadow: 0 1px 4px rgba(15, 23, 42, .05);
            min-height: 116px;
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
        .kp-amber::before { background: #E8C17A; }
        .kp-slate::before { background: #94A3B8; }
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
        .kp-blue .kpi-value { color: #4A7BA8; }
        .kp-green .kpi-value { color: #3D8B6E; }
        .kp-amber .kpi-value { color: #C49A3C; }
        .kp-slate .kpi-value { color: #64748B; }
        .kpi-sub {
            font-size: 11px;
            color: #9CA3AF;
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
        .heatmap-shell {
            background: #FFFFFF;
            border: 1px solid #E5E9F0;
            border-radius: 16px;
            padding: 14px 14px 10px 14px;
            box-shadow: 0 1px 4px rgba(15, 23, 42, .05);
            height: 100%;
        }
        .mini-note {
            color: #94A3B8;
            font-size: 11px;
            margin-top: -8px;
            margin-bottom: 12px;
        }
        .heatmap-legend {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-top: 10px;
            padding: 0 6px 2px 6px;
            color: #64748B;
            font-size: 11px;
            font-weight: 600;
        }
        .heatmap-legend-bar {
            flex: 1;
            min-width: 120px;
            height: 10px;
            border-radius: 999px;
            background: linear-gradient(90deg, #B42318 0%, #F97066 25%, #F4E29A 50%, #7BC67B 75%, #157F3B 100%);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_page_heading(title: str, subtitle: str = "") -> None:
    subtitle_html = f'<div class="page-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="page-head"><div><div class="page-title">{title}</div>{subtitle_html}</div></div>',
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str = "") -> None:
    subtitle_html = f'<div class="section-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f'<div class="section-head"><div class="section-title">{title}</div>{subtitle_html}</div>',
        unsafe_allow_html=True,
    )


def kpi_card(label: str, value: str, sub: str, css: str) -> str:
    return (
        f'<div class="kpi-card {css}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'<div class="kpi-sub">{sub}</div>'
        f"</div>"
    )


def get_excel_signature(path: Path) -> tuple[int, int]:
    stat = path.stat()
    return stat.st_mtime_ns, stat.st_size


def format_spanish_date(dt_value: datetime) -> str:
    return (
        f"{dt_value.day} de {MONTHS_ES[dt_value.month].lower()} del "
        f"{dt_value.year} a las {dt_value.strftime('%H:%M')}"
    )


def get_excel_last_update_text(path: Path) -> str:
    workbook = load_workbook(path, read_only=True)
    try:
        updated_at = workbook.properties.modified or workbook.properties.created
    finally:
        workbook.close()

    if updated_at is None:
        fallback = datetime.fromtimestamp(path.stat().st_mtime, BOGOTA_TZ)
        return format_spanish_date(fallback)

    if updated_at.tzinfo is None:
        updated_at = updated_at.replace(tzinfo=timezone.utc)

    return format_spanish_date(updated_at.astimezone(BOGOTA_TZ))


def parse_text_date(series: pd.Series) -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    cleaned = cleaned.replace({"": pd.NA, "nan": pd.NA, "NaT": pd.NA})
    return pd.to_datetime(cleaned, format="%d/%m/%Y", errors="coerce")


@st.cache_data(show_spinner=False)
def load_incidents_data(_signature: tuple[int, int]) -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH, dtype=str)
    df.columns = [str(col).strip() for col in df.columns]

    for col in ["PROYECTO", "FASE", "ESTADO"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    for col in ["CREADA EN", "ATENDIDA EN", "FECHA ESTADO", "FECHA LÍMITE"]:
        df[col] = parse_text_date(df[col])

    filtered = df[df["FASE"].isin(TARGET_PHASES)].copy()
    filtered = filtered[filtered["PROYECTO"].ne("")].copy()

    filtered["CATEGORIA_HEATMAP"] = filtered["FASE"].map(
        lambda phase: "Estructura" if phase in STRUCTURE_PHASES else "Acabados"
    )
    filtered["ESTADO_NORM"] = filtered["ESTADO"].str.upper()
    filtered["PROJECT_SORT"] = filtered["PROYECTO"].str.casefold()
    return filtered


@st.cache_data(show_spinner=False)
def build_monthly_metrics(df: pd.DataFrame) -> pd.DataFrame:
    max_year = int(
        max(
            df["CREADA EN"].dropna().dt.year.max(),
            df["ATENDIDA EN"].dropna().dt.year.max() if df["ATENDIDA EN"].notna().any() else 0,
            df["FECHA ESTADO"].dropna().dt.year.max() if df["FECHA ESTADO"].notna().any() else 0,
            df["FECHA LÍMITE"].dropna().dt.year.max() if df["FECHA LÍMITE"].notna().any() else 0,
        )
    )
    min_year = int(df["CREADA EN"].dropna().dt.year.min())

    records: list[dict] = []
    ordered = df.sort_values(["PROJECT_SORT", "CATEGORIA_HEATMAP"]).copy()
    today = datetime.now(BOGOTA_TZ)
    current_year = today.year
    last_completed_month = max(today.month - 1, 0)

    for year in range(min_year, max_year + 1):
        if year > current_year:
            continue

        month_limit = 12 if year < current_year else last_completed_month
        if month_limit <= 0:
            continue

        for month in range(1, month_limit + 1):
            cutoff = pd.Timestamp(datetime(year, month, monthrange(year, month)[1]))
            created_in_year_until_cutoff = (
                ordered["CREADA EN"].notna()
                & ordered["CREADA EN"].dt.year.eq(year)
                & (ordered["CREADA EN"] <= cutoff)
            )
            limit_reached = ordered["FECHA LÍMITE"].notna() & (cutoff > ordered["FECHA LÍMITE"])
            limit_pending = ordered["FECHA LÍMITE"].notna() & (cutoff <= ordered["FECHA LÍMITE"])

            action_positive = created_in_year_until_cutoff & (
                (
                    ordered["ESTADO_NORM"].isin(POSITIVE_STATES)
                    & ordered["FECHA ESTADO"].notna()
                    & (ordered["FECHA ESTADO"] <= cutoff)
                )
                | (
                    ordered["ESTADO_NORM"].eq("VENCIDA")
                    & ordered["ATENDIDA EN"].notna()
                    & (ordered["ATENDIDA EN"] <= cutoff)
                )
            )

            receptive_no_action = created_in_year_until_cutoff & (
                (
                    ordered["ESTADO_NORM"].eq("VENCIDA")
                    & ordered["ATENDIDA EN"].isna()
                    & limit_reached
                )
                | (
                    ordered["ESTADO_NORM"].eq("ABIERTA")
                    & limit_reached
                )
            )

            partial_action = (
                created_in_year_until_cutoff
                & ordered["ESTADO_NORM"].eq("ABIERTA")
                & limit_pending
            )

            new_items = (
                ordered["ESTADO_NORM"].eq("ABIERTA")
                & ordered["CREADA EN"].notna()
                & ordered["CREADA EN"].dt.year.eq(year)
                & ordered["CREADA EN"].dt.month.eq(month)
            )

            month_df = ordered.loc[created_in_year_until_cutoff, ["PROYECTO", "CATEGORIA_HEATMAP"]].copy()
            month_df["accion_positiva"] = action_positive[created_in_year_until_cutoff].astype(int).to_numpy()
            month_df["receptiva_sin_accion"] = receptive_no_action[created_in_year_until_cutoff].astype(int).to_numpy()
            month_df["accion_parcial"] = partial_action[created_in_year_until_cutoff].astype(int).to_numpy()
            month_df["datos_nuevos_mes"] = new_items[created_in_year_until_cutoff].astype(int).to_numpy()

            grouped = (
                month_df.groupby(["CATEGORIA_HEATMAP", "PROYECTO"], as_index=False)[
                    ["accion_positiva", "receptiva_sin_accion", "accion_parcial", "datos_nuevos_mes"]
                ]
                .sum()
            )

            if grouped.empty:
                continue

            grouped["acumulado"] = (
                grouped["accion_positiva"]
                + grouped["receptiva_sin_accion"]
                + grouped["accion_parcial"]
            )
            grouped["porcentaje_atendidas"] = grouped["accion_positiva"] / grouped["acumulado"]
            grouped["year"] = year
            grouped["month"] = month
            records.extend(grouped.to_dict("records"))

    metrics = pd.DataFrame(records)
    if metrics.empty:
        return metrics

    metrics.rename(columns={"CATEGORIA_HEATMAP": "categoria", "PROYECTO": "proyecto"}, inplace=True)
    metrics["project_sort"] = metrics["proyecto"].str.casefold()
    return metrics


def build_heatmap_option(
    category_df: pd.DataFrame,
    category_name: str,
    year: int,
    selected_project: str,
) -> tuple[dict, int]:
    if category_df.empty:
        return {}, 280

    months = [MONTHS_ES[month] for month in range(1, 13)]
    projects = (
        category_df[["proyecto", "project_sort"]]
        .drop_duplicates()
        .sort_values(["project_sort", "proyecto"])
        ["proyecto"]
        .tolist()
    )

    project_to_idx = {project: idx for idx, project in enumerate(projects)}
    data_points = []

    for row in category_df.itertuples(index=False):
        if row.acumulado <= 0:
            continue
        item = {
            "value": [
                int(row.month - 1),
                int(project_to_idx[row.proyecto]),
                round(float(row.porcentaje_atendidas) * 100, 2),
                int(row.accion_positiva),
                int(row.receptiva_sin_accion),
                int(row.accion_parcial),
                int(row.acumulado),
                int(row.datos_nuevos_mes),
            ]
        }
        if selected_project != "Todos" and row.proyecto == selected_project:
            item["itemStyle"] = {"borderColor": "#1D4ED8", "borderWidth": 2}
        data_points.append(item)

    axis_formatter = "__JS__function(value){ return value; }"
    if selected_project != "Todos":
        axis_formatter = (
            "__JS__function(value){"
            f"if (value === {json.dumps(selected_project, ensure_ascii=False)}) "
            "{ return '{selected|' + value + '}'; }"
            "return value; }"
        )

    option = {
        "backgroundColor": "rgba(0,0,0,0)",
        "animation": False,
        "grid": {
            "top": 52,
            "left": 140,
            "right": 18,
            "bottom": 22,
            "containLabel": False,
        },
        "title": {
            "text": category_name,
            "left": "center",
            "top": 8,
            "textStyle": {
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 15,
                "fontWeight": 800,
                "color": "#111827",
            },
        },
        "tooltip": {
            "trigger": "item",
            "confine": True,
            "backgroundColor": "#0F172A",
            "borderColor": "#1E293B",
            "borderWidth": 1,
            "textStyle": {
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 12,
                "color": "#F8FAFC",
            },
            "extraCssText": "border-radius:12px; box-shadow:0 10px 24px rgba(15,23,42,.3);",
            "formatter": (
                "__JS__function(params){"
                "const v = params.value || [];"
                "const pct = Number(v[2] || 0).toFixed(1) + '%';"
                "return ["
                "'<div style=\"font-weight:800;margin-bottom:6px;\">' + params.name + '</div>',"
                "'<div style=\"margin-bottom:4px;\"><span style=\"color:#93C5FD;\">Proyecto:</span> ' + params.data.project + '</div>',"
                "'<div style=\"margin-bottom:4px;\"><span style=\"color:#93C5FD;\">Acción positiva:</span> ' + v[3] + '</div>',"
                "'<div style=\"margin-bottom:4px;\"><span style=\"color:#93C5FD;\">Receptiva sin acción:</span> ' + v[4] + '</div>',"
                "'<div style=\"margin-bottom:4px;\"><span style=\"color:#93C5FD;\">Acción parcial:</span> ' + v[5] + '</div>',"
                "'<div style=\"margin-bottom:4px;\"><span style=\"color:#93C5FD;\">Acumulado proyecto:</span> ' + v[6] + '</div>',"
                "'<div style=\"margin-bottom:4px;\"><span style=\"color:#93C5FD;\">Datos nuevos del mes:</span> ' + v[7] + '</div>',"
                "'<div><span style=\"color:#93C5FD;\">% atendidas:</span> ' + pct + '</div>'"
                "].join('');"
                "}"
            ),
        },
        "xAxis": {
            "type": "category",
            "data": months,
            "splitArea": {"show": False},
            "axisTick": {"show": False},
            "axisLine": {"lineStyle": {"color": "#CBD5E1"}},
            "axisLabel": {
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 11,
                "color": "#64748B",
            },
        },
        "yAxis": {
            "type": "category",
            "data": projects,
            "splitArea": {"show": False},
            "axisTick": {"show": False},
            "axisLine": {"show": False},
            "axisLabel": {
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 11,
                "color": "#334155",
                "width": 126,
                "overflow": "truncate",
                "formatter": axis_formatter,
                "rich": {
                    "selected": {
                        "color": "#1D4ED8",
                        "fontWeight": 800,
                        "backgroundColor": "#DBEAFE",
                        "padding": [4, 8],
                        "borderRadius": 10,
                    }
                },
            },
        },
        "visualMap": {
            "show": False,
            "dimension": 2,
            "min": 0,
            "max": 100,
            "calculable": False,
            "orient": "horizontal",
            "left": "center",
            "bottom": 0,
            "itemWidth": 110,
            "itemHeight": 10,
            "text": ["100%", "0%"],
            "textStyle": {
                "fontFamily": "Manrope, sans-serif",
                "fontSize": 11,
                "color": "#64748B",
            },
            "inRange": {
                "color": ["#B42318", "#F97066", "#F4E29A", "#7BC67B", "#157F3B"]
            },
        },
        "series": [
            {
                "name": f"{category_name} {year}",
                "type": "heatmap",
                "encode": {
                    "x": 0,
                    "y": 1,
                    "value": 2,
                    "tooltip": [2, 3, 4, 5, 6, 7],
                },
                "data": [
                    {
                        **item,
                        "project": projects[item["value"][1]],
                        "name": months[item["value"][0]],
                    }
                    for item in data_points
                ],
                "label": {"show": False},
                "itemStyle": {
                    "borderColor": "#FFFFFF",
                    "borderWidth": 1,
                    "borderRadius": 7,
                },
                "emphasis": {
                    "itemStyle": {
                        "shadowBlur": 12,
                        "shadowColor": "rgba(30,41,59,.18)",
                    }
                },
            }
        ],
    }
    height = max(300, 130 + (len(projects) * 26))
    return option, height


def render_echarts(option: dict, height: int = 420) -> None:
    chart_id = f"echarts-{uuid4().hex}"
    option_json = json.dumps(option, ensure_ascii=False)
    html_code = f"""
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


def get_available_years(metrics_df: pd.DataFrame) -> list[int]:
    return sorted(metrics_df["year"].dropna().astype(int).unique().tolist(), reverse=True)


def get_default_year(years: list[int]) -> int:
    today = datetime.now(BOGOTA_TZ)
    preferred_year = today.year if today.month > 1 else today.year - 1
    return preferred_year if preferred_year in years else years[0]


def get_default_month_name() -> str:
    today = datetime.now(BOGOTA_TZ)
    month_number = today.month - 1
    if month_number <= 0:
        month_number = 12
    return MONTHS_ES[month_number]


def format_pct(value: float | None) -> str:
    if value is None or pd.isna(value):
        return "0.0%"
    return f"{value * 100:.1f}%"


def get_reference_month_number(year_df: pd.DataFrame, fallback_year: int) -> int:
    valid = year_df[(year_df["acumulado"] > 0) | (year_df["datos_nuevos_mes"] > 0)]
    if not valid.empty:
        return int(valid["month"].max())
    today = datetime.now(BOGOTA_TZ)
    if fallback_year < today.year:
        return 12
    return max(1, today.month - 1)


def render_heatmap_legend() -> None:
    st.markdown(
        """
        <div class="heatmap-legend">
            <span>0%</span>
            <div class="heatmap-legend-bar"></div>
            <span>100%</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


inject_base_styles()

if not EXCEL_PATH.exists():
    st.error(f"No se encontró el archivo {EXCEL_PATH.name} en la carpeta actual.")
    st.stop()

excel_signature = get_excel_signature(EXCEL_PATH)
last_update_text = get_excel_last_update_text(EXCEL_PATH)
incidents_df = load_incidents_data(excel_signature)
metrics_df = build_monthly_metrics(incidents_df)

if metrics_df.empty:
    st.error("No se encontraron datos válidos para construir el heatmap de notificaciones.")
    st.stop()

years = get_available_years(metrics_df)
default_year = get_default_year(years)
default_month_name = get_default_month_name()
project_options = ["Todos"] + sorted(metrics_df["proyecto"].dropna().unique().tolist(), key=str.casefold)

render_page_heading(
    "Notificaciones",
    "Seguimiento acumulado mensual de incidencias por proyecto con separación entre estructura y acabados.",
)

filter_col1, filter_col2, filter_col3, filter_col4 = st.columns([2.2, 1, 1, 1.3])
with filter_col1:
    selected_project = st.selectbox("Proyecto", project_options, index=0)
with filter_col2:
    selected_year = st.selectbox("Año", years, index=years.index(default_year))
with filter_col3:
    month_options = ["Todos"] + [MONTHS_ES[month] for month in range(1, 13)]
    default_month_index = month_options.index(default_month_name) if default_month_name in month_options else 0
    selected_month = st.selectbox("Mes KPI", month_options, index=default_month_index)
with filter_col4:
    st.markdown(
        kpi_card("Última actualización", last_update_text.split(" a las ")[0], "Fuente: Incidencias.xlsx", "kp-slate"),
        unsafe_allow_html=True,
    )

st.markdown(
    '<div class="info-note">El filtro de proyecto resalta la fila seleccionada sin ocultar los demás proyectos. '
    "El filtro de año cambia el heatmap. El filtro de mes queda disponible para los KPIs de la siguiente iteración.</div>",
    unsafe_allow_html=True,
)

selected_year_df = metrics_df[metrics_df["year"] == selected_year].copy()
selected_month_number = next((idx for idx, name in MONTHS_ES.items() if name == selected_month), None)
reference_month_number = selected_month_number or get_reference_month_number(selected_year_df, selected_year)
selected_month_slice = (
    selected_year_df[selected_year_df["month"] == selected_month_number].copy()
    if selected_month_number is not None
    else selected_year_df[selected_year_df["month"] == reference_month_number].copy()
)

summary_scope = selected_month_slice.copy()
if selected_project != "Todos":
    summary_scope = summary_scope[summary_scope["proyecto"] == selected_project]

accion_positiva_total = int(summary_scope["accion_positiva"].sum()) if not summary_scope.empty else 0
receptiva_total = int(summary_scope["receptiva_sin_accion"].sum()) if not summary_scope.empty else 0
parcial_total = int(summary_scope["accion_parcial"].sum()) if not summary_scope.empty else 0
acumulado_total = int(summary_scope["acumulado"].sum()) if not summary_scope.empty else 0
datos_nuevos_total = int(summary_scope["datos_nuevos_mes"].sum()) if not summary_scope.empty else 0
porcentaje_total = (accion_positiva_total / acumulado_total) if acumulado_total else 0

section_header(
    "KPIs de referencia",
    f"Corte referencial del filtro de mes: {selected_month if selected_month != 'Todos' else MONTHS_ES[reference_month_number]} {selected_year}",
)
kpi_cols = st.columns(5)
with kpi_cols[0]:
    st.markdown(
        kpi_card("Acción positiva", f"{accion_positiva_total:,}".replace(",", "."), "Cerradas, anuladas y vencidas atendidas", "kp-green"),
        unsafe_allow_html=True,
    )
with kpi_cols[1]:
    st.markdown(
        kpi_card("Receptiva sin acción", f"{receptiva_total:,}".replace(",", "."), "Vencidas sin atención al corte", "kp-amber"),
        unsafe_allow_html=True,
    )
with kpi_cols[2]:
    st.markdown(
        kpi_card("Acción parcial", f"{parcial_total:,}".replace(",", "."), "Abiertas dentro del plazo o en el límite", "kp-blue"),
        unsafe_allow_html=True,
    )
with kpi_cols[3]:
    st.markdown(
        kpi_card("Acumulado", f"{acumulado_total:,}".replace(",", "."), "Suma de las tres tipologías", "kp-slate"),
        unsafe_allow_html=True,
    )
with kpi_cols[4]:
    st.markdown(
        kpi_card("% atendidas", format_pct(porcentaje_total), f"Datos nuevos del mes: {datos_nuevos_total:,}".replace(",", "."), "kp-green"),
        unsafe_allow_html=True,
    )

section_header(
    "Heatmap mensual",
    f"Distribución del porcentaje de atendidas por proyecto durante {selected_year}.",
)

heatmap_col1, heatmap_col2 = st.columns(2)
for column, category in zip((heatmap_col1, heatmap_col2), HEATMAP_CATEGORIES):
    category_df = selected_year_df[selected_year_df["categoria"] == category].copy()
    category_df = category_df[(category_df["acumulado"] > 0) | (category_df["datos_nuevos_mes"] > 0)].copy()
    with column:
        st.markdown('<div class="heatmap-shell">', unsafe_allow_html=True)
        if category_df.empty:
            st.markdown(
                f'<div class="mini-note">No hay datos disponibles para {category.lower()} en {selected_year}.</div>',
                unsafe_allow_html=True,
            )
        else:
            option, chart_height = build_heatmap_option(
                category_df=category_df,
                category_name=category,
                year=selected_year,
                selected_project=selected_project,
            )
            render_echarts(option, height=chart_height)
            render_heatmap_legend()
        st.markdown("</div>", unsafe_allow_html=True)

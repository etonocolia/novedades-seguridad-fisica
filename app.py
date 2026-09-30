"""Dashboard local para analizar paradas no autorizadas y novedades relacionadas."""

from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
import unicodedata
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = BASE_DIR / "Informes.xlsx"
STOP_SHEET = "ParadasNoAutorizadas"
CONTEXT_SHEETS = {
    "Novedades de conductores": "NovedadesDeConductores",
    "Novedades de precinto": "NovedadesDePrecinto",
    "Unidades de transporte": "UnidadesDeTransporte",
}
REQUIRED_SHEETS = [
    "NovedadesDeConductores",
    "NovedadesDePrecinto",
    "UnidadesDeTransporte",
    STOP_SHEET,
]


def normalize_header(value: object) -> str:
    """Normaliza encabezados para hacer búsquedas sin depender de mayúsculas o tildes."""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(character for character in text if not unicodedata.combining(character))
    return " ".join(text.upper().strip().split())


def get_column(
    frame: pd.DataFrame,
    name: str,
    *,
    required: bool = False,
) -> pd.Series:
    """Obtiene una columna por nombre normalizado o devuelve valores vacíos."""
    wanted = normalize_header(name)
    for column in frame.columns:
        if normalize_header(column) == wanted:
            return frame[column]
    if required:
        raise ValueError(f"Falta la columna requerida: {name}")
    return pd.Series(pd.NA, index=frame.index, dtype="object")


@st.cache_data(ttl=900, max_entries=3, show_spinner="Leyendo el libro de novedades...")
def read_workbook(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    """Lee las cuatro hojas de datos; paradas tiene encabezados en la fila 6 de Excel."""
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message=".*extension is not supported.*",
            category=UserWarning,
        )
        excel_file = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
        missing = sorted(set(REQUIRED_SHEETS) - set(excel_file.sheet_names))
        if missing:
            raise ValueError("No se encontraron estas hojas requeridas: " + ", ".join(missing))

        sheets = {
            sheet_name: pd.read_excel(
                excel_file,
                sheet_name=sheet_name,
                header=5 if sheet_name == STOP_SHEET else 0,
            ).dropna(how="all")
            for sheet_name in REQUIRED_SHEETS
        }
    return sheets


def clean_text(series: pd.Series) -> pd.Series:
    """Limpia espacios y convierte cadenas vacías en valores faltantes."""
    values = series.astype("string").str.strip()
    return values.mask(values.eq(""), pd.NA)


def prepare_stops(raw: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, int]]:
    """Normaliza paradas y conserva solo un ejemplar de filas exactamente duplicadas."""
    duplicate_rows = int(raw.duplicated().sum())
    source = raw.drop_duplicates().copy()

    stops = pd.DataFrame(index=source.index)
    stops["date"] = pd.to_datetime(
        get_column(source, "FECHA", required=True), errors="coerce"
    ).dt.normalize()
    stops["timestamp"] = pd.to_datetime(
        get_column(source, "REGISTRO DE TIEMPO"), errors="coerce"
    )
    stops["date"] = stops["date"].fillna(stops["timestamp"].dt.normalize())
    missing_date_rows = int(stops["date"].isna().sum())
    stops = stops.loc[stops["date"].notna()].copy()

    for output, source_name in {
        "tractor": "PLACA MEDIO PROPULSOR",
        "storage": "PLACA MEDIO ALMACENADOR",
    }.items():
        stops[output] = clean_text(get_column(source, source_name)).str.upper()

    stops["site"] = clean_text(get_column(source, "SITIO"))
    stops["latitude"] = pd.to_numeric(get_column(source, "LATITUD"), errors="coerce")
    stops["longitude"] = pd.to_numeric(get_column(source, "LONGITUD"), errors="coerce")
    stops["description"] = clean_text(get_column(source, "DESCRIPCION DEL EVENTO"))
    stops["actions"] = clean_text(get_column(source, "ACCIONES"))
    stops["justification"] = clean_text(get_column(source, "JUSTIFICACION"))
    stops["hour"] = stops["timestamp"].dt.hour
    stops["coordinates_valid"] = (
        stops["latitude"].between(-90, 90)
        & stops["longitude"].between(-180, 180)
    )
    stops = stops.reset_index(drop=True)

    quality = {
        "duplicate_rows": duplicate_rows,
        "missing_date_rows": missing_date_rows,
        "missing_time_rows": int(stops["timestamp"].isna().sum()),
        "invalid_coordinate_rows": int((~stops["coordinates_valid"]).sum()),
    }
    return stops, quality


def prepare_context(raw: pd.DataFrame) -> pd.DataFrame:
    """Extrae las fechas y placas disponibles para cruces de contexto exactos."""
    date_values = pd.to_datetime(get_column(raw, "FECHA"), errors="coerce").dt.normalize()
    context = pd.DataFrame({"date": date_values})
    context["tractor"] = clean_text(get_column(raw, "PLACA MEDIO PROPULSOR")).str.upper()
    context["storage"] = clean_text(get_column(raw, "PLACA MEDIO ALMACENADOR")).str.upper()
    return context.loc[context["date"].notna()].reset_index(drop=True)


def add_context_matches(
    stops: pd.DataFrame,
    contexts: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    """Marca coincidencias solo por fecha y placa exactas; no implica causalidad."""
    result = stops.copy()
    matched_sources: list[list[str]] = [[] for _ in range(len(result))]
    event_dates = result["date"].dt.date.tolist()

    for label, context in contexts.items():
        keys: set[tuple[date, str]] = set()
        for row in context.itertuples(index=False):
            context_date = row.date.date()
            for plate in (row.tractor, row.storage):
                if pd.notna(plate):
                    keys.add((context_date, str(plate)))

        for position, row in enumerate(result.itertuples(index=False)):
            if any(
                pd.notna(plate) and (event_dates[position], str(plate)) in keys
                for plate in (row.tractor, row.storage)
            ):
                matched_sources[position].append(label)

    result["context_sources"] = [", ".join(labels) for labels in matched_sources]
    result["context_match"] = result["context_sources"].ne("")
    return result


def add_alerts(frame: pd.DataFrame, repeat_threshold: int) -> pd.DataFrame:
    """Aplica reglas legibles sobre recurrencia y coincidencias de otras hojas."""
    result = frame.copy()
    tractor_counts = result["tractor"].value_counts()
    storage_counts = result["storage"].value_counts()
    result["tractor_count"] = result["tractor"].map(tractor_counts).fillna(0).astype(int)
    result["storage_count"] = result["storage"].map(storage_counts).fillna(0).astype(int)
    result["repeat_signal"] = (
        result["tractor_count"].ge(repeat_threshold)
        | result["storage_count"].ge(repeat_threshold)
    )
    result["alert_points"] = (
        result["repeat_signal"].astype(int) + result["context_match"].astype(int)
    )
    result["priority"] = np.select(
        [result["alert_points"].ge(2), result["alert_points"].eq(1)],
        ["Prioridad alta", "Revisar"],
        default="Sin señal",
    )
    return result


def apply_filters(
    frame: pd.DataFrame,
    date_range: tuple[date, date],
    tractor: str,
    storage: str,
    site_query: str,
) -> pd.DataFrame:
    """Aplica filtros operativos sin modificar el conjunto de origen."""
    start, end = date_range
    filtered = frame.loc[frame["date"].dt.date.between(start, end)].copy()
    if tractor != "Todas":
        filtered = filtered.loc[filtered["tractor"] == tractor]
    if storage != "Todas":
        filtered = filtered.loc[filtered["storage"] == storage]
    if site_query.strip():
        filtered = filtered.loc[
            filtered["site"].fillna("").str.contains(site_query.strip(), case=False, regex=False)
        ]
    return filtered


def show_chart(fig: object) -> None:
    """Renderiza un gráfico Plotly usando el ancho disponible."""
    st.plotly_chart(fig, width="stretch")


st.set_page_config(page_title="Paradas no autorizadas", page_icon=":material/local_shipping:", layout="wide")
st.title("Control de paradas no autorizadas")
st.caption("Análisis de recurrencia, exposición operativa y coincidencias con otras novedades")

with st.sidebar:
    st.header("Fuente de datos")
    uploaded_file = st.file_uploader("Cargar otro libro Excel", type=["xlsx"])

if uploaded_file is not None:
    workbook_bytes = uploaded_file.getvalue()
    source_label = uploaded_file.name
elif WORKBOOK_PATH.exists():
    workbook_bytes = WORKBOOK_PATH.read_bytes()
    source_label = WORKBOOK_PATH.name
else:
    st.error("No se encontró Informes.xlsx. Cárguelo desde el panel lateral.")
    st.stop()

try:
    raw_sheets = read_workbook(workbook_bytes)
    stops, quality = prepare_stops(raw_sheets[STOP_SHEET])
    contexts = {
        label: prepare_context(raw_sheets[sheet_name])
        for label, sheet_name in CONTEXT_SHEETS.items()
    }
    stops = add_context_matches(stops, contexts)
except (ValueError, KeyError, OSError, ImportError) as error:
    st.error(f"No se pudo procesar el libro: {error}")
    st.stop()

if stops.empty:
    st.warning("La hoja de paradas no contiene eventos con fecha válida.")
    st.stop()

minimum_date = stops["date"].min().date()
maximum_date = stops["date"].max().date()
with st.sidebar:
    st.header("Filtros operativos")
    selected_dates = st.date_input(
        "Periodo",
        value=(minimum_date, maximum_date),
        min_value=minimum_date,
        max_value=maximum_date,
    )
    if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
        date_range = (selected_dates[0], selected_dates[1])
    elif isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 1:
        date_range = (selected_dates[0], selected_dates[0])
    elif isinstance(selected_dates, (tuple, list)) or selected_dates is None:
        date_range = (minimum_date, maximum_date)
    else:
        date_range = (selected_dates, selected_dates)

    tractor_options = ["Todas"] + sorted(stops["tractor"].dropna().unique().tolist())
    storage_options = ["Todas"] + sorted(stops["storage"].dropna().unique().tolist())
    selected_tractor = st.selectbox("Placa medio propulsor", tractor_options)
    selected_storage = st.selectbox("Placa medio almacenador", storage_options)
    site_query = st.text_input("Buscar sitio")
    repeat_threshold = st.number_input(
        "Eventos para marcar recurrencia",
        min_value=2,
        max_value=50,
        value=5,
        step=1,
        help="Umbral configurable para priorizar placas repetidas en el periodo filtrado.",
    )

filtered = apply_filters(
    stops,
    date_range,
    selected_tractor,
    selected_storage,
    site_query,
)
filtered = add_alerts(filtered, int(repeat_threshold))

st.caption(
    f"Fuente: {source_label} · Periodo disponible: {minimum_date:%d/%m/%Y} a "
    f"{maximum_date:%d/%m/%Y} · Los datos se procesan en esta sesión local."
)

if filtered.empty:
    st.info("No hay eventos para los filtros seleccionados.")
    st.stop()

with st.container(horizontal=True):
    st.metric("Paradas registradas", f"{len(filtered):,}")
    st.metric("Placas propulsoras", f"{filtered['tractor'].nunique(dropna=True):,}")
    st.metric("Sitios distintos", f"{filtered['site'].nunique(dropna=True):,}")
    st.metric("Coincidencia con otra hoja", f"{int(filtered['context_match'].sum()):,}")

st.info(
    "El libro no contiene cantidad de carga recibida/despachada ni un total confiable de viajes. "
    "Por eso no se calcula una tasa por viaje ni se afirma que hubo pérdida."
)

management_tab, operations_tab, quality_tab = st.tabs(
    ["Gerencia", "Operación", "Calidad y método"]
)

with management_tab:
    daily = (
        filtered.groupby(filtered["date"].dt.date)
        .size()
        .rename_axis("Día")
        .reset_index(name="Paradas")
    )
    hourly = (
        filtered.dropna(subset=["hour"])
        .groupby("hour")
        .size()
        .reindex(range(24), fill_value=0)
        .rename_axis("Hora")
        .reset_index(name="Paradas")
    )
    asset_counts = (
        filtered["storage"]
        .value_counts()
        .head(10)
        .rename_axis("Placa")
        .reset_index(name="Eventos")
    )
    top_storage_share = (
        float(asset_counts.iloc[0]["Eventos"]) / len(filtered) * 100
        if not asset_counts.empty
        else 0.0
    )
    top_columns = st.columns(2)
    with top_columns[0]:
        st.metric("Participación de la placa almacenadora más frecuente", f"{top_storage_share:.1f}%")
        st.caption("Concentración de registros, no tasa de riesgo por viaje.")
    with top_columns[1]:
        st.metric("Eventos con alguna señal", f"{int(filtered['alert_points'].gt(0).sum()):,}")
        st.caption("Recurrencia configurable o coincidencia exacta por fecha y placa.")

    chart_columns = st.columns(2)
    with chart_columns[0]:
        st.subheader("Evolución diaria")
        if not daily.empty:
            show_chart(
                px.line(daily, x="Día", y="Paradas", markers=True, labels={"Día": "Fecha"})
            )
    with chart_columns[1]:
        st.subheader("Distribución por hora")
        show_chart(
            px.bar(hourly, x="Hora", y="Paradas", labels={"Hora": "Hora del día"})
        )

    chart_columns = st.columns(2)
    with chart_columns[0]:
        st.subheader("Sitios con más registros")
        top_sites = (
            filtered["site"].dropna().value_counts().head(10)
            .rename_axis("Sitio").reset_index(name="Eventos")
        )
        if not top_sites.empty:
            show_chart(
                px.bar(top_sites.sort_values("Eventos"), x="Eventos", y="Sitio", orientation="h")
            )
    with chart_columns[1]:
        st.subheader("Placas almacenadoras con más registros")
        if not asset_counts.empty:
            show_chart(
                px.bar(asset_counts.sort_values("Eventos"), x="Eventos", y="Placa", orientation="h")
            )

    context_summary = pd.DataFrame(
        {
            "Hoja de contexto": list(contexts),
            "Registros disponibles": [len(context) for context in contexts.values()],
        }
    )
    st.subheader("Hojas complementarias")
    st.dataframe(context_summary, hide_index=True)

with operations_tab:
    st.subheader("Eventos para revisión")
    st.caption(
        f"Reglas: recurrencia si una placa aparece al menos {int(repeat_threshold)} veces "
        "en el periodo filtrado; coincidencia contextual si hay misma fecha y placa en otra hoja. "
        "Prioridad alta significa que se activan ambas reglas, no que exista pérdida confirmada."
    )

    alert_events = filtered.loc[filtered["alert_points"].gt(0)].copy()
    if alert_events.empty:
        st.success("No hay eventos que activen las reglas con los filtros actuales.")
    else:
        event_columns = [
            "timestamp", "date", "tractor", "storage", "site", "priority",
            "context_sources", "description", "actions", "justification",
        ]
        event_labels = {
            "timestamp": "Registro de tiempo",
            "date": "Fecha",
            "tractor": "Placa propulsora",
            "storage": "Placa almacenadora",
            "site": "Sitio",
            "priority": "Prioridad de revisión",
            "context_sources": "Coincidencia en hoja",
            "description": "Descripción",
            "actions": "Acciones",
            "justification": "Justificación",
        }
        table = alert_events[event_columns].rename(columns=event_labels)
        st.dataframe(table, hide_index=True, height=460)
        st.download_button(
            "Descargar eventos para revisión",
            data=table.to_csv(index=False).encode("utf-8-sig"),
            file_name="eventos_para_revision.csv",
            mime="text/csv",
            icon=":material/download:",
        )

    top_sites = (
        filtered["site"].dropna().value_counts().head(15)
        .rename_axis("site").reset_index(name="Paradas")
    )
    valid_coordinates = filtered.loc[
        filtered["coordinates_valid"] & filtered["site"].notna()
    ]
    coordinate_counts = (
        valid_coordinates.groupby(["site", "latitude", "longitude"])
        .size()
        .rename("Registros en coordenada")
        .reset_index()
        .sort_values("Registros en coordenada", ascending=False)
        .drop_duplicates("site")
    )
    map_sites = top_sites.merge(
        coordinate_counts[["site", "latitude", "longitude"]],
        on="site",
        how="left",
    ).dropna(subset=["latitude", "longitude"])

    st.subheader("Mapa de los 15 sitios con más paradas")
    if map_sites.empty:
        st.warning("Los 15 sitios principales no tienen coordenadas válidas para mostrar.")
    else:
        st.caption(
            "Fondo de calles OpenStreetMap (requiere conexión a internet). El tamaño representa "
            "todas las paradas del sitio con los filtros actuales. Cada marcador usa la coordenada "
            "válida más frecuente; los sitios sin coordenadas válidas se omiten."
        )
        map_figure = px.scatter_map(
            map_sites,
            lat="latitude",
            lon="longitude",
            size="Paradas",
            hover_name="site",
            hover_data={"Paradas": True, "latitude": ":.5f", "longitude": ":.5f"},
            map_style="open-street-map",
            center={
                "lat": float(map_sites["latitude"].median()),
                "lon": float(map_sites["longitude"].median()),
            },
            zoom=9,
            size_max=40,
            labels={"Paradas": "Paradas no autorizadas"},
        )
        map_figure.update_traces(marker={"color": "#f05a28", "opacity": 0.88})
        map_figure.update_layout(height=500, margin={"l": 0, "r": 0, "t": 10, "b": 0})
        show_chart(map_figure)

with quality_tab:
    st.subheader("Controles de calidad")
    quality_columns = st.columns(4)
    quality_columns[0].metric("Duplicados exactos excluidos", f"{quality['duplicate_rows']:,}")
    quality_columns[1].metric("Fechas no interpretables", f"{quality['missing_date_rows']:,}")
    quality_columns[2].metric("Horas no interpretables", f"{quality['missing_time_rows']:,}")
    quality_columns[3].metric("Coordenadas fuera de rango", f"{quality['invalid_coordinate_rows']:,}")

    st.subheader("Indicadores y límites")
    st.markdown(
        "- **Volumen y tendencia:** conteo de eventos únicos por periodo; no hay total de viajes para calcular tasas.\n"
        "- **Recurrencia:** frecuencia por placa y sitio; una concentración alta requiere validar asignación y calidad del registro.\n"
        "- **Coincidencias:** cruce exacto por fecha y placa con conductores, precintos o unidades; no demuestra relación causal.\n"
        "- **Duración:** no disponible; el registro de tiempo es una marca horaria, no hora de inicio y fin.\n"
        "- **Carga:** no hay peso/cantidad de salida y llegada; una pérdida física no puede deducirse solo con estas novedades.\n"
        "- **Mapa:** excluye coordenadas fuera de latitud [-90, 90] y longitud [-180, 180]."
    )

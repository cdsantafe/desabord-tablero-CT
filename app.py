"""
Dashboard ACUMULADO - Notificación al cliente y adherencia al radio (BEES Deliver)
Lee el export mensual/acumulado (Excel o CSV, hoja "Export") y muestra más información operativa.
Ejecutar local:  streamlit run app.py
"""
from html import escape
from io import BytesIO
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # sin interfaz gráfica (servidor)
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

# ----------------------------------------------------------------------
# Configuración general
# ----------------------------------------------------------------------
OBJETIVO = 88.0                    # % objetivo de notificación > 2 min
OBJETIVO_RADIO_DEFECTO = 95        # % objetivo de adherencia al radio (editable en la barra lateral)
HORA_INICIO_GRAFICO = 7            # el gráfico por hora arranca a las 7:00
ARCHIVO_DEFECTO = "Septiembre.xlsx"
ESTADOS_VISITA = ["CONCLUDED", "DEFINITELY_RETURNED", "PARTIAL_DELIVERY"]   # visitas con resultado final

# Columnas del export -> nombres internos
MAPA = {
    "FECHA": "fecha_raw", "_ID_PLANILLA": "planilla", "CHOFER BEES": "chofer", "CHOFER TRUCK": "chofer_truck",
    "_ID_CLIENTE": "cod_cliente", "NOMBRE_CLIENTE": "cliente",
    "FECHA_INICIO_VIAJE": "inicio_viaje", "FECHA_FIN_VIAJE": "fin_viaje", "ESTATUS_TOUR": "estatus_tour",
    "FECHA_NOTIFICACION": "notif", "FECHA_INICIO_VISITA": "arribo", "FECHA_FIN_VISITA": "fin_visita",
    "COD_CATEGORIA": "categoria", "Cliente Click Notificacion": "click_cliente",
    "ESTADO_FINAL_VISITA": "estado", "Intentos_Entrega": "intentos", "COD_MOTIVO": "motivo",
    "COMENTARIO": "comentario", "FLAG_RADIO_INICIO": "radio_ini", "FLAG_RADIO_FIN": "radio_fin",
    "FLAG_RADIO_ENTREGA_FINAL": "radio_ent", "MOTIVO_RADIO": "motivo_radio",
    "Tipo Entrega Rechazo": "tipo_rechazo", "Segmento Venta": "segmento", "PATENTE": "patente",
}
OBLIGATORIAS = ["FECHA", "_ID_PLANILLA", "CHOFER BEES", "FECHA_NOTIFICACION", "FECHA_INICIO_VISITA"]

MOTIVOS_ES = {
    "Closed": "Cliente cerrado", "Without Money": "Sin dinero", "Underinvoiced": "Subfacturado",
    "Under Invoiced": "Subfacturado", "Error order": "Error de pedido", "Out of hours": "Fuera de horario",
    "Out of route/area": "Fuera de ruta/zona", "Product in poor condition": "Producto en mal estado",
    "Out of Stock Product": "Producto sin stock", "Incorrect address": "Dirección incorrecta",
    "Return product not available": "Producto de retorno no disponible", "Cannot continue route": "No puede continuar la ruta",
    "Business Problems": "Problemas del negocio", "Without Containers": "Sin envases", "No Order": "Sin pedido",
    "Unspecified reason": "Motivo no especificado", "Badly Billed": "Mal facturado",
    "Container not available": "Envase no disponible", "Wrong Product": "Producto equivocado",
}

st.set_page_config(page_title="Acumulado - Notificación y Radio", page_icon="📲", layout="wide")

# ----------------------------------------------------------------------
# CSS personalizado
# ----------------------------------------------------------------------
st.markdown(
    """
    <style>
    .titulo {color:#D32F2F; font-size:2.4rem; font-weight:800; margin-bottom:0; letter-spacing:.5px;}
    .subtitulo {color:#444; font-size:1.2rem; font-weight:600; margin-top:0;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# Tabla ordenable, imagen JPEG y exportación a Excel
# ----------------------------------------------------------------------
PLANTILLA_TABLA = """
<style>
  body {font-family: "Source Sans Pro", sans-serif; margin:0;}
  #wrap {background:#fff; padding:4px;}
  #wrap:fullscreen {padding:24px; overflow:auto;}
  .barra {display:flex; gap:8px; justify-content:flex-end; margin-bottom:8px;}
  .barra button {border:1px solid #d9a800; background:#FFF8E1; border-radius:6px; padding:6px 12px;
                 cursor:pointer; font-size:.9rem;}
  .barra button:hover {background:#FFC933;}
  table {width:auto; max-width:__ANCHO__px; border-collapse:collapse;}
  #wrap:fullscreen table {width:100%; max-width:none;}
  th {background:#FFC933; color:#222; padding:6px 8px; font-size:.95rem; text-align:center; border:1px solid #d9a800;
      cursor:pointer; user-select:none;}
  th:hover {background:#FFB300;}
  td {padding:5px 8px; text-align:center; border:1px solid #e0e0e0; font-size:1.15rem; font-weight:600;}
  td.nom {text-align:left;}
  td.ok  {background:#C8E6C9; color:#1B5E20;}
  td.bad {background:#FFCDD2; color:#B71C1C;}
  tr.gral td {font-size:1.8rem; font-weight:800; padding:10px 8px;
              border-top:4px solid #D32F2F; border-bottom:4px solid #D32F2F;}
  tr.gral td.nom {background:#FFB300; color:#222; letter-spacing:1px;}
  tr.gral td.cli {background:#FFE082; color:#222;}
  tr.gral td.ok  {background:#2E7D32; color:#fff;}
  tr.gral td.bad {background:#C62828; color:#fff;}
  #wrap:fullscreen td {font-size:1.2rem;} #wrap:fullscreen tr.gral td {font-size:2rem;}
</style>
<div id="wrap">
  <div class="barra">
    <button onclick="fsTabla()">⛶ Maximizar tabla</button>
    <button onclick="fsDash()">🖥️ Dashboard en pantalla completa</button>
  </div>
  <table id="t">__CABECERA__<tbody>__FILAS__</tbody></table>
</div>
<script>
  const dirs = {};
  function ordenar(i) {
    const tb = document.querySelector('#t tbody');
    const asc = dirs[i] === undefined ? false : !dirs[i];   // primer clic: mayor a menor
    dirs[i] = asc;
    const filas = Array.from(tb.rows);
    filas.sort((a, b) => {
      const x = a.cells[i], y = b.cells[i];
      const vx = x.dataset.v !== undefined ? parseFloat(x.dataset.v) : x.textContent.toLowerCase();
      const vy = y.dataset.v !== undefined ? parseFloat(y.dataset.v) : y.textContent.toLowerCase();
      return (vx > vy ? 1 : vx < vy ? -1 : 0) * (asc ? 1 : -1);
    });
    filas.forEach(f => tb.appendChild(f));
    document.querySelectorAll('#t th').forEach((th, k) => {
      th.textContent = th.dataset.t + (k === i ? (asc ? ' ▲' : ' ▼') : '');
    });
  }
  document.querySelectorAll('#t th').forEach((th, k) => {
    th.dataset.t = th.textContent; th.onclick = () => ordenar(k);
  });
  function fsTabla() {
    const w = document.getElementById('wrap');
    if (document.fullscreenElement) { document.exitFullscreen(); return; }
    w.requestFullscreen().catch(() => alert('Tu navegador bloqueó la pantalla completa.'));
  }
  function fsDash() {
    try {
      const d = window.parent.document;
      if (d.fullscreenElement) { d.exitFullscreen(); return; }
      d.documentElement.requestFullscreen().catch(() => alert('Usá F11 para pantalla completa.'));
    } catch (e) { alert('Usá F11 para pantalla completa.'); }
  }
</script>
"""


def tabla_componente(tabla: pd.DataFrame, cols_num: list, cols_pct: list,
                     objetivo: float = OBJETIVO, ancho: int = 560) -> tuple[str, int]:
    """Tabla HTML ordenable (clic en encabezados) con la fila GENERAL fija arriba.
    cols_num: columnas de conteo; cols_pct: columnas de porcentaje coloreadas según el objetivo."""

    def fila(r, clase_tr=""):
        celdas = [f'<td class="nom">{escape(str(r["CHOFER"]))}</td>']
        for c in cols_num:
            celdas.append(f'<td class="cli" data-v="{int(r[c])}">{int(r[c])}</td>')
        for c in cols_pct:
            if pd.isna(r[c]):                      # sin datos (ej.: chofer sin visitas realizadas)
                celdas.append('<td data-v="-1">–</td>')
                continue
            k = "ok" if r[c] >= objetivo else "bad"
            celdas.append(f'<td class="{k}" data-v="{r[c]:.4f}">{r[c]:.0f}%</td>')
        return f'<tr class="{clase_tr}">' + "".join(celdas) + "</tr>"

    gral = tabla[tabla["CHOFER"] == "GENERAL"].iloc[0]
    choferes = tabla[tabla["CHOFER"] != "GENERAL"]
    titulos = "".join(f"<th>{escape(c)}</th>" for c in ["CHOFER", *cols_num, *cols_pct])
    cabecera = f"<thead><tr>{titulos}</tr>{fila(gral, 'gral')}</thead>"
    html = (PLANTILLA_TABLA.replace("__ANCHO__", str(ancho)).replace("__CABECERA__", cabecera)
            .replace("__FILAS__", "".join(fila(r) for _, r in choferes.iterrows())))
    alto = min(190 + 38 * len(choferes), 800)
    return html, alto


def generar_jpeg(tabla: pd.DataFrame, etiqueta: str, titulo: str, subtitulo: str, titulos: list,
                 cols_num: list, cols_pct: list, anchos: list, objetivo: float = OBJETIVO) -> bytes:
    """Imagen JPEG pensada para verse en el celular sin abrirla: formato vertical, letra grande y celdas compactas.
    titulos: encabezados de columna (con \\n para partir líneas); anchos: ancho de cada columna (suma 100)."""
    gral = tabla[tabla["CHOFER"] == "GENERAL"].iloc[0]
    chof = tabla[tabla["CHOFER"] != "GENERAL"]
    n_lin_tit = titulo.count("\n") + 1

    # Medidas en pulgadas (ancho fijo 7": se ve completo en una pantalla de celular)
    ancho_tab, h_cab, h_gral, h_fila, base = 7.0, 0.95, 1.0, 0.52, 0.1
    n_lin_sub = subtitulo.count("\n") + 1
    y_obj = 0.12 + 0.38 * n_lin_tit               # posición del subtítulo (pulgadas desde arriba)
    y_fecha = y_obj + 0.225 * n_lin_sub           # posición de la línea FECHA
    tope = y_fecha + 0.22
    tabla_in = h_cab + h_gral + h_fila * len(chof)
    alto = base + tabla_in + tope
    # Lienzo más cuadrado: ancho ~3/4 del alto (máx. 12"), con espacio en blanco a los lados
    ancho = min(12.0, max(ancho_tab, alto * 0.75))
    mx = (ancho - ancho_tab) / 2 / ancho          # margen lateral (fracción del ancho)
    fig = plt.figure(figsize=(ancho, alto), dpi=220, facecolor="white")

    # Encabezado
    fig.text(mx + 0.14 / ancho, 1 - 0.12 / alto, titulo, color="#D32F2F", fontsize=21, fontweight="bold",
             va="top", linespacing=1.05)
    fig.text(mx + 0.14 / ancho, 1 - y_obj / alto, subtitulo, color="#333", fontsize=13, fontweight="bold",
             va="top", linespacing=1.1)
    fig.text(mx + 0.14 / ancho, 1 - y_fecha / alto, f"FECHA: {etiqueta}", color="#333", fontsize=13, va="top")

    # Tabla dibujada a mano, pegada a los bordes
    ax = fig.add_axes([mx + 0.07 / ancho, base / alto, (ancho_tab - 0.14) / ancho, tabla_in / alto])
    ax.set_xlim(0, 100); ax.set_ylim(tabla_in, 0); ax.axis("off")
    xs = np.cumsum([0] + list(anchos))

    def celda(i, y, h, texto, fondo, color="#222", size=15, peso="normal", borde="#bbbbbb", lw=0.6, ha="center"):
        ax.add_patch(Rectangle((xs[i], y), anchos[i], h, facecolor=fondo, edgecolor=borde, linewidth=lw))
        x_txt = xs[i] + (1.5 if ha == "left" else anchos[i] / 2)
        ax.text(x_txt, y + h / 2, texto, ha=ha, va="center", fontsize=size, color=color, fontweight=peso,
                linespacing=1.05)

    for i, t in enumerate(titulos):
        celda(i, 0, h_cab, t, "#FFC933", size=13, peso="bold", borde="#d9a800", lw=0.8)

    n_num = len(cols_num)
    y = h_cab                                             # fila GENERAL
    celda(0, y, h_gral, "GENERAL", "#FFB300", size=26, peso="bold", ha="left")
    for j, c in enumerate(cols_num):
        celda(1 + j, y, h_gral, f"{int(gral[c])}", "#FFE082", size=26 if gral[c] < 1000 else 20, peso="bold")
    for j, c in enumerate(cols_pct):
        ok = gral[c] >= objetivo
        celda(1 + n_num + j, y, h_gral, f"{gral[c]:.0f}%", "#2E7D32" if ok else "#C62828", "white", 30, "bold")
    ax.add_patch(Rectangle((0, y), 100, h_gral, fill=False, edgecolor="#D32F2F", linewidth=3))

    y += h_gral                                           # choferes
    for _, r in chof.iterrows():
        nombre = str(r["CHOFER"])
        celda(0, y, h_fila, nombre, "white", size=14 if len(nombre) <= 17 else 12, ha="left")
        for j, c in enumerate(cols_num):
            celda(1 + j, y, h_fila, f"{int(r[c])}", "white", size=15)
        for j, c in enumerate(cols_pct):
            if pd.isna(r[c]):                       # sin dato
                celda(1 + n_num + j, y, h_fila, "–", "white", size=17)
                continue
            bien = r[c] >= objetivo
            celda(1 + n_num + j, y, h_fila, f"{r[c]:.0f}%", "#C8E6C9" if bien else "#FFCDD2",
                  "#1B5E20" if bien else "#B71C1C", size=17, peso="bold")
        y += h_fila

    buf = BytesIO()
    fig.savefig(buf, format="jpeg", pil_kwargs={"quality": 92})
    plt.close(fig)
    return buf.getvalue()


def a_excel(tabla: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tabla.to_excel(w, index=False, sheet_name="Notificación")
    return buf.getvalue()



# ----------------------------------------------------------------------
# Carga y preparación de datos
# ----------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo archivo…")
def leer_archivo(nombre: str, contenido: bytes) -> pd.DataFrame:
    """Lee un Excel (hoja Export) o CSV a partir de sus bytes."""
    buf = BytesIO(contenido)
    if nombre.lower().endswith(".csv"):
        df = pd.read_csv(buf, sep=None, engine="python")
    else:
        xl = pd.ExcelFile(buf)
        df = xl.parse("Export" if "Export" in xl.sheet_names else xl.sheet_names[0])
    df["archivo_origen"] = nombre
    return df


def preparar(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia el export y calcula anticipación, banderas de cumplimiento y de radio."""
    faltan = [c for c in OBLIGATORIAS if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas en el archivo: {', '.join(faltan)}")
    df = df.rename(columns=MAPA)
    for c in MAPA.values():                       # columnas opcionales ausentes -> vacías
        if c not in df.columns:
            df[c] = np.nan

    # Filas de pie de página / vacías del export (sin planilla)
    df = df[pd.to_numeric(df["planilla"], errors="coerce").notna()].copy()
    df["planilla"] = pd.to_numeric(df["planilla"]).astype("int64").astype(str)
    cod = pd.to_numeric(df["cod_cliente"], errors="coerce").astype("Int64").astype(str)
    df["cod_cliente"] = cod.replace("<NA>", "")
    df["fecha"] = pd.to_datetime(df["fecha_raw"], errors="coerce").dt.normalize()
    df = df[df["fecha"].notna()].copy()

    for c in ("inicio_viaje", "fin_viaje", "notif", "arribo", "fin_visita"):
        df[c] = pd.to_datetime(df[c], errors="coerce")   # horas del export: ya en hora local
    df["chofer"] = df["chofer"].fillna("Sin chofer").astype(str).str.strip().str.title()
    df["segmento"] = df["segmento"].fillna("Sin segmento")
    df["estado"] = df["estado"].astype(str).str.strip().str.upper()

    # Notificación: arribo - notificación (sin notificación => no cumple)
    df["anticipacion_seg"] = (df["arribo"] - df["notif"]).dt.total_seconds()
    df["cumple_2"] = df["anticipacion_seg"] >= 120
    df["cumple_5"] = df["anticipacion_seg"] >= 300
    df["hora"] = df["arribo"].dt.hour
    df["t_visita_min"] = (df["fin_visita"] - df["arribo"]).dt.total_seconds() / 60

    # Radio: 1 = dentro, 0 = fuera, vacío = sin dato
    df["en_radio_ini"] = np.where(df["radio_ini"].isna(), np.nan, (df["radio_ini"] == 1).astype(float))
    df["en_radio_fin"] = np.where(df["radio_fin"].isna(), np.nan, (df["radio_fin"] == 1).astype(float))

    # Resultado de la visita
    df["devuelta"] = df["estado"] == "DEFINITELY_RETURNED"
    df["motivo_es"] = df["motivo"].map(MOTIVOS_ES).fillna(df["motivo"])
    df["rechazo"] = df["tipo_rechazo"].notna()
    df["click"] = df["click_cliente"].astype(str).str.strip().str.lower().isin(["si", "sí"])
    return df.reset_index(drop=True)


@st.cache_data(show_spinner="Procesando datos…")
def cargar(archivos: tuple) -> tuple[pd.DataFrame, int, pd.DataFrame]:
    """Une uno o varios archivos. Si una visita está en más de un archivo, queda la del último."""
    partes = [leer_archivo(n, b) for n, b in archivos]
    resumen = pd.DataFrame({"Archivo": [n for n, _ in archivos], "Filas": [len(p) for p in partes]})
    datos = preparar(pd.concat(partes, ignore_index=True))
    antes = len(datos)
    datos = datos.drop_duplicates(subset=["planilla", "cod_cliente"], keep="last").reset_index(drop=True)
    return datos, antes - len(datos), resumen


def asignar_periodo(df: pd.DataFrame, frecuencia: str) -> pd.DataFrame:
    df = df.copy()
    if frecuencia == "Diario":
        df["periodo"] = df["fecha"].dt.strftime("%Y-%m-%d")
    elif frecuencia == "Semanal":
        iso = df["fecha"].dt.isocalendar()
        df["periodo"] = iso["year"].astype(str) + "-S" + iso["week"].astype(str).str.zfill(2)
    else:
        df["periodo"] = df["fecha"].dt.strftime("%Y-%m")
    return df


def mmss(minutos) -> str:
    """Minutos (decimal) -> 'mm:ss'."""
    if pd.isna(minutos):
        return "–"
    seg = int(round(float(minutos) * 60))
    return f"{seg // 60:02d}:{seg % 60:02d}"


def pct(serie: pd.Series) -> float:
    return float(serie.mean() * 100) if len(serie) else np.nan


# ----------------------------------------------------------------------
# Encabezado y carga de archivos
# ----------------------------------------------------------------------
st.markdown('<p class="titulo">TAREA CRÍTICA: NOTIFICACIÓN AL CLIENTE</p>', unsafe_allow_html=True)
st.markdown(f'<p class="subtitulo">Objetivo: {OBJETIVO:.0f}% de los clientes notificados con más de 2 minutos '
            'de anticipación · Vista acumulada</p>', unsafe_allow_html=True)

tipos = ["xlsx", "xls", "csv"]
sub_side = st.sidebar.file_uploader("📂 Cargar archivos (uno o varios)", type=tipos,
                                    accept_multiple_files=True, key="up_side")
sub_main = st.file_uploader("📂 Cargar archivos acumulados (Excel o CSV, podés subir varios meses)",
                            type=tipos, accept_multiple_files=True, key="up_main")

subidos, vistos = [], set()
for a in list(sub_main or []) + list(sub_side or []):
    if (a.name, a.size) not in vistos:
        vistos.add((a.name, a.size))
        subidos.append((a.name, a.getvalue()))

if subidos:
    archivos = tuple(subidos)
elif Path(ARCHIVO_DEFECTO).exists():
    archivos = ((ARCHIVO_DEFECTO, Path(ARCHIVO_DEFECTO).read_bytes()),)
else:
    st.info("Cargá un archivo de Excel o CSV (export acumulado) para comenzar.")
    st.stop()

try:
    datos, n_dup, resumen_archivos = cargar(archivos)
except Exception as e:
    st.error(f"No se pudo procesar el archivo: {e}")
    st.stop()

texto_origen = f"{len(resumen_archivos)} archivo(s) · {len(datos):,} visitas".replace(",", ".")
if n_dup:
    texto_origen += f" · {n_dup} repetidas descartadas"
st.caption(texto_origen)
with st.expander("Archivos incluidos"):
    st.dataframe(resumen_archivos, use_container_width=True, hide_index=True)

# ----------------------------------------------------------------------
# Filtros (sidebar)
# ----------------------------------------------------------------------
st.sidebar.header("Filtros")
frecuencia = st.sidebar.radio("Frecuencia", ["Diario", "Semanal", "Mensual"], horizontal=True)

f_min, f_max = datos["fecha"].min().date(), datos["fecha"].max().date()
rango = st.sidebar.date_input("Fecha / rango", value=(f_min, f_max), min_value=f_min, max_value=f_max)
if isinstance(rango, tuple) and len(rango) == 2:
    desde, hasta = rango
else:
    desde = hasta = rango[0] if isinstance(rango, tuple) else rango

choferes = sorted(datos["chofer"].unique())
sel_ch = st.sidebar.multiselect("Chofer", choferes, default=choferes)
segmentos = sorted(datos["segmento"].unique())
sel_seg = st.sidebar.multiselect("Segmento de venta", segmentos, default=segmentos)
planillas = sorted(datos["planilla"].unique())
sel_pl = st.sidebar.multiselect("Planilla (vacío = todas)", planillas, default=[])
objetivo_radio = st.sidebar.number_input("Objetivo % adherencia al radio", 0, 100, OBJETIVO_RADIO_DEFECTO, 1)

mask = (
    (datos["fecha"].dt.date >= desde) & (datos["fecha"].dt.date <= hasta)
    & datos["chofer"].isin(sel_ch) & datos["segmento"].isin(sel_seg)
)
if sel_pl:
    mask &= datos["planilla"].isin(sel_pl)
f = datos[mask].copy()                          # todas las visitas filtradas
d = f[f["arribo"].notna()].copy()               # visitas con hora de arribo (se evalúan)
if d.empty:
    st.info("Sin visitas con hora de arribo para los filtros seleccionados.")
    st.stop()
d = asignar_periodo(d, frecuencia)

etiqueta = f"{desde:%d/%m/%Y}" if desde == hasta else f"{desde:%d/%m/%Y} - {hasta:%d/%m/%Y}"
g2, g5 = pct(d["cumple_2"]), pct(d["cumple_5"])

# Tabla de notificación por chofer (GENERAL arriba)
C2, C5 = "Notificaciones mayores a 2 minutos", "Notificaciones mayores a 5 minutos"
tabla = (d.groupby("chofer").agg(Clientes=("arribo", "size"), c2=("cumple_2", "mean"), c5=("cumple_5", "mean"))
         .reset_index().rename(columns={"chofer": "CHOFER", "c2": C2, "c5": C5}))
tabla[[C2, C5]] *= 100
tabla = tabla.sort_values(C2, ascending=False).reset_index(drop=True)
tabla = pd.concat([pd.DataFrame([{"CHOFER": "GENERAL", "Clientes": len(d), C2: g2, C5: g5}]), tabla],
                  ignore_index=True)

(tab_res, tab1, tab2, tab3, tab4, tab5, tab7, tab6) = st.tabs([
    "🏠 Resumen", "📋 Por chofer", "📈 Evolución", "⚠️ Visitas < 2 min",
    "📍 Adherencia al radio", "🚚 Operación", "⏱️ Tiempo de visita", "🔁 Devoluciones"])

# ----------------------------------------------------------------------
# Resumen
# ----------------------------------------------------------------------
with tab_res:
    st.markdown(f"**FECHA:** {etiqueta}")
    base_res = f[f["estado"].isin(ESTADOS_VISITA)]
    r_res = d[d["inicio_viaje"].notna() & d["en_radio_ini"].notna()]
    c = st.columns(4)
    c[0].metric("Visitas con arribo", f"{len(d):,}".replace(",", "."))
    c[1].metric("Planillas", f"{f['planilla'].nunique():,}".replace(",", "."))
    c[2].metric("Choferes", d["chofer"].nunique())
    c[3].metric("Días", d["fecha"].nunique())
    c = st.columns(4)
    c[0].metric("% notificado > 2 min", f"{g2:.1f}%", f"{g2 - OBJETIVO:+.1f} pts vs objetivo")
    c[1].metric("% notificado > 5 min", f"{g5:.1f}%")
    rad = pct(r_res["en_radio_ini"]) if len(r_res) else np.nan
    c[2].metric("% adherencia al radio", "–" if pd.isna(rad) else f"{rad:.1f}%",
                None if pd.isna(rad) else f"{rad - objetivo_radio:+.1f} pts vs objetivo")
    c[3].metric("% devoluciones", f"{pct(base_res['devuelta']):.1f}%" if len(base_res) else "–")
    c = st.columns(4)
    c[0].metric("Clientes por planilla (prom.)", f"{len(d) / max(f['planilla'].nunique(), 1):.1f}")
    rutas = f.drop_duplicates("planilla")
    h_ruta = ((rutas["fin_viaje"] - rutas["inicio_viaje"]).dt.total_seconds() / 3600)
    h_ruta = h_ruta[(h_ruta > 0) & (h_ruta < 16)]
    c[1].metric("Tiempo de ruta prom. (h)", f"{h_ruta.mean():.1f}" if len(h_ruta) else "–")
    c[2].metric("Rechazos", int(f["rechazo"].sum()))
    c[3].metric("Clics del cliente en la notificación", int(f["click"].sum()))

    ranking = tabla[(tabla["CHOFER"] != "GENERAL") & (tabla["Clientes"] >= 10)]
    col_a, col_b = st.columns(2)
    fmt = {"Clientes": st.column_config.NumberColumn(format="%d"),
           C2: st.column_config.NumberColumn("% > 2 min", format="%.1f%%"),
           C5: st.column_config.NumberColumn("% > 5 min", format="%.1f%%")}
    col_a.markdown("**Mejores 5 (≥ 10 visitas)**")
    col_a.dataframe(ranking.head(5), use_container_width=True, hide_index=True, column_config=fmt)
    col_b.markdown("**A mejorar (5 más bajos)**")
    col_b.dataframe(ranking.tail(5).iloc[::-1], use_container_width=True, hide_index=True, column_config=fmt)

# ----------------------------------------------------------------------
# Por chofer (notificación)
# ----------------------------------------------------------------------
with tab1:
    st.markdown(f"**FECHA:** {etiqueta}")
    html_tabla, alto_tabla = tabla_componente(tabla, ["Clientes"], [C2, C5])
    components.html(html_tabla, height=alto_tabla, scrolling=True)
    st.caption("Hacé clic en un encabezado para ordenar (▲▼). La fila GENERAL queda fija arriba.")
    col_x, col_j, _ = st.columns([1, 1, 3])
    col_x.download_button("⬇️ Descargar Excel", a_excel(tabla.round(1)), "notificacion_por_chofer.xlsx",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="xl_notif")
    try:
        col_j.download_button(
            "🖼️ Descargar imagen (JPEG)",
            generar_jpeg(tabla, etiqueta, "TAREA CRÍTICA:\nNOTIFICACIÓN AL CLIENTE",
                         f"Objetivo: {OBJETIVO:.0f}% de los clientes notificados con más\nde 2 minutos de anticipación",
                         ["CHOFER", "Clientes", "Notificaciones\nmayores a\n2 minutos",
                          "Notificaciones\nmayores a\n5 minutos"],
                         ["Clientes"], [C2, C5], [37, 13, 25, 25]),
            "dashboard_notificacion.jpeg", "image/jpeg", key="jpeg_notif")
    except Exception as e:
        col_j.warning(f"No se pudo generar la imagen: {e}")

# ----------------------------------------------------------------------
# Evolución
# ----------------------------------------------------------------------
with tab2:
    evo = (d.groupby("periodo")
             .agg(**{"% > 2 min": ("cumple_2", "mean"), "% > 5 min": ("cumple_5", "mean")})
             .mul(100).reset_index())
    fig = px.line(evo.melt("periodo", var_name="Métrica", value_name="%"),
                  x="periodo", y="%", color="Métrica", markers=True,
                  color_discrete_sequence=["#D32F2F", "#F9A825"])
    fig.add_hline(y=OBJETIVO, line_dash="dash", line_color="green", annotation_text=f"Objetivo {OBJETIVO:.0f}%")
    fig.update_layout(yaxis_range=[0, 105], xaxis_title=frecuencia, yaxis_title="% cumplimiento")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader(f"Evolución según la hora del día (desde las {HORA_INICIO_GRAFICO}:00)")
    hor = (d[d["hora"] >= HORA_INICIO_GRAFICO].groupby("hora")
             .agg(Visitas=("arribo", "size"), **{"% > 2 min": ("cumple_2", "mean"), "% > 5 min": ("cumple_5", "mean")})
             .reset_index())
    if hor.empty:
        st.info(f"No hay visitas desde las {HORA_INICIO_GRAFICO}:00 para los filtros elegidos.")
    else:
        hor[["% > 2 min", "% > 5 min"]] *= 100
        hor["Franja"] = hor["hora"].astype(str).str.zfill(2) + ":00"
        fig_h = px.line(hor.melt(["hora", "Franja", "Visitas"], var_name="Métrica", value_name="%"),
                        x="Franja", y="%", color="Métrica", markers=True, hover_data=["Visitas"],
                        color_discrete_sequence=["#D32F2F", "#F9A825"])
        fig_h.add_hline(y=OBJETIVO, line_dash="dash", line_color="green", annotation_text=f"Objetivo {OBJETIVO:.0f}%")
        fig_h.update_layout(yaxis_range=[0, 105], xaxis_title="Hora de llegada al cliente",
                            yaxis_title="% cumplimiento")
        st.plotly_chart(fig_h, use_container_width=True)
        fig_v = px.bar(hor, x="Franja", y="Visitas", color_discrete_sequence=["#FFC933"])
        fig_v.update_layout(height=220, xaxis_title=None, yaxis_title="Visitas evaluadas")
        st.plotly_chart(fig_v, use_container_width=True)

    st.subheader("Mapa de calor: cumplimiento por chofer y día")
    metrica = st.radio("Indicador", ["% > 2 min", "% > 5 min"], horizontal=True, key="mc_metrica")
    col_m = "cumple_2" if metrica == "% > 2 min" else "cumple_5"
    piv = d.pivot_table(index="chofer", columns="fecha", values=col_m, aggfunc="mean") * 100
    piv.columns = [c.strftime("%d/%m") for c in piv.columns]
    fig_m = px.imshow(piv, color_continuous_scale=["#C62828", "#FFF176", "#2E7D32"], zmin=50, zmax=100,
                      aspect="auto", text_auto=".0f", labels={"color": "%"})
    fig_m.update_layout(height=140 + 24 * len(piv), xaxis_title=None, yaxis_title=None)
    st.plotly_chart(fig_m, use_container_width=True)

# ----------------------------------------------------------------------
# Visitas que no cumplieron 2 minutos
# ----------------------------------------------------------------------
with tab3:
    inc = d[~d["cumple_2"]].copy()
    vista = pd.DataFrame({
        "Fecha": inc["fecha"].dt.strftime("%Y-%m-%d"), "Chofer": inc["chofer"], "Planilla": inc["planilla"],
        "Código cliente": inc["cod_cliente"], "Cliente": inc["cliente"],
        "Notificación": inc["notif"].dt.strftime("%H:%M:%S"), "Arribo": inc["arribo"].dt.strftime("%H:%M:%S"),
        "Anticipación (seg)": inc["anticipacion_seg"].round(0),
    }).sort_values("Anticipación (seg)", na_position="first")
    st.caption(f"{len(vista)} visitas sin cumplir (negativo = notificó después de llegar; vacío = sin notificación).")
    st.dataframe(vista, use_container_width=True, hide_index=True)

# ----------------------------------------------------------------------
# Adherencia al radio
# ----------------------------------------------------------------------
with tab4:
    st.subheader("Adherencia al radio (% de click dentro del radio al iniciar la visita)")
    r = d[d["inicio_viaje"].notna() & d["en_radio_ini"].notna()].copy()   # ruta iniciada + arribo + dato de radio
    if r.empty:
        st.info("No hay visitas con ruta iniciada, arribo y dato de radio para los filtros elegidos.")
    else:
        c_cli, c_fuera, c_ini, c_fin = "Clientes", "Fuera del radio", "% en radio al inicio", "% en radio al fin"
        tr = (r.groupby("chofer")
                .agg(**{c_cli: ("arribo", "size"), c_fuera: ("en_radio_ini", lambda x: int((x == 0).sum())),
                        c_ini: ("en_radio_ini", "mean"), c_fin: ("en_radio_fin", "mean")})
                .reset_index().rename(columns={"chofer": "CHOFER"}))
        tr[[c_ini, c_fin]] *= 100
        tr = tr.sort_values(c_ini, ascending=False).reset_index(drop=True)
        gr = pd.DataFrame([{"CHOFER": "GENERAL", c_cli: len(r), c_fuera: int((r["en_radio_ini"] == 0).sum()),
                            c_ini: r["en_radio_ini"].mean() * 100, c_fin: r["en_radio_fin"].mean() * 100}])
        tr = pd.concat([gr, tr], ignore_index=True)

        st.markdown(f"**FECHA:** {etiqueta}  ·  Objetivo: {objetivo_radio}%")
        html_r, alto_r = tabla_componente(tr, [c_cli, c_fuera], [c_ini, c_fin], objetivo_radio, ancho=620)
        components.html(html_r, height=alto_r, scrolling=True)
        st.caption("Solo planillas con ruta iniciada y entregas con hora de arribo. "
                   "Hacé clic en un encabezado para ordenar.")
        bx, bj, _ = st.columns([1, 1, 3])
        bx.download_button("⬇️ Descargar Excel", a_excel(tr.round(1)), "adherencia_radio.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="xl_radio")
        try:
            bj.download_button(
                "🖼️ Descargar imagen (JPEG)",
                generar_jpeg(tr, etiqueta, "ADHERENCIA AL RADIO",
                             f"Objetivo: {objetivo_radio}% de los clientes con\nclick dentro del radio",
                             ["CHOFER", "Clientes", "Fuera del\nradio", "En radio\nal inicio", "En radio\nal fin"],
                             [c_cli, c_fuera], [c_ini, c_fin], [32, 14, 16, 19, 19], objetivo_radio),
                "adherencia_radio.jpeg", "image/jpeg", key="jpeg_radio")
        except Exception as e:
            bj.warning(f"No se pudo generar la imagen: {e}")

        fuera = r[r["en_radio_ini"] == 0]
        if not fuera.empty:
            st.markdown("**Visitas fuera del radio**")
            det = pd.DataFrame({
                "Fecha": fuera["fecha"].dt.strftime("%Y-%m-%d"), "Chofer": fuera["chofer"],
                "Planilla": fuera["planilla"], "Código cliente": fuera["cod_cliente"], "Cliente": fuera["cliente"],
                "Motivo fuera de radio": fuera["motivo_radio"],
            })
            st.dataframe(det, use_container_width=True, hide_index=True)
            mot = fuera["motivo_radio"].fillna("Sin motivo informado").value_counts().reset_index()
            mot.columns = ["Motivo", "Visitas"]
            st.markdown("**Motivos informados por el chofer**")
            st.dataframe(mot, use_container_width=True, hide_index=True)

# ----------------------------------------------------------------------
# Operación
# ----------------------------------------------------------------------
with tab5:
    st.subheader("Operación por chofer")
    base = f[f["estado"].isin(ESTADOS_VISITA)]
    rutas = f.drop_duplicates("planilla").copy()
    rutas["h_ruta"] = (rutas["fin_viaje"] - rutas["inicio_viaje"]).dt.total_seconds() / 3600
    rutas = rutas[(rutas["h_ruta"] > 0) & (rutas["h_ruta"] < 16)]
    op = pd.DataFrame({
        "Planillas": f.groupby("chofer")["planilla"].nunique(),
        "Visitas con arribo": d.groupby("chofer").size(),
        "Tiempo de ruta prom. (h)": rutas.groupby("chofer")["h_ruta"].mean(),
        "Clientes por planilla": d.groupby("chofer").size() / f.groupby("chofer")["planilla"].nunique(),
        "% notificadas al llegar (< 10 seg)": d.groupby("chofer")["categoria"]
                                               .apply(lambda s: (s == "<10 seconds").mean() * 100),
        "Devoluciones": base.groupby("chofer")["devuelta"].sum(),
        "% devoluciones": base.groupby("chofer")["devuelta"].mean() * 100,
        "Entregas parciales": base[base["estado"] == "PARTIAL_DELIVERY"].groupby("chofer").size(),
        "Rechazos": f[f["rechazo"]].groupby("chofer").size(),
    })
    for c_ in ("Devoluciones", "Entregas parciales", "Rechazos", "% devoluciones"):
        op[c_] = op[c_].fillna(0)
    op = op.reset_index().rename(columns={"chofer": "CHOFER"}).sort_values("Devoluciones", ascending=False)
    cfg = {c_: st.column_config.NumberColumn(format="%d") for c_ in
           ("Planillas", "Visitas con arribo", "Devoluciones", "Entregas parciales", "Rechazos")}
    cfg.update({"Tiempo de ruta prom. (h)": st.column_config.NumberColumn(format="%.1f"),
                "Clientes por planilla": st.column_config.NumberColumn(format="%.1f"),
                "% notificadas al llegar (< 10 seg)": st.column_config.NumberColumn(format="%.1f%%"),
                "% devoluciones": st.column_config.NumberColumn(format="%.1f%%")})
    st.dataframe(op, use_container_width=True, hide_index=True, column_config=cfg)
    st.caption("% notificadas al llegar: la notificación se envió con menos de 10 segundos de anticipación "
               "(prácticamente al arribar). % devoluciones sobre visitas con resultado final.")
    st.download_button("⬇️ Descargar Excel", a_excel(op.round(1)), "operacion_por_chofer.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="xl_op")

    st.subheader("Por segmento de venta")
    seg = pd.DataFrame({
        "Visitas con arribo": d.groupby("segmento").size(),
        "% > 2 min": d.groupby("segmento")["cumple_2"].mean() * 100,
        "% > 5 min": d.groupby("segmento")["cumple_5"].mean() * 100,
        "% en radio al inicio": r.groupby("segmento")["en_radio_ini"].mean() * 100 if not r.empty else np.nan,
        "% devoluciones": base.groupby("segmento")["devuelta"].mean() * 100,
    }).reset_index().rename(columns={"segmento": "Segmento"})
    st.dataframe(seg, use_container_width=True, hide_index=True, column_config={
        "Visitas con arribo": st.column_config.NumberColumn(format="%d"),
        **{k: st.column_config.NumberColumn(format="%.1f%%") for k in
           ("% > 2 min", "% > 5 min", "% en radio al inicio", "% devoluciones")}})

    st.subheader("Anticipación de la notificación (categorías del export)")
    cat = d["categoria"].replace({
        "<10 seconds": "< 10 seg (al llegar)", "<2 minutes": "10 seg – 2 min", "<5 minutes": "2 – 5 min",
        "<10 minutes": "5 – 10 min", ">=10 minutes": "≥ 10 min"})
    orden = ["< 10 seg (al llegar)", "10 seg – 2 min", "2 – 5 min", "5 – 10 min", "≥ 10 min"]
    dist = cat.value_counts().reindex(orden).fillna(0).reset_index()
    dist.columns = ["Anticipación", "Visitas"]
    dist["%"] = (dist["Visitas"] / dist["Visitas"].sum() * 100).round(1)
    fig_d = px.bar(dist, x="Anticipación", y="Visitas", color_discrete_sequence=["#FFC933"], text="%")
    fig_d.update_layout(height=320, xaxis_title=None)
    st.plotly_chart(fig_d, use_container_width=True)

# ----------------------------------------------------------------------
# Tiempo de visita (seguimiento por chofer y por cliente)
# ----------------------------------------------------------------------
with tab7:
    st.subheader("Tiempo de visita")
    st.caption("Tiempo entre el inicio y el fin de la visita registrados en la app (FECHA_INICIO_VISITA → "
               "FECHA_FIN_VISITA). Hoy es un indicador de seguimiento: la mayoría de las visitas dura segundos, "
               "así que a futuro conviene medirlo con un evento más cercano a la entrega real.")
    tv = d[d["fin_visita"].notna() & (d["t_visita_min"] >= 0)].copy()
    tope = st.number_input("Excluir visitas de más de (min), posibles errores de registro",
                           1, 1440, 180, 10, key="tv_tope")
    n_exc = int((tv["t_visita_min"] > tope).sum())
    tv = tv[tv["t_visita_min"] <= tope]
    if tv.empty:
        st.info("No hay visitas con inicio y fin registrados para los filtros elegidos.")
    else:
        c = st.columns(5)
        c[0].metric("Visitas medidas", f"{len(tv):,}".replace(",", "."))
        c[1].metric("Mediana (mm:ss)", mmss(tv["t_visita_min"].median()))
        c[2].metric("Promedio (min)", f"{tv['t_visita_min'].mean():.1f}")
        c[3].metric("Percentil 90 (min)", f"{tv['t_visita_min'].quantile(.9):.1f}")
        c[4].metric("Excluidas por el tope", n_exc)

        vista = st.radio("Ver por", ["Chofer", "Cliente"], horizontal=True, key="tv_vista")
        cfg_t = {"Visitas": st.column_config.NumberColumn(format="%d"),
                 "Mediana (min)": st.column_config.NumberColumn(format="%.2f"),
                 "Promedio (min)": st.column_config.NumberColumn(format="%.2f"),
                 "P90 (min)": st.column_config.NumberColumn(format="%.2f"),
                 "Máx. (min)": st.column_config.NumberColumn(format="%.1f"),
                 "% < 1 min": st.column_config.NumberColumn(format="%.1f%%"),
                 "% ≥ 10 min": st.column_config.NumberColumn(format="%.1f%%")}

        def resumen_tiempos(df_, claves):
            g = df_.groupby(claves)["t_visita_min"]
            out = pd.DataFrame({
                "Visitas": g.size(), "Mediana (min)": g.median(), "Promedio (min)": g.mean(),
                "P90 (min)": g.quantile(0.9), "Máx. (min)": g.max(),
                "% < 1 min": g.apply(lambda x: (x < 1).mean() * 100),
                "% ≥ 10 min": g.apply(lambda x: (x >= 10).mean() * 100),
            }).reset_index()
            out.insert(len(claves) + 1, "Mediana (mm:ss)", out["Mediana (min)"].map(mmss))
            return out

        if vista == "Chofer":
            tab_t = (resumen_tiempos(tv, ["chofer"]).rename(columns={"chofer": "CHOFER"})
                     .sort_values("Mediana (min)", ascending=False))
        else:
            cb1, cb2 = st.columns([2, 1])
            busq = cb1.text_input("Buscar cliente (nombre o código)", key="tv_busq").strip()
            min_vis = cb2.number_input("Mínimo de visitas por cliente", 1, 50, 2, 1, key="tv_minvis")
            base_c = tv
            if busq:
                base_c = tv[tv["cliente"].astype(str).str.contains(busq, case=False, na=False, regex=False)
                            | tv["cod_cliente"].str.contains(busq, na=False, regex=False)]
            tab_t = resumen_tiempos(base_c, ["cod_cliente", "cliente"]) if len(base_c) else pd.DataFrame()
            if not tab_t.empty:
                ult = base_c.groupby(["cod_cliente", "cliente"])["fecha"].max().dt.strftime("%Y-%m-%d")
                tab_t = tab_t.merge(ult.rename("Última visita").reset_index(), on=["cod_cliente", "cliente"])
                tab_t = (tab_t[tab_t["Visitas"] >= min_vis]
                         .rename(columns={"cod_cliente": "Código cliente", "cliente": "Cliente"})
                         .sort_values("Mediana (min)", ascending=False))
        if tab_t.empty:
            st.info("No hay resultados con esos criterios.")
        else:
            st.dataframe(tab_t, use_container_width=True, hide_index=True, column_config=cfg_t)
            st.download_button("⬇️ Descargar Excel", a_excel(tab_t.round(2)),
                               f"tiempo_visita_por_{vista.lower()}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="xl_tv")

        if vista == "Cliente" and busq and len(base_c):
            st.markdown("**Detalle de visitas del cliente buscado**")
            det_c = pd.DataFrame({
                "Fecha": base_c["fecha"].dt.strftime("%Y-%m-%d"), "Chofer": base_c["chofer"],
                "Planilla": base_c["planilla"], "Código cliente": base_c["cod_cliente"],
                "Cliente": base_c["cliente"], "Arribo": base_c["arribo"].dt.strftime("%H:%M:%S"),
                "Fin": base_c["fin_visita"].dt.strftime("%H:%M:%S"),
                "Tiempo (mm:ss)": base_c["t_visita_min"].map(mmss), "Tiempo (min)": base_c["t_visita_min"].round(2),
            }).sort_values(["Fecha", "Arribo"])
            st.dataframe(det_c, use_container_width=True, hide_index=True)

        col_g1, col_g2 = st.columns(2)
        cortes = [0, 1 / 6, 1, 5, 10, np.inf]
        etiquetas = ["< 10 seg", "10 seg – 1 min", "1 – 5 min", "5 – 10 min", "≥ 10 min"]
        dist_t = (pd.cut(tv["t_visita_min"], bins=cortes, labels=etiquetas, right=False)
                  .value_counts().reindex(etiquetas).fillna(0).reset_index())
        dist_t.columns = ["Tiempo de visita", "Visitas"]
        fig_t1 = px.bar(dist_t, x="Tiempo de visita", y="Visitas", color_discrete_sequence=["#FFC933"],
                        text="Visitas", title="Distribución del tiempo de visita")
        fig_t1.update_layout(height=340, xaxis_title=None)
        col_g1.plotly_chart(fig_t1, use_container_width=True)
        evo_t = (tv.groupby("periodo")["t_visita_min"].median().reset_index()
                   .rename(columns={"t_visita_min": "Mediana (min)"}))
        fig_t2 = px.line(evo_t, x="periodo", y="Mediana (min)", markers=True,
                         color_discrete_sequence=["#D32F2F"], title=f"Mediana del tiempo de visita ({frecuencia.lower()})")
        fig_t2.update_layout(height=340, xaxis_title=None)
        col_g2.plotly_chart(fig_t2, use_container_width=True)

        st.markdown("**Visitas con mayor tiempo (top 30)**")
        top = tv.nlargest(30, "t_visita_min")
        st.dataframe(pd.DataFrame({
            "Fecha": top["fecha"].dt.strftime("%Y-%m-%d"), "Chofer": top["chofer"], "Planilla": top["planilla"],
            "Código cliente": top["cod_cliente"], "Cliente": top["cliente"],
            "Tiempo (mm:ss)": top["t_visita_min"].map(mmss), "Tiempo (min)": top["t_visita_min"].round(1)}),
            use_container_width=True, hide_index=True)

# ----------------------------------------------------------------------
# Devoluciones
# ----------------------------------------------------------------------
with tab6:
    st.subheader("Devoluciones y entregas parciales")
    dev = f[f["estado"].isin(["DEFINITELY_RETURNED", "PARTIAL_DELIVERY"])].copy()
    if dev.empty:
        st.info("No hay devoluciones ni entregas parciales para los filtros elegidos.")
    else:
        c = st.columns(3)
        c[0].metric("Devoluciones", int((dev["estado"] == "DEFINITELY_RETURNED").sum()))
        c[1].metric("Entregas parciales", int((dev["estado"] == "PARTIAL_DELIVERY").sum()))
        c[2].metric("% devoluciones", f"{pct(base['devuelta']):.1f}%")

        mot = dev["motivo_es"].fillna("Sin motivo informado").value_counts().reset_index()
        mot.columns = ["Motivo", "Visitas"]
        fig_mot = px.bar(mot.iloc[::-1], x="Visitas", y="Motivo", orientation="h", text="Visitas",
                         color_discrete_sequence=["#D32F2F"])
        fig_mot.update_layout(height=120 + 28 * len(mot), yaxis_title=None)
        st.plotly_chart(fig_mot, use_container_width=True)

        por_ch = (dev.groupby("chofer").size().rename("Devoluciones y parciales").reset_index()
                  .rename(columns={"chofer": "CHOFER"}).sort_values("Devoluciones y parciales", ascending=False))
        col_a, col_b = st.columns([1, 2])
        col_a.markdown("**Por chofer**")
        col_a.dataframe(por_ch, use_container_width=True, hide_index=True)
        col_b.markdown("**Detalle**")
        det = pd.DataFrame({
            "Fecha": dev["fecha"].dt.strftime("%Y-%m-%d"), "Chofer": dev["chofer"], "Planilla": dev["planilla"],
            "Código cliente": dev["cod_cliente"], "Cliente": dev["cliente"],
            "Resultado": dev["estado"].map({"DEFINITELY_RETURNED": "Devuelta", "PARTIAL_DELIVERY": "Parcial"}),
            "Motivo": dev["motivo_es"],
        }).sort_values(["Fecha", "Chofer"])
        col_b.dataframe(det, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Descargar Excel (detalle)", a_excel(det), "devoluciones_detalle.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="xl_dev")

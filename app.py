import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import io
import os
import json
import shutil
import re

# ----------------- CONFIGURACIÓN DE PÁGINA -----------------
st.set_page_config(
    page_title="Control de HH - Montaje Industrial",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS de alta legibilidad técnica
st.markdown("""
<style>
    .main-header {
        font-size: 2.15rem;
        font-weight: 800;
        color: #1E3A8A;
        margin-bottom: 0.15rem;
        letter-spacing: -0.02em;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.4rem;
        font-weight: 400;
    }
    .kpi-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border: 1px solid #E5E7EB;
        margin-bottom: 14px;
        transition: transform 0.15s ease-in-out;
    }
    .kpi-title {
        font-size: 0.8rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #6B7280;
    }
    .kpi-value {
        font-size: 1.85rem;
        font-weight: 800;
        margin-top: 4px;
        margin-bottom: 2px;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 0.85rem;
        font-weight: 500;
    }
    .status-badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 0.78rem;
        font-weight: 700;
    }
    .badge-green { background: #DCFCE7; color: #166534; border: 1px solid #BBF7D0; }
    .badge-yellow { background: #FEF9C3; color: #854D0E; border: 1px solid #FEF08A; }
    .badge-red { background: #FEE2E2; color: #991B1B; border: 1px solid #FECACA; }
    .badge-blue { background: #DBEAFE; color: #1E40AF; border: 1px solid #BFDBFE; }
    .source-banner {
        background: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 8px;
        padding: 8px 14px;
        font-size: 0.88rem;
        color: #166534;
        margin-bottom: 15px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
</style>
""", unsafe_allow_html=True)

# Directorios de datos y persistencia
BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

ACTIVE_FILE_PATH = os.path.join(DATA_DIR, "informe_activo.xlsx")
METADATA_JSON_PATH = os.path.join(DATA_DIR, "metadata_proyecto.json")
DEFAULT_TEMPLATE_PATH = os.path.join(BASE_DIR, "plantilla_control_hh_montaje.xlsx")

# ----------------- PARSERS Y PROCESAMIENTO -----------------

def parse_curva_avance_sheet(xls, sheet_name):
    """
    Parsea una hoja tipo 'Curva de Avance 1' típica de informes semanales de ingeniería y montaje.
    """
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    
    meta = {
        "tipo": "informe_semanal_curva",
        "hoja": sheet_name,
        "proyecto": "Montaje Industrial",
        "contrato": "",
        "plazo": "",
        "corte_semana": "",
        "fecha_corte": ""
    }
    
    # Extraer encabezados de metadatos (primeras 15 filas)
    for i in range(min(15, len(df_raw))):
        row_str = " ".join([str(x) for x in df_raw.iloc[i].values if pd.notna(x)])
        if "PROYECTO" in row_str.upper() and len(row_str) > 10:
            meta["proyecto"] = re.sub(r'[\r\n\t]+', ' ', row_str).strip()
        if "CONTRATO" in row_str.upper():
            vals = [str(x).strip() for x in df_raw.iloc[i].values if pd.notna(x) and str(x).strip() not in [':', 'N° CONTRATO', 'CONTRATO']]
            if vals:
                meta["contrato"] = vals[-1]
        if "SEM " in row_str.upper() and not meta["corte_semana"]:
            m = re.search(r'SEM\s*\d+', row_str.upper())
            if m:
                meta["corte_semana"] = m.group(0)

    # Identificar fila de encabezados
    header_idx = None
    for i in range(min(30, len(df_raw))):
        row_vals = [str(x).lower() for x in df_raw.iloc[i].values if pd.notna(x)]
        if any('hh prog' in v or 'hh ganadas' in v or 'hh gastadas' in v or 'programa oferta' in v for v in row_vals):
            header_idx = i
            break
            
    if header_idx is None:
        return None, meta

    # Sub-fila de nombres de columnas
    sub_idx = header_idx + 1 if header_idx + 1 < len(df_raw) and any('desde' in str(x).lower() or 'hh prog' in str(x).lower() for x in df_raw.iloc[header_idx + 1].values if pd.notna(x)) else header_idx

    # Mapeo de columnas
    cols_map = {}
    row_sub = df_raw.iloc[sub_idx]
    for col_i, val in enumerate(row_sub):
        if pd.isna(val):
            continue
        v_clean = str(val).lower().strip()
        if 'desde' in v_clean:
            cols_map['Fecha_Desde'] = col_i
        elif 'hasta' in v_clean:
            cols_map['Fecha_Hasta'] = col_i
        elif 'hh prog' in v_clean and 'acum' not in v_clean:
            cols_map['Prog_Inc'] = col_i
        elif 'hh prog acum' in v_clean:
            cols_map['Prog_Cum'] = col_i
        elif 'hh ganadas' in v_clean and 'acum' not in v_clean:
            cols_map['Gan_Inc'] = col_i
        elif 'hh real acum' in v_clean or ('hh gan' in v_clean and 'acum' in v_clean):
            cols_map['Gan_Cum'] = col_i
        elif 'hh gastadas' in v_clean and 'acum' not in v_clean:
            cols_map['Gast_Inc'] = col_i
        elif 'hh gastadas acum' in v_clean:
            cols_map['Gast_Cum'] = col_i
        elif 'pf' in v_clean and 'acum' in v_clean:
            cols_map['PF_Cum'] = col_i
        elif 'pf' in v_clean and 'parcial' in v_clean:
            cols_map['PF_Inc'] = col_i

    # Ubicar columna de Semana
    sem_col = 2
    for col_i, val in enumerate(df_raw.iloc[header_idx]):
        if pd.notna(val) and 'semana' in str(val).lower():
            sem_col = col_i
            break
    cols_map['Semana'] = sem_col

    # Extraer filas
    data_rows = df_raw.iloc[sub_idx + 1:].copy()
    clean_records = []
    
    for _, r in data_rows.iterrows():
        sem_val = r[cols_map['Semana']] if 'Semana' in cols_map else None
        if pd.isna(sem_val) or not str(sem_val).strip():
            continue
        sem_str = str(sem_val).strip()
        if not (sem_str.upper().startswith('SEM') or sem_str.isdigit()):
            continue
            
        rec = {'Semana': sem_str}
        for k, c_idx in cols_map.items():
            if k != 'Semana':
                rec[k] = r[c_idx] if c_idx < len(r) else None
        clean_records.append(rec)
        
    df_res = pd.DataFrame(clean_records)
    if df_res.empty:
        return None, meta

    # Limpiar numéricos
    for num_col in ['Prog_Inc', 'Prog_Cum', 'Gan_Inc', 'Gan_Cum', 'Gast_Inc', 'Gast_Cum', 'PF_Inc', 'PF_Cum']:
        if num_col in df_res.columns:
            df_res[num_col] = pd.to_numeric(df_res[num_col], errors='coerce').fillna(0.0)
        else:
            df_res[num_col] = 0.0

    # Limpiar fechas
    for d_col in ['Fecha_Desde', 'Fecha_Hasta']:
        if d_col in df_res.columns:
            df_res[d_col] = pd.to_datetime(df_res[d_col], errors='coerce')

    return df_res, meta


def parse_flat_table(df_raw):
    """
    Parsea una tabla plana estándar con mapeo automático de columnas.
    """
    cols = list(df_raw.columns)
    
    def detect_col(candidates):
        col_map = {str(c).lower().strip().replace('_', ' ').replace('-', ' '): c for c in cols}
        for cand in candidates:
            cand_clean = cand.lower().strip().replace('_', ' ').replace('-', ' ')
            for k, original in col_map.items():
                if cand_clean in k:
                    return original
        return None

    c_sem = detect_col(["semana", "sem", "week", "periodo"])
    c_fec = detect_col(["fecha", "date", "corte", "desde", "dia"])
    c_prog = detect_col(["hh_programadas", "hh programadas", "programadas", "hh prog", "hh plan", "plan"])
    c_gan = detect_col(["hh_ganadas", "hh ganadas", "ganadas", "hh gan", "valor ganado", "ev"])
    c_gast = detect_col(["hh_gastadas", "hh gastadas", "gastadas", "hh reales", "hh real", "reales", "actual"])
    c_disc = detect_col(["especialidad", "disciplina", "area", "frente", "rubro"])

    df = df_raw.copy()
    
    # Asignar nombres estandarizados
    df['Semana'] = df[c_sem].astype(str) if c_sem else [f"Sem {i+1}" for i in range(len(df))]
    df['Fecha_Desde'] = pd.to_datetime(df[c_fec], errors='coerce') if c_fec else pd.date_range(start="2026-01-01", periods=len(df), freq='W')
    df['Fecha_Hasta'] = df['Fecha_Desde'] + pd.Timedelta(days=6)
    
    df['Prog_Inc'] = pd.to_numeric(df[c_prog], errors='coerce').fillna(0.0) if c_prog else 0.0
    df['Gan_Inc'] = pd.to_numeric(df[c_gan], errors='coerce').fillna(0.0) if c_gan else 0.0
    df['Gast_Inc'] = pd.to_numeric(df[c_gast], errors='coerce').fillna(0.0) if c_gast else 0.0
    
    if c_disc:
        df['Disciplina'] = df[c_disc].astype(str).fillna("General")
    else:
        df['Disciplina'] = "General"

    # Agrupar si hay desglose por disciplina para obtener la curva general
    df_grouped = df.groupby('Semana', as_index=False, sort=False).agg({
        'Fecha_Desde': 'min',
        'Fecha_Hasta': 'max',
        'Prog_Inc': 'sum',
        'Gan_Inc': 'sum',
        'Gast_Inc': 'sum'
    })
    
    df_grouped['Prog_Cum'] = df_grouped['Prog_Inc'].cumsum()
    df_grouped['Gan_Cum'] = df_grouped['Gan_Inc'].cumsum()
    df_grouped['Gast_Cum'] = df_grouped['Gast_Inc'].cumsum()
    
    # Factor de productividad PF (Gastadas / Ganadas)
    df_grouped['PF_Cum'] = np.where(df_grouped['Gan_Cum'] > 0, df_grouped['Gast_Cum'] / df_grouped['Gan_Cum'], 1.0)
    df_grouped['PF_Inc'] = np.where(df_grouped['Gan_Inc'] > 0, df_grouped['Gast_Inc'] / df_grouped['Gan_Inc'], 1.0)

    meta = {
        "tipo": "tabla_estandar",
        "hoja": "Datos",
        "proyecto": "Control de Horas Hombre",
        "contrato": "General",
        "plazo": f"{len(df_grouped)} semanas",
        "corte_semana": df_grouped['Semana'].iloc[-1] if not df_grouped.empty else "",
        "fecha_corte": ""
    }
    
    return df_grouped, meta, df


def load_excel_file(file_path_or_bytes):
    """
    Intenta cargar el archivo evaluando primero si tiene hoja de Curva de Avance,
    y si no, como tabla plana.
    """
    try:
        xls = pd.ExcelFile(file_path_or_bytes)
        # Buscar hojas tipo Curva de Avance
        curva_sheets = [s for s in xls.sheet_names if 'curva' in s.lower() and 'avance' in s.lower()]
        for s in curva_sheets:
            df, meta = parse_curva_avance_sheet(xls, s)
            if df is not None and not df.empty:
                return df, meta, None
        
        # Si no tiene o falló, buscar hoja con datos o primera hoja
        sheet_choice = "Control_HH_Montaje" if "Control_HH_Montaje" in xls.sheet_names else xls.sheet_names[0]
        df_flat_raw = pd.read_excel(xls, sheet_name=sheet_choice)
        df_grouped, meta, df_full = parse_flat_table(df_flat_raw)
        meta["hoja"] = sheet_choice
        return df_grouped, meta, df_full
    except Exception as e:
        st.error(f"Error procesando archivo Excel: {e}")
        return None, None, None

# ----------------- SIDEBAR Y CARGA DE ARCHIVO -----------------

with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/crane.png", width=64)
    st.title("Montaje Industrial")
    st.caption("**Sistema de Control de HH & Valor Ganado**")
    st.markdown("---")

    st.subheader("📂 Carga del Informe Semanal")
    st.write("Sube el entregable semanal (.xlsx) para actualizar la Curva S y los indicadores.")
    
    uploaded_file = st.file_uploader(
        "Subir entregable semanal de Excel", 
        type=["xlsx", "xls"],
        help="Arrastra aquí el archivo semanal entregable para actualizar el proyecto."
    )

    # Si se sube un nuevo archivo, guardarlo como activo en data/
    if uploaded_file is not None:
        try:
            with open(ACTIVE_FILE_PATH, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Guardar metadatos de carga
            info_carga = {
                "nombre_archivo": uploaded_file.name,
                "fecha_carga": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "tamano_bytes": uploaded_file.size
            }
            with open(METADATA_JSON_PATH, "w", encoding="utf-8") as f:
                json.dump(info_carga, f, ensure_ascii=False, indent=2)
                
            st.sidebar.success(f"✅ Nuevo informe cargado: {uploaded_file.name}")
        except Exception as e:
            st.sidebar.error(f"Error guardando archivo: {e}")

    # Selector de fuente de datos si existe archivo guardado
    tiene_archivo_guardado = os.path.exists(ACTIVE_FILE_PATH)
    
    opciones_fuente = []
    if tiene_archivo_guardado:
        # Obtener nombre original
        nombre_guardado = "Informe Guardado"
        if os.path.exists(METADATA_JSON_PATH):
            try:
                with open(METADATA_JSON_PATH, "r", encoding="utf-8") as f:
                    meta_guardada = json.load(f)
                    nombre_guardado = meta_guardada.get("nombre_archivo", "Informe Guardado")
            except Exception:
                pass
        opciones_fuente.append(f"📌 {nombre_guardado} (Activo)")
    
    opciones_fuente.append("🏗️ Proyecto de Demostración (Simulado)")

    fuente_seleccionada = st.radio("Fuente de datos en uso:", opciones_fuente, index=0)

    st.markdown("---")
    st.subheader("⚙️ Configuración")
    horas_semana_trabajador = st.number_input(
        "Jornada semanal (HH/persona/sem)", 
        min_value=30, 
        max_value=70, 
        value=45, 
        step=1,
        help="Horas hombre semanales por trabajador para estimar la dotación de cuadrillas equivalentes."
    )

    st.markdown("---")
    # Botón de Descarga de Plantilla
    if os.path.exists(DEFAULT_TEMPLATE_PATH):
        with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
            template_bytes = f.read()
        st.download_button(
            label="📥 Descargar Plantilla Excel Base",
            data=template_bytes,
            file_name="plantilla_control_hh_montaje.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descargar archivo Excel modelo con formato listo."
        )

# ----------------- CARGA DE DATOS EFECTIVA -----------------

df_data = None
metadata = None
df_full_detail = None
archivo_origen_nombre = ""

if "📌" in fuente_seleccionada and os.path.exists(ACTIVE_FILE_PATH):
    df_data, metadata, df_full_detail = load_excel_file(ACTIVE_FILE_PATH)
    archivo_origen_nombre = fuente_seleccionada.replace("📌", "").replace("(Activo)", "").strip()
else:
    # Cargar proyecto demo
    if os.path.exists(DEFAULT_TEMPLATE_PATH):
        df_data, metadata, df_full_detail = load_excel_file(DEFAULT_TEMPLATE_PATH)
        archivo_origen_nombre = "Proyecto Simulado de Demostración (14 Semanas)"

if df_data is None or df_data.empty:
    st.error("No se pudo cargar la información del proyecto. Por favor carga un archivo Excel válido.")
    st.stop()

# ----------------- PROCESAMIENTO DE SEMANAS Y CORTE ACTUAL -----------------

# Identificar corte actual (última semana donde hubo ejecución: Gast_Inc > 0 o Gan_Inc > 0)
ejecutadas_mask = (df_data["Gast_Inc"] > 0) | (df_data["Gan_Inc"] > 0)
df_ejecutadas = df_data[ejecutadas_mask]

if not df_ejecutadas.empty:
    indice_corte = df_ejecutadas.index[-1]
    fila_corte = df_data.loc[indice_corte]
else:
    indice_corte = df_data.index[-1]
    fila_corte = df_data.iloc[-1]

semana_corte_label = fila_corte["Semana"]
fecha_corte_label = ""
if pd.notna(fila_corte.get("Fecha_Hasta")):
    try:
        fecha_corte_label = pd.to_datetime(fila_corte["Fecha_Hasta"]).strftime("%d-%b-%Y")
    except Exception:
        fecha_corte_label = str(fila_corte.get("Fecha_Hasta"))[:10]

# Totales y Métricas al Corte
prog_total_proyecto = df_data["Prog_Cum"].iloc[-1]
prog_al_corte = fila_corte["Prog_Cum"]
gan_al_corte = fila_corte["Gan_Cum"]
gast_al_corte = fila_corte["Gast_Cum"]

# Variaciones acumuladas
sv_acum = gan_al_corte - prog_al_corte  # Variación de cronograma (HH)
cv_acum = gan_al_corte - gast_al_corte  # Variación de rendimiento/costo (HH)

# Índices
spi_actual = (gan_al_corte / prog_al_corte) if prog_al_corte > 0 else 1.0
cpi_actual = (gan_al_corte / gast_al_corte) if gast_al_corte > 0 else 1.0

# Factor de Productividad (PF = Horas Gastadas / Horas Ganadas en montaje chileno/minero)
pf_actual = (gast_al_corte / gan_al_corte) if gan_al_corte > 0 else (fila_corte.get("PF_Cum", 1.0))

# Avances porcentuales
pct_avance_prog = (prog_al_corte / prog_total_proyecto * 100) if prog_total_proyecto > 0 else 0.0
pct_avance_real = (gan_al_corte / prog_total_proyecto * 100) if prog_total_proyecto > 0 else 0.0

# Pronóstico al término (EAC y VAC)
eac_hh = (prog_total_proyecto / cpi_actual) if cpi_actual > 0 else prog_total_proyecto
vac_hh = prog_total_proyecto - eac_hh

# Dotación equivalente última semana
dotacion_real_semana = (fila_corte["Gast_Inc"] / horas_semana_trabajador) if horas_semana_trabajador > 0 else 0
dotacion_prog_semana = (fila_corte["Prog_Inc"] / horas_semana_trabajador) if horas_semana_trabajador > 0 else 0

# ----------------- ENCABEZADO Y BANNER INFORMATIVO -----------------

st.markdown('<div class="main-header">🏗️ Control de Horas Hombre (HH) — Montaje Industrial</div>', unsafe_allow_html=True)

# Proyecto y detalles contractuales
nombre_proy = metadata.get("proyecto", "Proyecto Montaje Industrial")
contrato_txt = f" • Contrato: <b>{metadata.get('contrato')}</b>" if metadata.get('contrato') else ""
st.markdown(f'<div class="sub-header">{nombre_proy}{contrato_txt} • Corte activo: <b>{semana_corte_label}</b> ({fecha_corte_label})</div>', unsafe_allow_html=True)

# Banner de archivo activo
st.markdown(f"""
<div class="source-banner">
    <div>
        📄 <b>Informe Semanal Activo:</b> {archivo_origen_nombre} &nbsp;|&nbsp; 
        <b>Hoja analizada:</b> <code>{metadata.get('hoja')}</code> &nbsp;|&nbsp;
        <b>Total Semanas en Planilla:</b> {len(df_data)} semanas
    </div>
    <div style="font-weight: 600;">
        ⏱️ Semana Activa: <span style="background: #166534; color: white; padding: 2px 8px; border-radius: 6px;">{semana_corte_label}</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ----------------- TARJETAS KPI DE RESUMEN EJECUTIVO -----------------

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">📅 HH Programadas (PV)</div>
        <div class="kpi-value" style="color: #1E40AF;">{prog_al_corte:,.0f} <span style="font-size: 0.95rem; font-weight: 600; color:#6B7280;">HH</span></div>
        <div class="kpi-sub" style="color: #374151;">
            Plan acumulado: <b>{pct_avance_prog:.1f}%</b> de {prog_total_proyecto:,.0f} HH
        </div>
    </div>
    """, unsafe_allow_html=True)

with k2:
    color_gan = "#059669" if gan_al_corte >= prog_al_corte else "#D97706"
    desvio_txt = f"{sv_acum:+,.0f} HH vs plan"
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">🎯 HH Ganadas (EV - Avance Real)</div>
        <div class="kpi-value" style="color: {color_gan};">{gan_al_corte:,.0f} <span style="font-size: 0.95rem; font-weight: 600; color:#6B7280;">HH</span></div>
        <div class="kpi-sub" style="color: #374151;">
            Avance Físico: <b>{pct_avance_real:.1f}%</b> ({desvio_txt})
        </div>
    </div>
    """, unsafe_allow_html=True)

with k3:
    color_gast = "#DC2626" if gast_al_corte > gan_al_corte else "#059669"
    balance_txt = "Sobreconsumo" if cv_acum < 0 else "Ahorro"
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">⏱️ HH Gastadas (AC - Reales)</div>
        <div class="kpi-value" style="color: {color_gast};">{gast_al_corte:,.0f} <span style="font-size: 0.95rem; font-weight: 600; color:#6B7280;">HH</span></div>
        <div class="kpi-sub" style="color: {color_gast};">
            {balance_txt}: <b>{abs(cv_acum):,.0f} HH</b> ({cv_acum:+,.0f})
        </div>
    </div>
    """, unsafe_allow_html=True)

with k4:
    # Semáforo de productividad
    if cpi_actual >= 1.0 and spi_actual >= 1.0:
        badge_cls = "badge-green"
        estado_txt = "En Plazo y Eficiente"
    elif cpi_actual >= 0.90 and spi_actual >= 0.90:
        badge_cls = "badge-yellow"
        estado_txt = "Alerta de Desviación"
    else:
        badge_cls = "badge-red"
        estado_txt = "Crítico: Requiere Mitigación"

    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">⚡ Rendimiento & Productividad</div>
        <div style="display: flex; gap: 14px; align-items: baseline; margin-top: 4px;">
            <div><span style="font-size: 0.75rem; color:#6B7280; font-weight: 600;">CPI:</span> <span class="kpi-value" style="font-size: 1.4rem; color: {'#059669' if cpi_actual >= 1 else '#DC2626'};">{cpi_actual:.2f}</span></div>
            <div><span style="font-size: 0.75rem; color:#6B7280; font-weight: 600;">SPI:</span> <span class="kpi-value" style="font-size: 1.4rem; color: {'#059669' if spi_actual >= 1 else '#D97706'};">{spi_actual:.2f}</span></div>
            <div><span style="font-size: 0.75rem; color:#6B7280; font-weight: 600;">PF:</span> <span class="kpi-value" style="font-size: 1.4rem; color: {'#059669' if pf_actual <= 1 else '#DC2626'};">{pf_actual:.2f}</span></div>
        </div>
        <div style="margin-top: 5px;">
            <span class="status-badge {badge_cls}">{estado_txt}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ----------------- PESTAÑAS PRINCIPALES DE ANÁLISIS -----------------

tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Curva 'S' Acumulada",
    "📊 Histograma Semanal y Dotación",
    "🎯 Rendimiento y Productividad (CPI / SPI / PF)",
    "📋 Tabla Detallada y Exportación"
])

# ================= TAB 1: CURVA 'S' =================
with tab1:
    st.subheader("Curva 'S' Acumulada de Horas Hombre")
    st.caption("Comparación visual de la línea base contractual (HH Programadas), avance físico real (HH Ganadas) y consumo en obra (HH Gastadas).")

    # Controles de visualización de la curva
    col_c1, col_c2 = st.columns([3, 1])
    with col_c2:
        ver_futuras = st.checkbox(
            "Mostrar semanas futuras planificadas", 
            value=True, 
            help="Muestra el plan contractual completo hasta el término del contrato."
        )

    # Preparar dataset para el gráfico
    if ver_futuras:
        df_plot = df_data.copy()
    else:
        df_plot = df_data.loc[:indice_corte].copy()

    # Para las semanas futuras (posteriores al corte actual), Ganadas y Gastadas no deben graficarse como líneas planas
    df_plot["Gan_Plot"] = np.where(df_plot.index <= indice_corte, df_plot["Gan_Cum"], np.nan)
    df_plot["Gast_Plot"] = np.where(df_plot.index <= indice_corte, df_plot["Gast_Cum"], np.nan)

    fig_s = go.Figure()

    # 1. Traza Programada (Plan Base)
    fig_s.add_trace(go.Scatter(
        x=df_plot["Semana"],
        y=df_plot["Prog_Cum"],
        mode="lines+markers",
        name="HH Programadas (PV - Plan Base)",
        line=dict(color="#2563EB", width=3, dash="dash"),
        marker=dict(size=5),
        hovertemplate="<b>Programadas</b>: %{y:,.0f} HH<br>Semana: %{x}<extra></extra>"
    ))

    # 2. Traza Ganada (Avance Real)
    fig_s.add_trace(go.Scatter(
        x=df_plot["Semana"],
        y=df_plot["Gan_Plot"],
        mode="lines+markers",
        name="HH Ganadas (EV - Avance Físico Real)",
        line=dict(color="#059669", width=3.5),
        marker=dict(size=6),
        hovertemplate="<b>Ganadas</b>: %{y:,.0f} HH<br>Semana: %{x}<extra></extra>"
    ))

    # 3. Traza Gastada (Costo Real)
    fig_s.add_trace(go.Scatter(
        x=df_plot["Semana"],
        y=df_plot["Gast_Plot"],
        mode="lines+markers",
        name="HH Gastadas (AC - Horas Reales Invertidas)",
        line=dict(color="#DC2626", width=3),
        marker=dict(size=5),
        hovertemplate="<b>Gastadas</b>: %{y:,.0f} HH<br>Semana: %{x}<extra></extra>"
    ))

    # Línea vertical del corte actual
    fig_s.add_shape(
        type="line",
        x0=semana_corte_label,
        x1=semana_corte_label,
        y0=0,
        y1=1,
        yref="paper",
        line=dict(color="#4B5563", width=1.5, dash="dot")
    )
    fig_s.add_annotation(
        x=semana_corte_label,
        y=1,
        yref="paper",
        text=f"Corte: {semana_corte_label}",
        showarrow=False,
        xanchor="right",
        yanchor="top",
        font=dict(color="#4B5563", size=11),
        bgcolor="rgba(255, 255, 255, 0.85)"
    )

    fig_s.update_layout(
        title="Evolución Acumulada de Horas Hombre (Curva S)",
        xaxis_title="Semana de Control",
        yaxis_title="Horas Hombre Acumuladas (HH)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template="plotly_white",
        height=540,
        margin=dict(l=40, r=40, t=60, b=40)
    )

    st.plotly_chart(fig_s, use_container_width=True)

    # Gráficos de Variaciones Semanales Acumuladas
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        df_var = df_data.loc[:indice_corte].copy()
        df_var["SV_Cum"] = df_var["Gan_Cum"] - df_var["Prog_Cum"]
        fig_sv = px.bar(
            df_var,
            x="Semana",
            y="SV_Cum",
            title="Variación de Cronograma: SV = Ganadas - Programadas (HH)",
            color="SV_Cum",
            color_continuous_scale=["#EF4444", "#10B981"],
            color_continuous_midpoint=0,
            labels={"SV_Cum": "Desvío en HH", "Semana": "Semana"}
        )
        fig_sv.update_layout(template="plotly_white", height=320, coloraxis_showscale=False)
        st.plotly_chart(fig_sv, use_container_width=True)

    with col_v2:
        df_var["CV_Cum"] = df_var["Gan_Cum"] - df_var["Gast_Cum"]
        fig_cv = px.bar(
            df_var,
            x="Semana",
            y="CV_Cum",
            title="Variación de Rendimiento: CV = Ganadas - Gastadas (HH)",
            color="CV_Cum",
            color_continuous_scale=["#EF4444", "#10B981"],
            color_continuous_midpoint=0,
            labels={"CV_Cum": "Balance HH", "Semana": "Semana"}
        )
        fig_cv.update_layout(template="plotly_white", height=320, coloraxis_showscale=False)
        st.plotly_chart(fig_cv, use_container_width=True)

# ================= TAB 2: HISTOGRAMA Y DOTACIÓN =================
with tab2:
    st.subheader("Histograma Periódico de HH y Dotación de Mano de Obra")
    st.caption("Consumo semanal de recursos de montaje y dotación equivalente de personal en faena.")

    df_hist = df_data.loc[:indice_corte].copy()
    df_hist["Dotacion_Real"] = df_hist["Gast_Inc"] / horas_semana_trabajador
    df_hist["Dotacion_Prog"] = df_hist["Prog_Inc"] / horas_semana_trabajador

    fig_hist = go.Figure()

    fig_hist.add_trace(go.Bar(
        x=df_hist["Semana"],
        y=df_hist["Prog_Inc"],
        name="HH Programadas en la Semana",
        marker_color="#93C5FD",
        hovertemplate="%{y:,.0f} HH<extra></extra>"
    ))

    fig_hist.add_trace(go.Bar(
        x=df_hist["Semana"],
        y=df_hist["Gan_Inc"],
        name="HH Ganadas en la Semana",
        marker_color="#34D399",
        hovertemplate="%{y:,.0f} HH<extra></extra>"
    ))

    fig_hist.add_trace(go.Bar(
        x=df_hist["Semana"],
        y=df_hist["Gast_Inc"],
        name="HH Gastadas en la Semana",
        marker_color="#F87171",
        hovertemplate="%{y:,.0f} HH<extra></extra>"
    ))

    # Dotación en eje Y secundario
    fig_hist.add_trace(go.Scatter(
        x=df_hist["Semana"],
        y=df_hist["Dotacion_Real"],
        name=f"Dotación Real Estimada (pers @ {horas_semana_trabajador}h/sem)",
        yaxis="y2",
        mode="lines+markers",
        line=dict(color="#111827", width=2.5),
        marker=dict(size=6)
    ))

    fig_hist.update_layout(
        title="Histograma Semanal de Horas Hombre y Dotación Equivalente",
        barmode="group",
        xaxis_title="Semana de Control",
        yaxis_title="Horas Hombre Semanales (HH)",
        yaxis2=dict(
            title="Dotación (Trabajadores en Faena)",
            overlaying="y",
            side="right",
            showgrid=False
        ),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template="plotly_white",
        height=500
    )

    st.plotly_chart(fig_hist, use_container_width=True)

    # Tarjetas de resumen de dotación
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric(
            "Dotación en Faena al Último Corte", 
            f"{dotacion_real_semana:.0f} trabajadores", 
            f"{(dotacion_real_semana - dotacion_prog_semana):+.0f} vs programada"
        )
    with m2:
        peak_gast = df_hist["Gast_Inc"].max()
        sem_peak = df_hist.loc[df_hist["Gast_Inc"].idxmax()]["Semana"]
        st.metric("Pico Semanal (Peak)", f"{peak_gast:,.0f} HH", f"{peak_gast/horas_semana_trabajador:.0f} trabajadores en {sem_peak}")
    with m3:
        prom_gast = df_hist[df_hist["Gast_Inc"] > 0]["Gast_Inc"].mean() if len(df_hist[df_hist["Gast_Inc"] > 0]) > 0 else 0
        st.metric("Consumo Promedio Semanal", f"{prom_gast:,.0f} HH/sem", f"{prom_gast/horas_semana_trabajador:.0f} trabajadores prom.")

# ================= TAB 3: CPI & SPI & PF =================
with tab3:
    st.subheader("Índices de Rendimiento: CPI, SPI y Factor de Productividad (PF)")
    st.caption("Comportamiento histórico de la productividad de las cuadrillas de montaje y cumplimiento de plazos.")

    df_idx = df_data.loc[:indice_corte].copy()
    df_idx["CPI_Calc"] = np.where(df_idx["Gast_Cum"] > 0, df_idx["Gan_Cum"] / df_idx["Gast_Cum"], 1.0)
    df_idx["SPI_Calc"] = np.where(df_idx["Prog_Cum"] > 0, df_idx["Gan_Cum"] / df_idx["Prog_Cum"], 1.0)
    
    # PF Acumulado (Gastadas / Ganadas)
    if "PF_Cum" in df_idx.columns and df_idx["PF_Cum"].max() > 0:
        df_idx["PF_Display"] = df_idx["PF_Cum"]
    else:
        df_idx["PF_Display"] = np.where(df_idx["Gan_Cum"] > 0, df_idx["Gast_Cum"] / df_idx["Gan_Cum"], 1.0)

    fig_idx = go.Figure()

    fig_idx.add_trace(go.Scatter(
        x=df_idx["Semana"],
        y=df_idx["CPI_Calc"],
        mode="lines+markers",
        name="CPI Acumulado (Ganadas / Gastadas)",
        line=dict(color="#2563EB", width=3),
        marker=dict(size=6)
    ))

    fig_idx.add_trace(go.Scatter(
        x=df_idx["Semana"],
        y=df_idx["SPI_Calc"],
        mode="lines+markers",
        name="SPI Acumulado (Ganadas / Programadas)",
        line=dict(color="#059669", width=3),
        marker=dict(size=6)
    ))

    fig_idx.add_trace(go.Scatter(
        x=df_idx["Semana"],
        y=df_idx["PF_Display"],
        mode="lines+markers",
        name="Factor Productividad PF (Gastadas / Ganadas)",
        line=dict(color="#D97706", width=2.5, dash="dot"),
        marker=dict(size=5)
    ))

    fig_idx.add_hline(y=1.0, line_dash="dash", line_color="#DC2626", annotation_text="Meta 1.00")

    fig_idx.update_layout(
        title="Tendencia de CPI, SPI y Factor de Productividad (PF)",
        xaxis_title="Semana de Control",
        yaxis_title="Índice",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        template="plotly_white",
        height=450
    )

    st.plotly_chart(fig_idx, use_container_width=True)

    # Diagnóstico y Pronóstico
    c_p1, c_p2 = st.columns([1.2, 1])
    with c_p1:
        st.markdown("#### Matriz de Posicionamiento (SPI vs CPI)")
        fig_scatter = go.Figure()

        # Cuadrantes
        fig_scatter.add_shape(type="rect", x0=1.0, y0=1.0, x1=1.5, y1=1.5, fillcolor="rgba(16, 185, 129, 0.15)", line_width=0)
        fig_scatter.add_shape(type="rect", x0=0.0, y0=1.0, x1=1.0, y1=1.5, fillcolor="rgba(245, 158, 11, 0.15)", line_width=0)
        fig_scatter.add_shape(type="rect", x0=1.0, y0=0.0, x1=1.5, y1=1.0, fillcolor="rgba(249, 115, 22, 0.15)", line_width=0)
        fig_scatter.add_shape(type="rect", x0=0.0, y0=0.0, x1=1.0, y1=1.0, fillcolor="rgba(239, 68, 68, 0.15)", line_width=0)

        # Punto actual
        fig_scatter.add_trace(go.Scatter(
            x=[spi_actual],
            y=[cpi_actual],
            mode="markers+text",
            text=[f"Posición Actual<br>(SPI: {spi_actual:.2f}, CPI: {cpi_actual:.2f})"],
            textposition="top center",
            marker=dict(size=16, color="#1E3A8A", symbol="diamond"),
            name="Estado Actual"
        ))

        fig_scatter.update_layout(
            xaxis=dict(title="SPI (Cronograma: < 1 Atrasado | > 1 Adelantado)", range=[0, 1.5]),
            yaxis=dict(title="CPI (Rendimiento: < 1 Sobreconsumo | > 1 Ahorro)", range=[0, 1.5]),
            template="plotly_white",
            height=360,
            showlegend=False
        )
        st.plotly_chart(fig_scatter, use_container_width=True)

    with c_p2:
        st.markdown("#### Proyección al Término del Contrato")
        st.markdown(f"""
        - **Presupuesto Total Contractual**: `{prog_total_proyecto:,.0f} HH`
        - **Estimado al Término (EAC)**: `{eac_hh:,.0f} HH`
        - **Varianza al Término (VAC)**: `{vac_hh:+,.0f} HH`
        - **Factor de Productividad Actual**: `{pf_actual:.2f}` *(gasto de {pf_actual:.2f} HH reales por cada HH valorizada)*
        """)
        if cpi_actual < 1.0:
            st.error(f"⚠️ Al rendimiento actual (CPI {cpi_actual:.2f}), se proyecta un sobrecosto de **{abs(vac_hh):,.0f} HH** adicionales sobre el presupuesto base.")
        else:
            st.success(f"✅ Con el rendimiento actual (CPI {cpi_actual:.2f}), el proyecto se proyecta con un ahorro de **{vac_hh:,.0f} HH**.")

# ================= TAB 4: TABLA DETALLADA =================
with tab4:
    st.subheader("Registro Semana a Semana y Exportación")
    st.caption("Detalle cronológico completo con indicadores calculados.")

    df_export = df_data.copy()
    df_export["CPI"] = np.where(df_export["Gast_Cum"] > 0, df_export["Gan_Cum"] / df_export["Gast_Cum"], np.nan)
    df_export["SPI"] = np.where(df_export["Prog_Cum"] > 0, df_export["Gan_Cum"] / df_export["Prog_Cum"], np.nan)
    df_export["SV_HH"] = df_export["Gan_Cum"] - df_export["Prog_Cum"]
    df_export["CV_HH"] = df_export["Gan_Cum"] - df_export["Gast_Cum"]

    cols_show = ["Semana", "Prog_Inc", "Prog_Cum", "Gan_Inc", "Gan_Cum", "Gast_Inc", "Gast_Cum", "CPI", "SPI", "SV_HH", "CV_HH"]
    if "Fecha_Desde" in df_export.columns:
        cols_show.insert(1, "Fecha_Desde")

    # Formateo de tabla
    df_view = df_export[cols_show].copy()
    if "Fecha_Desde" in df_view.columns:
        df_view["Fecha_Desde"] = df_view["Fecha_Desde"].dt.strftime("%Y-%m-%d").fillna("")

    st.dataframe(
        df_view.rename(columns={
            "Prog_Inc": "Prog Sem (HH)",
            "Prog_Cum": "Prog Acum (HH)",
            "Gan_Inc": "Gan Sem (HH)",
            "Gan_Cum": "Gan Acum (HH)",
            "Gast_Inc": "Gast Sem (HH)",
            "Gast_Cum": "Gast Acum (HH)",
            "SV_HH": "Desvío SV (HH)",
            "CV_HH": "Balance CV (HH)"
        }),
        use_container_width=True,
        hide_index=True
    )

    # Descarga a Excel
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, sheet_name='Analisis_Semanas_HH', index=False)
    excel_bytes = buf.getvalue()

    st.download_button(
        label="📥 Descargar Análisis Completo en Excel",
        data=excel_bytes,
        file_name=f"analisis_hh_{semana_corte_label}_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ----------------- PIE DE PÁGINA -----------------
st.markdown("---")
st.caption(f"🏗️ **Control de HH y Curva S — Montaje Industrial** • Última actualización: {datetime.now().strftime('%d-%m-%Y %H:%M')}")

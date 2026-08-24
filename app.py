
import streamlit as st
import pandas as pd
import plotly.express as px
from openpyxl import load_workbook
import os, tempfile, shutil, zipfile, re
import xml.etree.ElementTree as ET
from datetime import datetime
from io import BytesIO
from supabase import create_client

st.set_page_config(
    page_title="Control Fases SFCO211",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)



# v19 · Estilo visual uniforme
st.markdown("""
<style>
/* Menú lateral azul más claro */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div:first-child {
    background: #6F91B3 !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span {
    color: #FFFFFF !important;
}

/* Elemento seleccionado del menú: contraste suave */
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
    background: rgba(255,255,255,0.16) !important;
    border-radius: 8px !important;
}

/* Tamaño visual uniforme para gráficos */
[data-testid="stVegaLiteChart"],
[data-testid="stPlotlyChart"],
[data-testid="stPyplot"] {
    width: 100% !important;
    max-width: 100% !important;
}

/* Área de trabajo estable */
.block-container {
    max-width: 1450px !important;
}
</style>
""", unsafe_allow_html=True)

BASE = os.path.join(os.path.dirname(__file__), "CONTROL_FASES_SFCO211.xlsx")
LOGO = os.path.join(os.path.dirname(__file__), "logo_san_francisco.png")
PHASES = ["FASE1", "FASE2", "FASE3", "FASE4"]

@st.cache_resource(show_spinner=False)
def get_supabase():
    try:
        url = str(st.secrets["supabase"]["url"]).strip()
        key = str(st.secrets["supabase"]["service_role_key"]).strip()

        if not url.startswith("https://") or not url.endswith(".supabase.co"):
            raise ValueError("La URL de Supabase no tiene el formato esperado.")

        if not key.startswith("sb_secret_"):
            raise ValueError("La clave secreta no comienza con sb_secret_.")

        client = create_client(url, key)
        client.table("phase_updates").select("id").limit(1).execute()
        return client
    except Exception as e:
        st.session_state["supabase_error"] = str(e)
        return None

def db_ready():
    return get_supabase() is not None

def supabase_status():
    if db_ready():
        return True, "Conectada"
    return False, st.session_state.get("supabase_error", "No se pudo inicializar Supabase.")

def db_phase_updates():
    client = get_supabase()
    if client is None:
        return []
    try:
        res = client.table("phase_updates").select("*").execute()
        return res.data or []
    except Exception:
        return []

def db_upsert_phase_update(phase, excel_row, activity, percent, username):
    client = get_supabase()
    if client is None:
        return False
    payload = {
        "phase": phase,
        "excel_row": int(excel_row),
        "activity": str(activity),
        "percent": float(percent),
        "updated_by": str(username),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    try:
        client.table("phase_updates").upsert(
            payload, on_conflict="phase,excel_row,activity"
        ).execute()
        return True
    except Exception:
        return False

def db_weekly_history():
    client = get_supabase()
    if client is None:
        return []
    try:
        res = client.table("weekly_history").select("*").order("update_date").execute()
        return res.data or []
    except Exception:
        return []

def db_latest_weekly_snapshot():
    """Devuelve el último corte oficial directamente desde Supabase, sin cache.
    Esta función es la fuente maestra para Dashboard, Ruta Crítica y Gantt.
    """
    client = get_supabase()
    if client is None:
        return None
    try:
        res = (
            client.table("weekly_history")
            .select("*")
            .order("update_date", desc=True)
            .limit(1)
            .execute()
        )
        rows = res.data or []
        return rows[0] if rows else None
    except Exception:
        return None

def db_register_weekly(summaries, general, username):
    client = get_supabase()
    if client is None:
        return False, "Base de datos no configurada."
    date_value = datetime.now().date().isoformat()
    payload = {
        "update_date": date_value,
        "general": float(general),
        "fase1": float(summaries["FASE1"]),
        "fase2": float(summaries["FASE2"]),
        "fase3": float(summaries["FASE3"]),
        "fase4": float(summaries["FASE4"]),
        "updated_by": str(username),
    }
    try:
        client.table("weekly_history").upsert(
            payload, on_conflict="update_date"
        ).execute()
        return True, f"Actualización semanal guardada en línea: {date_value}."
    except Exception as e:
        return False, f"No se pudo guardar en la base de datos: {e}"


st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f4f7fb;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#071f36 0%,#0a3b63 100%);}
[data-testid="stSidebar"] *{color:#fff;}
.block-container{padding-top:1.1rem;max-width:1700px;}
.sf-title{font-size:30px;font-weight:850;color:#10233d;letter-spacing:-.4px;margin-bottom:1px}
.sf-sub{font-size:14px;color:#708096;margin-bottom:22px}
.card{background:#fff;border:1px solid #e2e8f0;border-radius:15px;padding:18px;box-shadow:0 4px 16px rgba(15,35,60,.04)}
.klabel{font-size:11px;font-weight:800;color:#77849a;text-transform:uppercase}
.kvalue{font-size:30px;font-weight:850;color:#10233d;margin-top:5px}
.good{color:#179b59}.warn{color:#e79b13}.bad{color:#df4545}.blue{color:#1678e7}
.section{font-size:17px;font-weight:850;color:#1b2d47;margin:14px 0 10px}
.phase-card{background:#fff;border:1px solid #dfe6ef;border-radius:16px;padding:16px 18px;
box-shadow:0 4px 14px rgba(15,35,60,.05);margin-bottom:8px}
.phase-name{font-size:12px;font-weight:800;color:#6f7d92;text-transform:uppercase}
.phase-pct{font-size:35px;font-weight:900;color:#10233d;line-height:1.1;margin-top:4px}
.general-box{background:linear-gradient(135deg,#102f50 0%,#1769aa 100%);border-radius:17px;
padding:19px 22px;color:white;box-shadow:0 5px 18px rgba(15,35,60,.12)}
.general-label{font-size:12px;font-weight:800;opacity:.82;text-transform:uppercase}
.general-pct{font-size:39px;font-weight:900;line-height:1.05;margin-top:4px}

.stButton>button,.stDownloadButton>button{border-radius:9px;font-weight:750}
div[data-testid="stDataFrame"],div[data-testid="stDataEditor"]{
 background:#fff;border-radius:12px;border:1px solid #e2e8f0;padding:5px;
}
div[data-testid="stDataFrame"] {font-size:12px;}

@media (max-width: 768px) {
  .block-container {padding: .65rem .65rem 1.5rem .65rem;}
  .sf-title {font-size:22px !important; line-height:1.1;}
  .sf-sub {font-size:12px !important; margin-bottom:12px;}
  .card {padding:12px; min-height:96px;}
  .kvalue {font-size:24px !important;}
  .phase-pct {font-size:27px !important;}
  .general-pct {font-size:31px !important;}
  div[data-testid="stHorizontalBlock"] {gap:.45rem;}
  [data-testid="stSidebar"] img {max-height:150px; object-fit:contain;}
  .stButton>button,.stDownloadButton>button {
      min-height:44px;
      font-size:15px;
  }
}
</style>
""", unsafe_allow_html=True)

APP_VERSION = "v26"

# Al cambiar de versión, reemplazar cualquier copia temporal antigua por
# la plantilla depurada incluida en esta versión. Los avances reales se
# vuelven a aplicar desde Supabase.
if "workbook_path" not in st.session_state:
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    shutil.copy2(BASE, p)
    st.session_state.workbook_path = p
    st.session_state.app_version = APP_VERSION
elif st.session_state.get("app_version") != APP_VERSION:
    shutil.copy2(BASE, st.session_state.workbook_path)
    st.session_state.app_version = APP_VERSION

def get_users():
    """
    Usuarios almacenados en Streamlit Secrets.
    Formato:
    [users.admin]
    password = "..."
    name = "..."
    role = "Administrador"
    """
    try:
        users = st.secrets.get("users", {})
        return users
    except Exception:
        return {}

def login_screen():
    users = get_users()

    if not users:
        st.error("No hay usuarios configurados todavía.")
        st.info(
            "Configura los usuarios en Streamlit → App settings → Secrets. "
            "Usa el formato incluido en el archivo secrets.toml.example."
        )
        st.stop()

    st.markdown("## 🔐 Ingreso al Control de Obra")
    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Clave", type="password")
        submitted = st.form_submit_button("INGRESAR", use_container_width=True)

    if submitted:
        if username in users and str(users[username].get("password", "")) == password:
            st.session_state.authenticated = True
            st.session_state.username = username
            st.session_state.user_name = users[username].get("name", username)
            st.session_state.user_role = users[username].get("role", "Usuario")
            st.session_state.user_phase = users[username].get("phase", "ALL")
            st.rerun()
        else:
            st.error("Usuario o clave incorrectos.")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    if os.path.exists(LOGO):
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            st.image(LOGO, use_container_width=True)
    login_screen()
    st.stop()

def phase_scale(sheet_name):
    # El archivo usa FASE1 en 0-100 y FASE2/3/4 en 0-1.
    return 100.0 if sheet_name == "FASE1" else 1.0

def to_percent(value, phase):
    try:
        if value is None or value == "":
            return 0.0
        x = float(value)
        return x if phase_scale(phase) == 100.0 else x * 100.0
    except:
        return 0.0

def from_percent(value, phase):
    x = max(0.0, min(100.0, float(value)))
    return x if phase_scale(phase) == 100.0 else x / 100.0

@st.cache_data(show_spinner=False)
def get_sheet_names(path):
    wb = load_workbook(path, read_only=True, data_only=False)
    names = wb.sheetnames
    wb.close()
    return names

@st.cache_data(show_spinner=False, ttl=3)
def load_phase(path, phase):
    wb = load_workbook(path, read_only=False, data_only=False)
    ws = wb[phase]
    headers = [c.value for c in ws[1]]
    rows, excel_rows = [], []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(v is not None for v in vals):
            continue
        rows.append(vals)
        excel_rows.append(r)
    wb.close()
    df = pd.DataFrame(rows, columns=headers)
    df["_excel_row"] = excel_rows

    for upd in db_phase_updates():
        try:
            if upd.get("phase") != phase:
                continue
            excel_row = int(upd.get("excel_row"))
            activity = upd.get("activity")
            pct = float(upd.get("percent"))
            if activity in df.columns:
                mask = df["_excel_row"] == excel_row
                if mask.any():
                    df.loc[mask, activity] = from_percent(pct, phase)
        except Exception:
            pass
    return df

def phase_editor_df(path, phase):
    """
    Devuelve la misma tabla que se usa en "Fases completas".
    Todas las partidas se muestran en 0-100 y el avance real se recalcula
    directamente desde las partidas.
    """
    df = load_phase(path, phase).copy()
    id_cols = ["Fase", "Piso", "Torre", "Departamento", "_excel_row"]
    progress_col = "% Avance Real Depto"

    activity_cols = [
        c for c in df.columns
        if c not in id_cols and c != progress_col
    ]

    for col in activity_cols:
        vals = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        if phase_scale(phase) == 1.0:
            vals = vals * 100.0
        df[col] = vals.round(1)

    if activity_cols:
        df[progress_col] = df[activity_cols].mean(axis=1).round(1)
    else:
        df[progress_col] = 0.0

    return df

def phase_numeric_percent_df(path, phase):
    """
    FUENTE ÚNICA PARA DASHBOARD Y GRÁFICOS.
    Usa exactamente las partidas visibles en "Fases completas" y calcula
    el % Avance Real Depto directamente desde esas partidas, sin depender
    de las fórmulas guardadas/caché del Excel.
    """
    df = phase_editor_df(path, phase).copy()

    id_cols = ["Fase", "Piso", "Torre", "Departamento", "_excel_row"]
    progress_col = "% Avance Real Depto"

    activity_cols = [
        c for c in df.columns
        if c not in id_cols and c != progress_col
    ]

    # Convertir todas las partidas a números 0-100.
    for col in activity_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Recalcular avance real del departamento desde las partidas.
    if activity_cols:
        df[progress_col] = df[activity_cols].mean(axis=1).round(2)
    else:
        df[progress_col] = 0.0

    return df

def phase_summary(path, phase):
    """
    AVANCE OFICIAL DE FASE.
    Replica el criterio de la hoja RESUMEN de la planilla:
    - toma las partidas de Fases completas,
    - calcula el promedio de cada partida por piso,
    - considera los pisos 1 a 9 con el mismo peso,
    - luego obtiene el promedio general de todas las partidas.
    """
    df = phase_numeric_percent_df(path, phase)

    if df.empty or "Piso" not in df.columns:
        return 0.0

    skip = {
        "Fase", "Piso", "Torre", "Departamento",
        "% Avance Real Depto", "_excel_row"
    }
    activity_cols = [c for c in df.columns if c not in skip]

    if not activity_cols:
        return 0.0

    # La planilla RESUMEN contempla Piso 1 a Piso 9.
    official_floors = list(range(1, 10))
    activity_results = []

    for activity in activity_cols:
        floor_results = []

        for piso in official_floors:
            piso_rows = df[pd.to_numeric(df["Piso"], errors="coerce") == piso]

            if piso_rows.empty:
                floor_avg = 0.0
            else:
                vals = pd.to_numeric(
                    piso_rows[activity],
                    errors="coerce"
                ).fillna(0.0)
                floor_avg = float(vals.mean()) if len(vals) else 0.0

            floor_results.append(floor_avg)

        activity_results.append(
            sum(floor_results) / len(official_floors)
        )

    return float(round(
        sum(activity_results) / len(activity_results)
    )) if activity_results else 0.0

def phase_by_floor(path, phase):
    df = phase_numeric_percent_df(path, phase)
    if "Piso" not in df.columns or "% Avance Real Depto" not in df.columns:
        return pd.DataFrame(columns=["Piso","Avance (%)"])
    out = (
        df.groupby("Piso", dropna=True)["% Avance Real Depto"]
        .mean()
        .reset_index()
        .rename(columns={"% Avance Real Depto":"Avance (%)"})
    )
    out["Avance (%)"] = out["Avance (%)"].round(1)
    return out

def phase_by_tower(path, phase):
    df = phase_numeric_percent_df(path, phase)
    if "Torre" not in df.columns or "% Avance Real Depto" not in df.columns:
        return pd.DataFrame(columns=["Torre","Avance (%)"])
    out = (
        df.groupby("Torre", dropna=True)["% Avance Real Depto"]
        .mean()
        .reset_index()
        .rename(columns={"% Avance Real Depto":"Avance (%)"})
    )
    out["Avance (%)"] = out["Avance (%)"].round(1)
    return out

def phase_activity_summary(path, phase):
    df = phase_numeric_percent_df(path, phase)
    skip = {"Fase","Piso","Torre","Departamento","% Avance Real Depto","_excel_row"}
    rows=[]
    for col in df.columns:
        if col not in skip:
            vals = pd.to_numeric(df[col], errors="coerce")
            if vals.notna().any():
                rows.append({"Partida":col, "Avance (%)":round(float(vals.fillna(0).mean()),1)})
    return pd.DataFrame(rows).sort_values("Avance (%)", ascending=False)

def phase_activity_floor_matrix(path, phase):
    """Matriz de partidas x pisos, alimentada directamente por Fases completas."""
    df = phase_numeric_percent_df(path, phase)
    skip = {"Fase","Piso","Torre","Departamento","% Avance Real Depto","_excel_row"}
    activities = [c for c in df.columns if c not in skip]
    rows = []
    piso_num = pd.to_numeric(df.get("Piso"), errors="coerce")
    for activity in activities:
        rec = {"Fase": phase.replace("FASE", "FASE "), "Partida": activity}
        vals_all = []
        for piso in range(1,10):
            vals = pd.to_numeric(df.loc[piso_num == piso, activity], errors="coerce").fillna(0.0)
            avg = float(vals.mean()) if len(vals) else 0.0
            avg = round(avg, 1)
            rec[f"Piso {piso}"] = avg
            vals_all.append(avg)
        real = round(sum(vals_all)/9, 1) if vals_all else 0.0
        rec["% Avance Real"] = real
        if real >= 100:
            rec["Estado"] = "✅ Completado"
        elif real >= 50:
            rec["Estado"] = "🟡 En progreso"
        elif real > 0:
            rec["Estado"] = "🟠 En riesgo"
        else:
            rec["Estado"] = "⚪ No iniciado"
        rows.append(rec)
    return pd.DataFrame(rows)


# Ruta Crítica oficial definida por el usuario (imagen de referencia 24-08-2026).
# Cada etiqueta visible se enlaza con la columna real de la fase para conservar actualización automática.
CRITICAL_ROUTE_COMPONENTS = {
    "FASE1": [
        ("Trazado tabique/Yeso/auxiliares vent. Y puertas", "Trazado tabique/Yeso/auxiliares vent. Y puertas"),
        ("Grada Buque (9cm)", "Grada Buque (9cm)"),
        ("Rasgo ventana - terrazas", "Rasgo ventana - terrazas"),
        ("Yeso Muro", "Yeso Muro"),
        ("Yeso Cielo", "Yeso Cielo"),
        ("Sanitario vertical y ramal", "Sanitario vertical y ramal"),
        ("Tabique 1° Cara", "Tabique 1° Cara"),
        ("Tabique 2° cara", "Tabique 2° cara "),
    ],
    "FASE2": [
        ("Huinchas tabiques", "Huinchas tabiques"),
        ("Instalación Marco Puerta Acceso", "Instalación Marco Puerta Acceso"),
        ("Marcos Puertas Interiores", "Marcos Puertas Interiores"),
        ("Impermeabilizacion Baños y cocina", "Impermeabilizacion Baños"),
        ("Instalación de Receptaculo", "Instalacion de tina/Receptaculo"),
        ("Protección Tina/receptaculo", "Protección Tina/receptaculo"),
        ("Instalación Ventana", "Instalación Ventana"),
        ("Ceramica Pisos", "Ceramica Pisos"),
        ("Ceramica Muros", "Ceramica Muros"),
        ("Cerámica de Terrazas", "Cerámica de Terrazas"),
    ],
    "FASE3": [
        ("Instalación piernas de closet", "Instalación piernas de closet"),
        ("Instalación Cornisas", "Instalación Cornisas"),
        ("Puertas interiores", "Puertas interiores"),
        ("Puertas de acceso y pasillo", "Puertas de acceso y pasillo"),
        ("Cerradura Puertas", "Cerradura Puertas "),
        ("Cableado Eléctrico y CCDD", "Cableado Eléctrico y CCDD"),
        ("Empaste Muro Cielo", "Empaste Muro Cielo"),
        ("Instalación de Baranda de Cristal", "Instalación de Baranda de Crital"),
        ("Losalin Cielo Departamento", "Losalin Cielo Departamento "),
        ("Instalacion Muebles de cocina (Base)", "Instalacion de Muebles de cocina (Base)"),
        ("Instalacion Muebles de cocina (Aereo)", "Instalacion Muebles de cocina (Aereo)"),
        ("Instalación cubiertas cocina", "Instalación cubiertas cocina"),
    ],
    "FASE4": [
        ("Instalacion de Artefactos Sanitarios.", "Instalacion de Artefactos Sanitarios."),
        ("Instalación de Shower Door", "Instalación de Shower Door y Barra Cortina"),
        ("Kit de cocina (Horno-Encimera -Campana - Lavaplato)", "Kit de cocina (Horno-Encimera -Campana - Lavaplato)"),
        ("Griferia.", "Griferia."),
        ("Piso Flotante + cubrejunta + junquillo", "Piso Flotante + cubrejunta + junquillo"),
        ("2da Mano de Pintura", "2° mano de pintura. (puertas, , cielo de baño, )"),
        ("Instalación de Papel Mural", "Instalación Papel Mural"),
        ("Topes en piso y Accesorios", "Topes en piso y Accesorios"),
    ],
}

def critical_route_matrix(path, phase):
    """Matriz oficial de Ruta Crítica según selección indicada por el usuario."""
    df = phase_numeric_percent_df(path, phase)
    if df.empty:
        return pd.DataFrame()
    piso_num = pd.to_numeric(df.get("Piso"), errors="coerce")
    rows = []
    for visible_name, source_col in CRITICAL_ROUTE_COMPONENTS.get(phase, []):
        if source_col not in df.columns:
            continue
        rec = {"Fase": phase.replace("FASE", "FASE "), "Partida": visible_name}
        floor_vals = []
        for piso in range(1, 10):
            vals = pd.to_numeric(df.loc[piso_num == piso, source_col], errors="coerce").fillna(0.0)
            avg = round(float(vals.mean()) if len(vals) else 0.0, 1)
            rec[f"Piso {piso}"] = avg
            floor_vals.append(avg)
        real = round(sum(floor_vals) / 9, 1) if floor_vals else 0.0
        rec["% Avance Real"] = real
        if real >= 100:
            rec["Estado"] = "✅ Completado"
        elif real >= 50:
            rec["Estado"] = "🟡 En progreso"
        elif real > 0:
            rec["Estado"] = "🟠 En riesgo"
        else:
            rec["Estado"] = "⚪ No iniciado"
        rows.append(rec)
    return pd.DataFrame(rows)

def _pct_style(v):
    try:
        x = float(v)
    except Exception:
        return ""
    if x >= 99.999:
        return "background-color:#22c55e;color:#062b13;font-weight:800"
    if x >= 50:
        return "background-color:#fde047;color:#3d3000;font-weight:700"
    return "background-color:#fb923c;color:#421900;font-weight:700"

def style_percent_df(df, percent_columns=None):
    """Semáforo de 3 tonos: verde=100%, amarillo=50-99%, naranjo=0-49%."""
    if percent_columns is None:
        percent_columns = [c for c in df.columns if "%" in str(c) or str(c).lower().startswith("piso ")]
    percent_columns = [c for c in percent_columns if c in df.columns]
    styler = df.style
    if percent_columns:
        styler = styler.map(_pct_style, subset=percent_columns)
        styler = styler.format({c:"{:.1f}%" for c in percent_columns}, na_rep="")
    return styler

# Fechas oficiales tomadas de Libro2.xlsx enviado por el usuario.
GANTT_PLAN = {
    "FASE1": [("Piso 1","2026-06-01",30),("Piso 2","2026-07-01",30),("Piso 3","2026-08-01",31),("Piso 4","2026-09-01",31),("Piso 5","2026-10-01",30),("Piso 6","2026-11-01",31),("Piso 7","2026-12-01",30),("Piso 8","2027-01-01",31),("Piso 9","2027-02-01",31)],
    "FASE2": [("Piso 1","2026-10-02",39),("Piso 2","2026-11-02",31),("Piso 3","2026-12-02",30),("Piso 4","2027-01-02",31),("Piso 5","2027-02-01",30),("Piso 6","2027-03-04",31),("Piso 7","2027-04-04",31),("Piso 8","2027-05-02",28),("Piso 9","2027-06-02",31)],
    "FASE3": [("Piso 1","2026-11-02",31),("Piso 2","2026-12-02",30),("Piso 3","2027-01-02",31),("Piso 4","2027-02-01",30),("Piso 5","2027-03-04",31),("Piso 6","2027-04-04",31),("Piso 7","2027-05-02",28),("Piso 8","2027-06-02",31),("Piso 9","2027-07-02",30)],
    "FASE4": [("Piso 1","2026-12-02",30),("Piso 2","2027-01-02",31),("Piso 3","2027-02-01",30),("Piso 4","2027-03-04",31),("Piso 5","2027-04-04",31),("Piso 6","2027-05-02",28),("Piso 7","2027-06-02",31),("Piso 8","2027-07-02",30),("Piso 9","2027-08-02",31)],
}

def build_gantt_df(path, phases):
    rows=[]
    for ph in phases:
        pf = phase_by_floor(path, ph)
        floor_progress = {}
        if not pf.empty:
            for _, rr in pf.iterrows():
                try: floor_progress[int(float(rr["Piso"]))] = float(rr["Avance (%)"])
                except Exception: pass
        for floor_name, start_txt, days in GANTT_PLAN[ph]:
            piso_num = int(floor_name.split()[-1])
            start = pd.to_datetime(start_txt)
            rows.append({
                "Fase": ph.replace("FASE","Fase "), "Piso": floor_name,
                "Tarea": f'{ph.replace("FASE","Fase ")} · {floor_name}',
                "Inicio": start, "Término": start + pd.Timedelta(days=int(days)),
                "Días": int(days), "Progreso (%)": round(floor_progress.get(piso_num,0.0),1)
            })
    return pd.DataFrame(rows)

def recalc_department_progress(ws, row, phase):
    first_activity = 5
    last_activity = ws.max_column - 1
    vals = []
    for c in range(first_activity, last_activity + 1):
        v = ws.cell(row, c).value
        try:
            if v is not None and v != "":
                vals.append(float(v))
        except:
            pass
    result = sum(vals) / len(vals) if vals else 0
    ws.cell(row, ws.max_column).value = result

def save_department(path, phase, excel_row, activity_values):
    wb = load_workbook(path)
    ws = wb[phase]
    headers = [c.value for c in ws[1]]
    header_to_col = {str(h): i+1 for i,h in enumerate(headers) if h is not None}

    username = st.session_state.get("user_name", st.session_state.get("username", "Usuario"))
    for activity, pct in activity_values.items():
        col = header_to_col.get(activity)
        if col:
            ws.cell(excel_row, col).value = from_percent(pct, phase)
            db_upsert_phase_update(phase, excel_row, activity, pct, username)

    recalc_department_progress(ws, excel_row, phase)
    wb.save(path)
    wb.close()
    # Forzar refresco inmediato de todas las vistas que dependen del avance.
    st.cache_data.clear()

def save_full_phase(path, phase, edited_df, original_df):
    """
    Guarda solamente las celdas modificadas de una fase completa.
    Los porcentajes visibles 0–100 vuelven a la escala interna correcta del Excel.
    """
    wb = load_workbook(path)
    ws = wb[phase]

    id_cols = ["Fase", "Piso", "Torre", "Departamento", "_excel_row"]
    headers = [c.value for c in ws[1]]
    header_to_col = {str(h): i+1 for i,h in enumerate(headers) if h is not None}

    changed_rows = set()
    changed_cells = 0

    for idx in edited_df.index:
        excel_row = int(original_df.loc[idx, "_excel_row"])

        for col in edited_df.columns:
            if col in id_cols or col == "% Avance Real Depto":
                continue

            new_val = edited_df.loc[idx, col]
            old_val = original_df.loc[idx, col]

            # Compare numerically where possible.
            try:
                new_num = float(new_val)
            except:
                new_num = 0.0
            try:
                old_num = float(old_val)
            except:
                old_num = 0.0

            if abs(new_num - old_num) < 1e-9:
                continue

            excel_col = header_to_col.get(str(col))
            if not excel_col:
                continue

            ws.cell(excel_row, excel_col).value = from_percent(new_num, phase)
            username = st.session_state.get("user_name", st.session_state.get("username", "Usuario"))
            db_upsert_phase_update(phase, excel_row, col, new_num, username)
            changed_rows.add(excel_row)
            changed_cells += 1

    for excel_row in changed_rows:
        recalc_department_progress(ws, excel_row, phase)

    wb.save(path)
    wb.close()
    # Forzar refresco inmediato de Dashboard, gráficos, Gantt y editores.
    st.cache_data.clear()
    return changed_cells, len(changed_rows)


def _xlsx_col_number(cell_ref):
    letters = "".join(ch for ch in str(cell_ref) if ch.isalpha())
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n

def _xlsx_col_letter(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out

def build_formatted_export():
    """
    Exporta directamente desde la plantilla XLSX depurada de la aplicación.
    No reabre ni vuelve a guardar el libro con openpyxl: modifica únicamente
    los valores numéricos de las partidas dentro del paquete XLSX.
    Esto conserva el formato y evita que Excel tenga que reparar el archivo.
    """
    with open(BASE, "rb") as f:
        base_bytes = f.read()

    updates = db_phase_updates()
    source = BytesIO(base_bytes)
    output = BytesIO()

    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel_doc = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_pkg_rel = "http://schemas.openxmlformats.org/package/2006/relationships"

    with zipfile.ZipFile(source, "r") as zin:
        names = set(zin.namelist())

        # Shared strings para interpretar los encabezados de FASE1–FASE4.
        shared = []
        if "xl/sharedStrings.xml" in names:
            ss_root = ET.fromstring(zin.read("xl/sharedStrings.xml"))
            for si in ss_root.findall(f"{{{ns_main}}}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{ns_main}}}t")))

        workbook_root = ET.fromstring(zin.read("xl/workbook.xml"))
        rel_root = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))

        rel_targets = {}
        for rel in rel_root:
            rel_targets[rel.attrib.get("Id")] = rel.attrib.get("Target", "")

        sheet_paths = {}
        sheets_node = workbook_root.find(f"{{{ns_main}}}sheets")
        if sheets_node is not None:
            for sh in sheets_node:
                name = sh.attrib.get("name")
                rid = sh.attrib.get(f"{{{ns_rel_doc}}}id")
                target = rel_targets.get(rid, "")
                if target.startswith("/"):
                    path = target.lstrip("/")
                else:
                    path = os.path.normpath(os.path.join("xl", target)).replace("\\\\", "/")
                sheet_paths[name] = path

        # Preparar modificaciones por XML de hoja.
        per_sheet = {}
        for upd in updates:
            try:
                phase = str(upd.get("phase", "")).upper()
                if phase not in PHASES or phase not in sheet_paths:
                    continue
                per_sheet.setdefault(phase, []).append(upd)
            except Exception:
                continue

        replacements = {}

        for phase, phase_updates in per_sheet.items():
            path = sheet_paths[phase]
            if path not in names:
                continue

            root = ET.fromstring(zin.read(path))
            sheet_data = root.find(f"{{{ns_main}}}sheetData")
            if sheet_data is None:
                continue

            rows = {
                int(r.attrib.get("r", "0")): r
                for r in sheet_data.findall(f"{{{ns_main}}}row")
            }

            # Leer encabezados desde fila 1.
            header_to_col = {}
            header_row = rows.get(1)
            if header_row is not None:
                for c in header_row.findall(f"{{{ns_main}}}c"):
                    ref = c.attrib.get("r", "")
                    col_num = _xlsx_col_number(ref)
                    ctype = c.attrib.get("t")
                    value = ""
                    if ctype == "inlineStr":
                        is_node = c.find(f"{{{ns_main}}}is")
                        if is_node is not None:
                            value = "".join(t.text or "" for t in is_node.iter(f"{{{ns_main}}}t"))
                    else:
                        v = c.find(f"{{{ns_main}}}v")
                        if v is not None and v.text is not None:
                            if ctype == "s":
                                try:
                                    value = shared[int(v.text)]
                                except Exception:
                                    value = ""
                            else:
                                value = v.text
                    if value:
                        header_to_col[str(value)] = col_num

            for upd in phase_updates:
                try:
                    row_num = int(upd.get("excel_row"))
                    activity = str(upd.get("activity"))
                    pct = float(upd.get("percent"))
                    col_num = header_to_col.get(activity)
                    if not col_num:
                        continue

                    raw_value = from_percent(pct, phase)
                    row = rows.get(row_num)
                    if row is None:
                        continue

                    target_ref = f"{_xlsx_col_letter(col_num)}{row_num}"
                    cells = list(row.findall(f"{{{ns_main}}}c"))
                    cell = next((c for c in cells if c.attrib.get("r") == target_ref), None)

                    if cell is None:
                        cell = ET.Element(f"{{{ns_main}}}c", {"r": target_ref})
                        # Copiar estilo de una celda cercana de la misma fila.
                        nearest = None
                        nearest_dist = 10**9
                        for other in cells:
                            ref = other.attrib.get("r", "")
                            oc = _xlsx_col_number(ref)
                            if oc >= 5:
                                d = abs(oc - col_num)
                                if d < nearest_dist:
                                    nearest = other
                                    nearest_dist = d
                        if nearest is not None and "s" in nearest.attrib:
                            cell.set("s", nearest.attrib["s"])

                        inserted = False
                        for idx, other in enumerate(cells):
                            if _xlsx_col_number(other.attrib.get("r", "")) > col_num:
                                row.insert(idx, cell)
                                inserted = True
                                break
                        if not inserted:
                            row.append(cell)

                    # Convertir a número manteniendo el estilo.
                    cell.attrib.pop("t", None)
                    for child in list(cell):
                        if child.tag in {
                            f"{{{ns_main}}}f",
                            f"{{{ns_main}}}v",
                            f"{{{ns_main}}}is",
                        }:
                            cell.remove(child)
                    v = ET.SubElement(cell, f"{{{ns_main}}}v")
                    v.text = str(raw_value)

                except Exception:
                    continue

            replacements[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        # Forzar recálculo de fórmulas al abrir Excel.
        calc_pr = workbook_root.find(f"{{{ns_main}}}calcPr")
        if calc_pr is None:
            calc_pr = ET.SubElement(workbook_root, f"{{{ns_main}}}calcPr")
        calc_pr.set("calcMode", "auto")
        calc_pr.set("fullCalcOnLoad", "1")
        calc_pr.set("forceFullCalc", "1")
        replacements["xl/workbook.xml"] = ET.tostring(
            workbook_root, encoding="utf-8", xml_declaration=True
        )

        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = replacements.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)

    output.seek(0)
    return output.getvalue()


def summary_display(path):
    wb = load_workbook(path, read_only=True, data_only=False)
    ws = wb["RESUMEN"]
    data = [[c.value for c in row] for row in ws.iter_rows()]
    wb.close()
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data[1:], columns=data[0])

    pct_cols = [c for c in df.columns if c not in ["Fase","Partidas "]]
    def phase_name(v):
        s = str(v).strip().replace(" ","").upper()
        return s if s in PHASES else None

    for idx,row in df.iterrows():
        phase = phase_name(row.get("Fase"))
        if not phase:
            continue
        for col in pct_cols:
            try:
                v = row[col]
                if v is None or v == "":
                    df.at[idx,col] = ""
                else:
                    x=float(v)
                    if phase_scale(phase)==1.0:
                        x*=100
                    df.at[idx,col]=f"{x:.1f}%"
            except:
                pass
    return df

def normalize_route_pct(value):
    try:
        if value is None or value == "":
            return 0.0
        x = float(value)
        if 0 <= x <= 1:
            return x * 100.0
        if 0 <= x <= 100:
            return x
        return 0.0
    except Exception:
        return 0.0

@st.cache_data(show_spinner=False)
def route_critical_data(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    if "SEGUIMIENTO R.CRITICA" not in wb.sheetnames:
        wb.close()
        return pd.DataFrame()

    ws = wb["SEGUIMIENTO R.CRITICA"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(rows) <= 1:
        return pd.DataFrame()

    headers = [
        v if v is not None else f"Columna {i+1}"
        for i, v in enumerate(rows[0])
    ]
    df = pd.DataFrame(rows[1:], columns=headers)

    if "Piso" in df.columns:
        df = df[df["Piso"].notna()].copy()

    id_cols = ["Columna 1", "Piso", "Torre", "Departamento"]
    old_progress = "Av. Depto/Piso"
    activity_cols = [c for c in df.columns if c not in id_cols and c != old_progress]

    for col in activity_cols:
        df[col] = df[col].map(normalize_route_pct)

    if activity_cols:
        df["Avance Ruta Crítica"] = df[activity_cols].mean(axis=1).round(1)
    else:
        df["Avance Ruta Crítica"] = 0.0

    return df

def route_detail_only(path):
    df = route_critical_data(path).copy()
    if df.empty:
        return df
    if "Departamento" in df.columns:
        df = df[df["Departamento"].astype(str).str.strip().str.lower() != "todos"].copy()
    return df

def route_overall(path):
    df = route_detail_only(path)
    if df.empty or "Avance Ruta Crítica" not in df.columns:
        return 0.0
    vals = pd.to_numeric(df["Avance Ruta Crítica"], errors="coerce").dropna()
    return round(float(vals.mean()), 1) if len(vals) else 0.0

def route_by_floor(path):
    df = route_detail_only(path)
    if df.empty or "Piso" not in df.columns:
        return pd.DataFrame()
    out = df.groupby("Piso", dropna=True)["Avance Ruta Crítica"].mean().reset_index()
    out = out.rename(columns={"Avance Ruta Crítica": "Avance (%)"})
    out["Avance (%)"] = out["Avance (%)"].round(1)
    return out

def route_by_tower(path):
    df = route_detail_only(path)
    if df.empty or "Torre" not in df.columns:
        return pd.DataFrame()
    out = df.groupby("Torre", dropna=True)["Avance Ruta Crítica"].mean().reset_index()
    out = out.rename(columns={"Avance Ruta Crítica": "Avance (%)"})
    out["Avance (%)"] = out["Avance (%)"].round(1)
    return out

def route_by_activity(path):
    df = route_detail_only(path)
    if df.empty:
        return pd.DataFrame()
    skip = {"Columna 1","Piso","Torre","Departamento","Av. Depto/Piso","Avance Ruta Crítica"}
    rows = []
    for col in df.columns:
        if col in skip:
            continue
        vals = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(vals):
            rows.append({"Partida": col, "Avance (%)": round(float(vals.mean()), 1)})
    return pd.DataFrame(rows).sort_values("Avance (%)", ascending=False)

def route_display(path, torre="Todas", piso="Todos"):
    df = route_detail_only(path).copy()
    if df.empty:
        return df
    if torre != "Todas" and "Torre" in df.columns:
        df = df[df["Torre"].astype(str) == str(torre)]
    if piso != "Todos" and "Piso" in df.columns:
        df = df[df["Piso"] == piso]

    drop_cols = [c for c in ["Columna 1", "Av. Depto/Piso"] if c in df.columns]
    if drop_cols:
        df = df.drop(columns=drop_cols)

    id_cols = ["Piso","Torre","Departamento"]
    for col in df.columns:
        if col not in id_cols:
            vals = pd.to_numeric(df[col], errors="coerce")
            df[col] = vals.map(lambda x: f"{x:.1f}%" if pd.notna(x) else "")
    return df

def ensure_history_sheet(path):
    wb = load_workbook(path)
    if "HISTORIAL_SEMANAL" not in wb.sheetnames:
        ws = wb.create_sheet("HISTORIAL_SEMANAL")
        ws.append([
            "Fecha actualización","Avance General","Fase 1","Fase 2","Fase 3","Fase 4","Usuario"
        ])
    wb.save(path)
    wb.close()

def register_weekly_snapshot(path, summaries, general, username, force=False):
    now = datetime.now()
    if now.weekday() != 4 and not force:
        return False, "La actualización semanal corresponde a los viernes."

    if db_ready():
        ok, msg = db_register_weekly(summaries, general, username)
        if not ok:
            return False, msg
    else:
        msg = "Base online no configurada; se guardará solo en el Excel de esta sesión."

    ensure_history_sheet(path)
    date_value = now.strftime("%Y-%m-%d")
    wb = load_workbook(path)
    ws = wb["HISTORIAL_SEMANAL"]
    target_row = None
    for r in range(2, ws.max_row + 1):
        existing = ws.cell(r, 1).value
        if existing is not None and str(existing)[:10] == date_value:
            target_row = r
            break
    vals = [
        date_value,float(general),float(summaries["FASE1"]),float(summaries["FASE2"]),
        float(summaries["FASE3"]),float(summaries["FASE4"]),username
    ]
    if target_row:
        for c,v in enumerate(vals, start=1):
            ws.cell(target_row,c).value = v
    else:
        ws.append(vals)
    wb.save(path)
    wb.close()
    return True, msg

@st.cache_data(show_spinner=False, ttl=5)
def load_weekly_history(path):
    rows = db_weekly_history()
    if rows:
        df = pd.DataFrame(rows).rename(columns={
            "update_date":"Fecha actualización",
            "general":"Avance General",
            "fase1":"Fase 1",
            "fase2":"Fase 2",
            "fase3":"Fase 3",
            "fase4":"Fase 4",
            "updated_by":"Usuario",
        })
        df["Fecha actualización"] = pd.to_datetime(df["Fecha actualización"], errors="coerce")
        return df.sort_values("Fecha actualización")

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if "HISTORIAL_SEMANAL" not in wb.sheetnames:
            wb.close()
            return pd.DataFrame()
        ws = wb["HISTORIAL_SEMANAL"]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if len(rows) <= 1:
            return pd.DataFrame()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df["Fecha actualización"] = pd.to_datetime(df["Fecha actualización"], errors="coerce")
        return df.sort_values("Fecha actualización")
    except Exception:
        return pd.DataFrame()

def next_friday_text():
    today = datetime.now().date()
    days = (4 - today.weekday()) % 7
    if days == 0:
        return "Hoy corresponde actualización semanal."
    nxt = today + pd.Timedelta(days=days)
    return f"Próxima actualización: viernes {nxt.strftime('%d-%m-%Y')}."

def raw_sheet_df(path, sheet_name):
    wb = load_workbook(path, read_only=True, data_only=False)
    ws = wb[sheet_name]
    data = [[c.value for c in row] for row in ws.iter_rows()]
    wb.close()
    width=max((len(r) for r in data), default=1)
    data=[r+[None]*(width-len(r)) for r in data]
    return pd.DataFrame(data, columns=[f"Columna {i+1}" for i in range(width)])


def current_role():
    return st.session_state.get("user_role", "Usuario")

def current_phase():
    return str(st.session_state.get("user_phase", "ALL")).upper()

def allowed_phases():
    role = current_role()
    phase = current_phase()
    if role in ["Administrador", "Visor"]:
        return PHASES
    if role == "Editor" and phase in PHASES:
        return [phase]
    return []

def can_edit_phase(phase):
    return current_role() == "Administrador" or (
        current_role() == "Editor" and current_phase() == phase
    )

def can_access_admin_tools():
    return current_role() == "Administrador"

with st.sidebar:
    if os.path.exists(LOGO):
        st.image(LOGO, use_container_width=True)

    st.markdown("## CONTROL DE OBRA")
    st.caption("EDIFICIO SAN FRANCISCO 211 · PAZ")
    st.markdown("---")

    role = current_role()
    assigned_phase = current_phase()

    if role == "Administrador":
        menu_options = [
            "📊 Dashboard",
            "📆 Comparación semanal",
            "📈 Gráficos por fase",
            "📅 Carta Gantt",
            "🧱 Actualizar avances",
            "🏢 Fases completas",
            "⚠️ Ruta crítica",
            "⬆️ Importar / Exportar",
        ]
    elif role == "Editor":
        menu_options = [
            "📊 Dashboard",
            "📈 Gráficos por fase",
            "📅 Carta Gantt",
            "🧱 Actualizar avances",
            "🏢 Fases completas",
        ]
    elif role == "Visor":
        menu_options = [
            "📊 Dashboard",
            "📆 Comparación semanal",
            "📈 Gráficos por fase",
            "📅 Carta Gantt",
            "🏢 Fases completas",
            "⚠️ Ruta crítica",
        ]
    else:
        menu_options = ["📊 Dashboard"]

    page = st.radio("Menú", menu_options, label_visibility="collapsed")

    st.markdown("---")
    st.caption(f"👤 {st.session_state.get('user_name','Usuario')}")
    st.caption(f"Rol: {role}")
    st.caption("Acceso: Todas las fases" if assigned_phase == "ALL"
               else f"Acceso: {assigned_phase.replace('FASE','Fase ')}")

    ok_db, db_msg = supabase_status()
    st.caption("🟢 Base online conectada" if ok_db else "🟠 Base online con error")

    if st.button("🚪 Cerrar sesión", use_container_width=True):
        for k in ["authenticated","username","user_name","user_role","user_phase"]:
            st.session_state.pop(k, None)
        st.rerun()

hlogo, htitle = st.columns([1.1,4.9], vertical_alignment="center")
with hlogo:
    if os.path.exists(LOGO):
        st.image(LOGO, use_container_width=True)
with htitle:
    st.markdown('<div class="sf-title">CONTROL FASES SAN FRANCISCO 211</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="sf-sub">Panel de control profesional · {datetime.now().strftime("%d-%m-%Y %H:%M")}</div>',
    unsafe_allow_html=True,
)

sync_a, sync_b = st.columns([4,1])
with sync_b:
    if st.button("🔄 Sincronizar datos", use_container_width=True, key="global_sync"):
        st.cache_data.clear()
        st.rerun()
with sync_a:
    online_now = db_phase_updates()
    if online_now:
        stamps = [str(x.get("updated_at","")) for x in online_now if x.get("updated_at")]
        last_online = max(stamps) if stamps else ""
        st.caption(f"🟢 Base online sincronizada · última modificación: {last_online[:19].replace('T',' ')}")
    else:
        st.caption("🟢 Base online conectada · sin cambios detallados registrados")

# Avance calculado desde el detalle actual (partidas/departamentos).
live_summaries = {p: phase_summary(st.session_state.workbook_path, p) for p in PHASES}
live_general = round(sum(live_summaries.values()) / 4, 1)

# AVANCE OFICIAL VISIBLE: lectura DIRECTA del último registro de weekly_history.
# No usa cache para evitar que un computador muestre un corte anterior.
summaries = dict(live_summaries)
general = live_general
official_snapshot_date = None
official_snapshot_user = None
official_snapshot_source = "avance en vivo"
_latest_official = db_latest_weekly_snapshot()
if _latest_official:
    try:
        summaries = {
            "FASE1": float(pd.to_numeric(_latest_official.get("fase1"), errors="coerce") or 0),
            "FASE2": float(pd.to_numeric(_latest_official.get("fase2"), errors="coerce") or 0),
            "FASE3": float(pd.to_numeric(_latest_official.get("fase3"), errors="coerce") or 0),
            "FASE4": float(pd.to_numeric(_latest_official.get("fase4"), errors="coerce") or 0),
        }
        general = float(pd.to_numeric(_latest_official.get("general"), errors="coerce") or 0)
        official_snapshot_date = pd.to_datetime(_latest_official.get("update_date"), errors="coerce")
        official_snapshot_user = _latest_official.get("updated_by")
        official_snapshot_source = "weekly_history · Supabase"
    except Exception:
        summaries = dict(live_summaries)
        general = live_general

# AVANCE GENERAL SIEMPRE VISIBLE EN TODAS LAS PÁGINAS
st.markdown('<div class="section">AVANCE GENERAL DEL PROYECTO</div>', unsafe_allow_html=True)
gcol, p1col, p2col, p3col, p4col = st.columns([1.15,1,1,1,1])

with gcol:
    st.markdown(
        f'<div class="general-box"><div class="general-label">Avance General</div>'
        f'<div class="general-pct">{general:.1f}%</div></div>',
        unsafe_allow_html=True
    )
    st.progress(min(max(general/100,0),1))

for col, phase in zip([p1col,p2col,p3col,p4col], PHASES):
    val = summaries[phase]
    with col:
        st.markdown(
            f'<div class="phase-card"><div class="phase-name">{phase.replace("FASE","Fase ")}</div>'
            f'<div class="phase-pct">{val:.1f}%</div></div>',
            unsafe_allow_html=True
        )
        st.progress(min(max(val/100,0),1))

if official_snapshot_date is not None and not pd.isna(official_snapshot_date):
    _who = f" · registrado por {official_snapshot_user}" if official_snapshot_user else ""
    st.caption(
        f"✅ Corte oficial usado: {official_snapshot_date.strftime('%d-%m-%Y')} · "
        f"Fuente: {official_snapshot_source}{_who}. "
        f"General {general:.1f}% · F1 {summaries['FASE1']:.1f}% · "
        f"F2 {summaries['FASE2']:.1f}% · F3 {summaries['FASE3']:.1f}% · F4 {summaries['FASE4']:.1f}%."
    )
else:
    st.warning("No se encontró un corte oficial en weekly_history; se muestra temporalmente el avance calculado desde las partidas actuales.")
st.divider()

if page == "📊 Dashboard":
    ok_db, db_msg = supabase_status()
    if ok_db:
        st.success("🟢 Supabase conectado correctamente.")
    else:
        st.warning("🟠 Supabase aún no está conectado.")
        with st.expander("Diagnóstico de conexión"):
            st.code(db_msg)
            st.markdown(
                "Formato esperado en Streamlit Secrets:\n\n"
                "[supabase]\n"
                'url = "https://TU-PROYECTO.supabase.co"\n'
                'service_role_key = "sb_secret_..."'
            )

    if official_snapshot_date is not None and not pd.isna(official_snapshot_date):
        st.info(
            f"📌 Dashboard sincronizado con el último corte oficial de Comparación semanal: "
            f"{official_snapshot_date.strftime('%d-%m-%Y')} · "
            f"General {general:.1f}% · Fase 1 {summaries['FASE1']:.1f}% · "
            f"Fase 2 {summaries['FASE2']:.1f}% · Fase 3 {summaries['FASE3']:.1f}% · Fase 4 {summaries['FASE4']:.1f}%."
        )
    else:
        st.warning("Dashboard sin corte oficial disponible; se está mostrando el avance en vivo.")
    c1,c2,c3,c4,c5=st.columns(5)
    cards=[
        (c1,"AVANCE GENERAL",f"{general:.1f}%","blue"),
        (c2,"FASE 1",f"{summaries['FASE1']:.0f}%","good"),
        (c3,"FASE 2",f"{summaries['FASE2']:.0f}%","warn"),
        (c4,"FASE 3",f"{summaries['FASE3']:.0f}%","warn"),
        (c5,"FASE 4",f"{summaries['FASE4']:.0f}%","bad"),
    ]
    for c,lab,val,cl in cards:
        c.markdown(f'<div class="card"><div class="klabel">{lab}</div><div class="kvalue {cl}">{val}</div></div>',unsafe_allow_html=True)

    st.write("")
    left,right=st.columns([1.25,1])
    with left:
        st.markdown('<div class="section">AVANCE GENERAL POR FASE · DESDE FASES COMPLETAS</div>',unsafe_allow_html=True)
        chart=pd.DataFrame({
            "Fase":["Fase 1","Fase 2","Fase 3","Fase 4"],
            "Avance (%)":[summaries[p] for p in PHASES],
        })
        st.bar_chart(chart.set_index("Fase"),height=330)
        compare = chart.copy()
        compare["Avance"] = compare["Avance (%)"].map(lambda x: f"{x:.1f}%")
        compare_num = chart[["Fase","Avance (%)"]].copy()
        st.dataframe(style_percent_df(compare_num, ["Avance (%)"]), use_container_width=True, hide_index=True)
    with right:
        st.markdown('<div class="section">ESTADO ACTUAL</div>',unsafe_allow_html=True)
        for p in PHASES:
            v=summaries[p]
            st.write(f"**{p.replace('FASE','Fase ')} — {v:.1f}%**")
            st.progress(min(max(v/100,0),1))

    st.markdown('<div class="section">AVANCE POR PISO · DESDE FASES COMPLETAS</div>',unsafe_allow_html=True)
    selected_phase=st.selectbox("Fase para detalle",allowed_phases(),format_func=lambda x:x.replace("FASE","Fase "))
    piso=phase_by_floor(st.session_state.workbook_path,selected_phase)
    if not piso.empty:
        st.bar_chart(piso.set_index("Piso"),height=330)

elif page == "📆 Comparación semanal":
    st.markdown('<div class="section">COMPARACIÓN DE AVANCE · REGISTRO DE LOS VIERNES</div>', unsafe_allow_html=True)

    st.info(
        "Cada punto del gráfico corresponde al avance registrado en un viernes. "
        "Se compara el Avance General y las cuatro fases a través del tiempo."
    )

    is_friday = datetime.now().weekday() == 4
    role = st.session_state.get("user_role", "Usuario")
    user = st.session_state.get(
        "user_name",
        st.session_state.get("username", "Usuario")
    )

    top1, top2 = st.columns([2,1])

    with top1:
        if is_friday:
            st.success("✅ Hoy es viernes: corresponde registrar el avance semanal.")
        else:
            st.info("📅 " + next_friday_text())

    with top2:
        if current_role() == "Administrador" and st.button(
            "📌 REGISTRAR AVANCE DE HOY",
            type="primary",
            use_container_width=True
        ):
            ok, msg = register_weekly_snapshot(
                st.session_state.workbook_path,
                live_summaries,
                live_general,
                user,
                force=(role == "Administrador")
            )
            load_weekly_history.clear()

            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.warning(msg)

    if not is_friday and role == "Administrador":
        st.caption(
            "Administrador: puedes registrar manualmente fuera del viernes "
            "para regularizar o probar el historial."
        )

    # Mostrar siempre el avance ACTUAL EN VIVO, aunque el historial semanal no se haya registrado todavía.
    st.markdown("### Avance actual en vivo")
    a1,a2,a3,a4,a5 = st.columns(5)
    a1.metric("Avance General", f"{live_general:.1f}%")
    a2.metric("Fase 1", f"{live_summaries['FASE1']:.0f}%")
    a3.metric("Fase 2", f"{live_summaries['FASE2']:.0f}%")
    a4.metric("Fase 3", f"{live_summaries['FASE3']:.0f}%")
    a5.metric("Fase 4", f"{live_summaries['FASE4']:.0f}%")
    st.caption("Estos valores cambian inmediatamente al guardar avances. El bloque siguiente corresponde al historial registrado de los viernes.")

    hist = load_weekly_history(st.session_state.workbook_path)

    if hist.empty:
        st.warning(
            "Todavía no existen viernes registrados. "
            "Presiona «Registrar avance de hoy» para crear el primer punto del historial."
        )
    else:
        hist = hist.copy()
        hist["Fecha actualización"] = pd.to_datetime(
            hist["Fecha actualización"],
            errors="coerce"
        )
        hist = hist.dropna(subset=["Fecha actualización"]).sort_values("Fecha actualización")

        metric_cols = ["Avance General","Fase 1","Fase 2","Fase 3","Fase 4"]
        for c in metric_cols:
            hist[c] = pd.to_numeric(hist[c], errors="coerce")

        latest = hist.iloc[-1]

        st.markdown("### Último viernes registrado · histórico")
        k1,k2,k3,k4,k5 = st.columns(5)
        k1.metric("Avance General", f"{latest['Avance General']:.1f}%")
        k2.metric("Fase 1", f"{latest['Fase 1']:.0f}%")
        k3.metric("Fase 2", f"{latest['Fase 2']:.0f}%")
        k4.metric("Fase 3", f"{latest['Fase 3']:.0f}%")
        k5.metric("Fase 4", f"{latest['Fase 4']:.0f}%")

        chart_df = hist[["Fecha actualización"] + metric_cols].copy()
        chart_df["Viernes"] = chart_df["Fecha actualización"].dt.strftime("%d-%m-%Y")
        chart_df = chart_df.set_index("Viernes")[metric_cols]

        st.markdown("### Evolución semanal del proyecto")
        st.line_chart(chart_df, height=400)

        st.caption(
            "Eje horizontal: fecha de actualización de cada viernes. "
            "Eje vertical: porcentaje de avance 0%–100%."
        )

        if len(hist) >= 2:
            current = hist.iloc[-1]
            previous = hist.iloc[-2]

            st.markdown("### Comparación con el viernes anterior")

            comparison = pd.DataFrame({
                "Indicador": metric_cols,
                "Viernes anterior": [previous[c] for c in metric_cols],
                "Último viernes": [current[c] for c in metric_cols],
            })
            comparison["Variación (pp)"] = (
                comparison["Último viernes"] -
                comparison["Viernes anterior"]
            ).round(1)

            comp_styler = style_percent_df(comparison, ["Viernes anterior","Último viernes"])
            comp_styler = comp_styler.format({"Variación (pp)":"{:+.1f}"})
            st.dataframe(comp_styler, use_container_width=True, hide_index=True)

        st.markdown("### Historial de viernes registrados")

        shown = hist.copy()
        shown["Fecha actualización"] = shown["Fecha actualización"].dt.strftime("%d-%m-%Y")

        for c in metric_cols:
            shown[c] = shown[c].map(
                lambda x: f"{x:.1f}%" if pd.notna(x) else ""
            )

        visible_cols = [
            "Fecha actualización",
            "Avance General",
            "Fase 1",
            "Fase 2",
            "Fase 3",
            "Fase 4",
            "Usuario"
        ]
        visible_cols = [c for c in visible_cols if c in shown.columns]

        st.dataframe(
            shown[visible_cols],
            use_container_width=True,
            hide_index=True
        )

        st.success(
            f"Historial disponible: {len(hist)} actualización(es) registrada(s)."
        )

elif page == "📈 Gráficos por fase":
    st.markdown('<div class="section">GRÁFICOS DE CADA FASE · FUENTE: FASES COMPLETAS</div>',unsafe_allow_html=True)
    st.caption("Todos los gráficos se generan directamente desde los datos de «Fases completas» y trabajan en escala 0%–100%.")

    for phase in allowed_phases():
        st.markdown(f"## {phase.replace('FASE','Fase ')}")
        st.metric("AVANCE GENERAL DE LA FASE", f"{summaries[phase]:.0f}%")
        st.progress(min(max(summaries[phase]/100,0),1))
        a,b=st.columns(2)

        with a:
            st.markdown("**Avance por piso**")
            pf=phase_by_floor(st.session_state.workbook_path,phase)
            if not pf.empty:
                st.bar_chart(pf.set_index("Piso"),height=300)

        with b:
            st.markdown("**Avance por torre**")
            tw=phase_by_tower(st.session_state.workbook_path,phase)
            if not tw.empty:
                st.bar_chart(tw.set_index("Torre"),height=300)

        st.markdown("**Avance promedio por partida**")
        act=phase_activity_summary(st.session_state.workbook_path,phase)
        if not act.empty:
            st.bar_chart(act.set_index("Partida"),height=400)
        st.divider()

elif page == "🧱 Actualizar avances":
    if current_role() == "Visor":
        st.error("Tu perfil es solo de lectura.")
        st.stop()
    st.markdown('<div class="section">ACTUALIZAR AVANCE DE DEPARTAMENTO</div>',unsafe_allow_html=True)
    st.success("Todos los controles de avance se muestran y editan de 0% a 100%.")

    phase=st.selectbox("Fase",allowed_phases(),format_func=lambda x:x.replace("FASE","Fase "))
    df=load_phase(st.session_state.workbook_path,phase)

    towers=sorted([str(x) for x in df["Torre"].dropna().unique()])
    torre=st.selectbox("Torre",towers)
    f1=df[df["Torre"].astype(str)==torre]

    pisos=sorted(f1["Piso"].dropna().unique(),key=lambda x:float(x) if str(x).replace(".","",1).isdigit() else str(x))
    piso=st.selectbox("Piso",pisos)
    f2=f1[f1["Piso"]==piso]

    deps=sorted(f2["Departamento"].dropna().unique(),key=lambda x:str(x))
    depto=st.selectbox("Departamento",deps)
    rec=f2[f2["Departamento"]==depto].iloc[0]
    excel_row=int(rec["_excel_row"])

    progress_col="% Avance Real Depto"
    current=to_percent(rec[progress_col],phase)
    st.metric("Avance actual del departamento",f"{current:.1f}%")
    st.progress(min(max(current/100,0),1))

    skip=["Fase","Piso","Torre","Departamento",progress_col,"_excel_row"]
    activities=[c for c in df.columns if c not in skip]

    st.markdown('<div class="section">PARTIDAS</div>',unsafe_allow_html=True)
    edits={}
    cols=st.columns(2)
    for i,activity in enumerate(activities):
        val=to_percent(rec[activity],phase)
        with cols[i%2]:
            edits[activity]=st.number_input(
                f"{activity} (%)",
                min_value=0.0,max_value=100.0,
                value=float(round(val,1)),step=5.0,format="%.1f",
                key=f"{phase}-{excel_row}-{i}"
            )

    calculated=round(sum(edits.values())/len(edits),1) if edits else 0.0
    st.markdown(f"### Nuevo avance calculado: **{calculated:.1f}%**")
    st.progress(min(max(calculated/100,0),1))

    if not can_edit_phase(phase):
        st.warning("No tienes permiso para modificar esta fase.")
    elif st.button("💾 GUARDAR ACTUALIZACIÓN",type="primary",use_container_width=True):
        save_department(st.session_state.workbook_path,phase,excel_row,edits)
        st.success(f"Departamento {depto} actualizado correctamente. Dashboard, gráficos y Gantt se refrescarán con este nuevo avance.")
        st.rerun()

elif page == "🏢 Fases completas":
    st.markdown('<div class="section">EDITAR FASE COMPLETA</div>', unsafe_allow_html=True)
    st.success("Ahora puedes editar directamente todas las partidas de cada Fase. Los valores se trabajan de 0% a 100%.")

    phase = st.selectbox(
        "Selecciona una fase",
        allowed_phases(),
        format_func=lambda x:x.replace("FASE","Fase ")
    )

    original = phase_editor_df(st.session_state.workbook_path, phase)

    # No mostrar _excel_row en pantalla.
    visible_cols = [c for c in original.columns if c != "_excel_row"]
    editor_source = original[visible_cols].copy()

    # Proteger columnas identificadoras y el cálculo total.
    disabled_cols = ["Fase", "Piso", "Torre", "Departamento", "% Avance Real Depto"]

    st.caption(
        "Puedes modificar las partidas. Las columnas Fase, Piso, Torre, Departamento y "
        "% Avance Real Depto quedan protegidas; el avance total se recalcula automáticamente."
    )

    column_cfg = {}
    for c in visible_cols:
        if c not in ["Fase", "Piso", "Torre", "Departamento"]:
            column_cfg[c] = st.column_config.NumberColumn(
                c,
                min_value=0.0,
                max_value=100.0,
                step=5.0,
                format="%.1f%%"
            )

    edited = st.data_editor(
        editor_source,
        use_container_width=True,
        height=700,
        num_rows="fixed",
        disabled=disabled_cols,
        column_config=column_cfg,
        key=f"full_editor_{phase}"
    )

    with st.expander("🎨 Vista semáforo de porcentajes", expanded=False):
        preview_cols = [c for c in visible_cols if c not in ["Fase","Piso","Torre","Departamento"]]
        st.caption("Verde = 100% · Amarillo = 50% a 99% · Naranjo = 0% a 49%")
        st.dataframe(
            style_percent_df(editor_source[visible_cols], preview_cols),
            use_container_width=True, height=520, hide_index=True
        )

    csave, cinfo = st.columns([1,2])
    with csave:
        if not can_edit_phase(phase):
            st.info("Modo solo lectura: no tienes permiso para modificar esta fase.")
        elif st.button("💾 GUARDAR CAMBIOS DE LA FASE", type="primary", use_container_width=True):
            edited_full = original.copy()
            for c in visible_cols:
                edited_full[c] = edited[c]

            changed_cells, changed_rows = save_full_phase(
                st.session_state.workbook_path,
                phase,
                edited_full,
                original
            )
            st.success(
                f"Guardado correctamente: {changed_cells} celdas modificadas "
                f"en {changed_rows} departamentos. Dashboard, gráficos y Gantt usarán estos mismos datos inmediatamente."
            )
            st.rerun()

    with cinfo:
        st.info(
            "Al guardar, la aplicación recalcula automáticamente el % Avance Real Depto "
            "de los departamentos modificados y el Dashboard se actualiza."
        )

elif page == "📋 Resumen":
    st.caption("El RESUMEN se convierte automáticamente a porcentaje según la fase.")
    st.dataframe(summary_display(st.session_state.workbook_path),use_container_width=True,height=700)

elif page == "📅 Avance semanal":
    names=get_sheet_names(st.session_state.workbook_path)
    opts=[n for n in names if "AVANCE SEMANAL" in n.upper()]
    sh=opts[0] if opts else names[0]
    st.dataframe(raw_sheet_df(st.session_state.workbook_path,sh),use_container_width=True,height=700)

elif page == "📅 Carta Gantt":
    st.markdown('<div class="section">CARTA GANTT · PROGRAMA OFICIAL SAN FRANCISCO 211</div>', unsafe_allow_html=True)
    st.caption("Fechas verificadas contra Libro2.xlsx: inicio del proyecto 01-06-2026. El progreso se actualiza desde la base online.")

    phase_filter = st.selectbox(
        "Mostrar", ["Todas las fases"] + [p.replace("FASE","Fase ") for p in allowed_phases()],
        key="gantt_phase_filter"
    )
    selected = allowed_phases() if phase_filter == "Todas las fases" else [phase_filter.upper().replace("FASE ","FASE")]
    gdf = build_gantt_df(st.session_state.workbook_path, selected)

    if gdf.empty:
        st.warning("No hay datos disponibles para construir la Carta Gantt.")
    else:
        fig = px.timeline(
            gdf, x_start="Inicio", x_end="Término", y="Tarea", color="Fase", text="Progreso (%)",
            hover_data={"Piso":True,"Días":True,"Progreso (%)":':.1f',"Inicio":'|%d-%m-%Y',"Término":'|%d-%m-%Y'}
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="inside")
        fig.update_yaxes(autorange="reversed", title=None)
        fig.update_xaxes(title="Programación oficial", tickformat="%d-%m-%Y")
        fig.update_layout(height=max(520,30*len(gdf)+170), legend_title_text="Fase", margin=dict(l=10,r=10,t=30,b=10))
        st.plotly_chart(fig, use_container_width=True)

        detail = gdf[["Fase","Piso","Inicio","Término","Días","Progreso (%)"]].copy()
        detail["Inicio"] = detail["Inicio"].dt.strftime("%d-%m-%Y")
        detail["Término"] = detail["Término"].dt.strftime("%d-%m-%Y")
        st.dataframe(style_percent_df(detail, ["Progreso (%)"]), use_container_width=True, hide_index=True, height=420)

elif page == "⚠️ Ruta crítica":
    st.markdown('<div class="section">RUTA CRÍTICA DEL PROYECTO · ACTUALIZACIÓN EN LÍNEA</div>', unsafe_allow_html=True)
    st.caption("Ruta Crítica oficial según la imagen adjunta: selección específica de partidas por fase, Piso 1 a Piso 9 y avance real. Se recalcula automáticamente desde Fases completas y Supabase.")

    phase_filter = st.selectbox(
        "Fase", ["Todas las fases"] + [p.replace("FASE","Fase ") for p in allowed_phases()],
        key="route_matrix_phase"
    )
    selected_phases = allowed_phases() if phase_filter == "Todas las fases" else [phase_filter.upper().replace("FASE ","FASE")]

    mats = {ph: critical_route_matrix(st.session_state.workbook_path, ph) for ph in selected_phases}
    valid_mats = {ph:m for ph,m in mats.items() if not m.empty}

    if not valid_mats:
        st.warning("No existen partidas disponibles para construir la Ruta Crítica.")
    else:
        # KPIs oficiales superiores: último corte registrado en Comparación semanal.
        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("AVANCE GENERAL OFICIAL", f"{general:.1f}%")
        c2.metric("FASE 1", f"{summaries['FASE1']:.0f}%")
        c3.metric("FASE 2", f"{summaries['FASE2']:.0f}%")
        c4.metric("FASE 3", f"{summaries['FASE3']:.0f}%")
        c5.metric("FASE 4", f"{summaries['FASE4']:.0f}%")
        if official_snapshot_date is not None:
            st.info(f"Corte oficial mostrado: {official_snapshot_date.strftime('%d-%m-%Y')}. Las partidas se leen del último detalle guardado en la base online.")

        st.markdown("### Ruta crítica oficial · partidas definidas por fase")
        pct_cols = [f"Piso {i}" for i in range(1,10)] + ["% Avance Real"]
        for ph in selected_phases:
            matrix = valid_mats.get(ph)
            if matrix is None or matrix.empty:
                continue
            phase_label = ph.replace("FASE", "FASE ")
            st.markdown(f"#### {phase_label} · {len(matrix)} partidas críticas")
            # Mostrar todas las filas de la fase sin recortar componentes.
            table_height = max(220, min(900, 38 * (len(matrix) + 1)))
            st.dataframe(
                style_percent_df(matrix, pct_cols),
                use_container_width=True, hide_index=True, height=table_height
            )

        st.caption("Semáforo: 🟩 100% completado · 🟨 50%–99% en progreso · 🟧 0%–49% pendiente/en riesgo. La Ruta Crítica muestra únicamente las partidas definidas en la imagen de referencia y se actualiza desde FASE1–FASE4.")

        st.markdown("### Carta Gantt vinculada a la Ruta Crítica")
        gdf = build_gantt_df(st.session_state.workbook_path, selected_phases)
        if not gdf.empty:
            fig = px.timeline(gdf, x_start="Inicio", x_end="Término", y="Tarea", color="Fase", text="Progreso (%)")
            fig.update_traces(texttemplate="%{text:.0f}%", textposition="inside")
            fig.update_yaxes(autorange="reversed", title=None)
            fig.update_xaxes(title="Fechas oficiales (Libro2.xlsx)", tickformat="%d-%m-%Y")
            fig.add_vline(x=pd.Timestamp.today().timestamp()*1000, line_dash="dash", line_color="red")
            fig.update_layout(height=max(500,27*len(gdf)+150), margin=dict(l=10,r=10,t=20,b=10))
            st.plotly_chart(fig, use_container_width=True)

elif page == "⬆️ Importar / Exportar":
    if not can_access_admin_tools():
        st.error("Solo el Administrador puede importar o exportar archivos.")
        st.stop()
    st.info(
        "La descarga usa siempre la plantilla Excel depurada de esta versión. "
        "AVANCE SEMANAL 1 queda enlazado a las partidas de FASE1–FASE4."
    )

    uploaded=st.file_uploader("Cargar una nueva versión de CONTROL_FASES_SFCO211.xlsx",type=["xlsx"])
    if uploaded is not None:
        with open(st.session_state.workbook_path,"wb") as f:
            f.write(uploaded.getbuffer())
        load_phase.clear(); get_sheet_names.clear()
        st.success("Archivo cargado.")
        st.rerun()

    st.markdown("### Respaldo de base online")
    online_updates = db_phase_updates()
    if online_updates:
        backup_df = pd.DataFrame(online_updates)
        preferred = [c for c in ["phase","excel_row","activity","percent","updated_by","updated_at","id"] if c in backup_df.columns]
        other = [c for c in backup_df.columns if c not in preferred]
        backup_df = backup_df[preferred + other].sort_values(
            [c for c in ["updated_at","phase","excel_row"] if c in backup_df.columns],
            ascending=True
        )
        c1,c2,c3 = st.columns(3)
        c1.metric("Cambios guardados online", f"{len(backup_df):,}".replace(",","."))
        c2.metric("Departamentos con cambios", backup_df[[c for c in ["phase","excel_row"] if c in backup_df.columns]].drop_duplicates().shape[0])
        latest = str(backup_df["updated_at"].max()) if "updated_at" in backup_df.columns else "—"
        c3.metric("Última actualización", latest[:19].replace("T"," "))
        st.success("La base online tiene detalle de avances. Descarga este respaldo antes de importar o reemplazar cualquier Excel.")
        st.download_button(
            "🛟 DESCARGAR RESPALDO BASE ONLINE (CSV)",
            data=backup_df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"RESPALDO_BASE_ONLINE_SFCO211_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
        with st.expander("Ver últimos cambios guardados online"):
            st.dataframe(backup_df.tail(100), use_container_width=True, height=420)
    else:
        st.warning("No se encontraron cambios detallados en phase_updates. No importes otro Excel hasta revisar la conexión de Supabase.")

    st.markdown("### Exportar Excel recuperado")
    st.caption(
        "Esta descarga parte de la plantilla oficial y aplica TODOS los cambios detallados que actualmente existen en Supabase. "
        "Úsala como respaldo recuperado del estado online."
    )

    export_data = build_formatted_export()

    st.download_button(
        "⬇️ Descargar Excel RECUPERADO desde base online",
        data=export_data,
        file_name="CONTROL_FASES_SFCO211_RECUPERADO_BASE_ONLINE.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if st.button("Restablecer archivo original"):
        shutil.copy2(BASE,st.session_state.workbook_path)
        load_phase.clear(); get_sheet_names.clear()
        st.rerun()


import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os, tempfile, shutil
from datetime import datetime
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

if "workbook_path" not in st.session_state:
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    shutil.copy2(BASE, p)
    st.session_state.workbook_path = p

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

@st.cache_data(show_spinner=False)
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
    load_phase.clear()
    get_sheet_names.clear()

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
    load_phase.clear()
    get_sheet_names.clear()
    return changed_cells, len(changed_rows)

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

@st.cache_data(show_spinner=False)
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
            "🧱 Actualizar avances",
            "🏢 Fases completas",
            "⚠️ Ruta crítica",
            "⬆️ Importar / Exportar",
        ]
    elif role == "Editor":
        menu_options = [
            "📊 Dashboard",
            "📈 Gráficos por fase",
            "🧱 Actualizar avances",
            "🏢 Fases completas",
        ]
    elif role == "Visor":
        menu_options = [
            "📊 Dashboard",
            "📆 Comparación semanal",
            "📈 Gráficos por fase",
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

summaries = {p:phase_summary(st.session_state.workbook_path,p) for p in PHASES}
general = round(sum(summaries.values()) / 4, 1)

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

st.caption("Los avances oficiales se calculan con el criterio de la hoja RESUMEN y se redondean a porcentaje entero, igual que en la planilla: Fase 1, Fase 2, Fase 3 y Fase 4.")
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

    st.info("Fuente de los indicadores y gráficos: partidas actuales de «Fases completas». El avance se recalcula automáticamente, sin depender de fórmulas almacenadas en Excel.")
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
        st.dataframe(compare[["Fase","Avance"]], use_container_width=True, hide_index=True)
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
                summaries,
                general,
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

        st.markdown("### Último viernes registrado")
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

            st.dataframe(
                comparison.style.format({
                    "Viernes anterior": "{:.1f}%",
                    "Último viernes": "{:.1f}%",
                    "Variación (pp)": "{:+.1f}",
                }),
                use_container_width=True,
                hide_index=True
            )

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
        st.success(f"Departamento {depto} actualizado correctamente.")
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
                f"en {changed_rows} departamentos. Los gráficos usarán estos mismos datos."
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

elif page == "⚠️ Ruta crítica":
    st.markdown('<div class="section">RUTA CRÍTICA · AVANCE REAL</div>', unsafe_allow_html=True)

    rdf = route_detail_only(st.session_state.workbook_path)
    roverall = route_overall(st.session_state.workbook_path)

    if rdf.empty:
        st.error("No se encontraron datos utilizables en SEGUIMIENTO R.CRITICA.")
    else:
        floors = rdf["Piso"].nunique() if "Piso" in rdf.columns else 0
        deps = rdf["Departamento"].nunique() if "Departamento" in rdf.columns else 0
        active = int((rdf["Avance Ruta Crítica"] > 0).sum())

        k1,k2,k3,k4 = st.columns(4)
        k1.metric("AVANCE RUTA CRÍTICA", f"{roverall:.1f}%")
        k2.metric("PISOS CONTROLADOS", str(floors))
        k3.metric("DEPARTAMENTOS", str(deps))
        k4.metric("DEP. CON AVANCE", str(active))

        st.progress(min(max(roverall/100,0),1))
        st.success("Ruta Crítica se calcula directamente desde las partidas de SEGUIMIENTO R.CRITICA.")

        left,right = st.columns(2)
        with left:
            st.markdown("### Avance por piso")
            pf = route_by_floor(st.session_state.workbook_path)
            if not pf.empty:
                st.bar_chart(pf.set_index("Piso"), height=400)

        with right:
            st.markdown("### Avance por torre")
            tw = route_by_tower(st.session_state.workbook_path)
            if not tw.empty:
                st.bar_chart(tw.set_index("Torre"), height=400)

        st.markdown("### Avance promedio por partida crítica")
        act = route_by_activity(st.session_state.workbook_path)
        if not act.empty:
            st.bar_chart(act.set_index("Partida"), height=400)

        st.markdown("### Detalle por departamento")
        f1,f2 = st.columns(2)

        with f1:
            torre_options = ["Todas"] + sorted([str(x) for x in rdf["Torre"].dropna().unique()]) if "Torre" in rdf.columns else ["Todas"]
            torre_filter = st.selectbox("Torre", torre_options, key="ruta_torre")

        with f2:
            base_floor = rdf.copy()
            if torre_filter != "Todas" and "Torre" in base_floor.columns:
                base_floor = base_floor[base_floor["Torre"].astype(str) == torre_filter]
            piso_options = ["Todos"]
            if "Piso" in base_floor.columns:
                vals = list(base_floor["Piso"].dropna().unique())
                piso_options += sorted(vals, key=lambda x: float(x) if str(x).replace(".","",1).isdigit() else str(x))
            piso_filter = st.selectbox("Piso", piso_options, key="ruta_piso")

        display = route_display(st.session_state.workbook_path, torre_filter, piso_filter)
        st.dataframe(display, use_container_width=True, height=650)

elif page == "⬆️ Importar / Exportar":
    if not can_access_admin_tools():
        st.error("Solo el Administrador puede importar o exportar archivos.")
        st.stop()
    st.info(
        "El Excel descargado incluye también el HISTORIAL_SEMANAL registrado en esta sesión, "
        "para conservar las comparaciones de los viernes."
    )

    uploaded=st.file_uploader("Cargar una nueva versión de CONTROL_FASES_SFCO211.xlsx",type=["xlsx"])
    if uploaded is not None:
        with open(st.session_state.workbook_path,"wb") as f:
            f.write(uploaded.getbuffer())
        load_phase.clear(); get_sheet_names.clear()
        st.success("Archivo cargado.")
        st.rerun()

    with open(st.session_state.workbook_path,"rb") as f:
        data=f.read()

    st.download_button(
        "⬇️ Descargar Excel actualizado",
        data=data,
        file_name="CONTROL_FASES_SFCO211_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if st.button("Restablecer archivo original"):
        shutil.copy2(BASE,st.session_state.workbook_path)
        load_phase.clear(); get_sheet_names.clear()
        st.rerun()

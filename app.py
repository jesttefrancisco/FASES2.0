
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os, tempfile, shutil
from datetime import datetime

st.set_page_config(
    page_title="Control Fases SFCO211",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = os.path.join(os.path.dirname(__file__), "CONTROL_FASES_SFCO211.xlsx")
LOGO = os.path.join(os.path.dirname(__file__), "logo_san_francisco.png")
PHASES = ["FASE1", "FASE2", "FASE3", "FASE4"]

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
    df = phase_numeric_percent_df(path, phase)
    col = "% Avance Real Depto"
    if col not in df.columns or len(df) == 0:
        return 0.0
    return round(float(pd.to_numeric(df[col], errors="coerce").fillna(0).mean()), 1)

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

    for activity, pct in activity_values.items():
        col = header_to_col.get(activity)
        if col:
            ws.cell(excel_row, col).value = from_percent(pct, phase)

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
    """Normaliza Ruta Crítica: 0-1 -> 0-100; 0-100 se conserva."""
    try:
        if value is None or value == "":
            return 0.0
        x = float(value)
        if 0 <= x <= 1:
            return x * 100.0
        if 0 <= x <= 100:
            return x
        return x
    except:
        return None

@st.cache_data(show_spinner=False)
def route_critical_data(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["SEGUIMIENTO R.CRITICA"]
    headers = [c.value if c.value is not None else f"Columna {i+1}" for i,c in enumerate(ws[1])]
    rows=[]
    for r in range(2, ws.max_row+1):
        vals=[ws.cell(r,c).value for c in range(1,ws.max_column+1)]
        if not any(v is not None for v in vals):
            continue
        rows.append(vals)
    wb.close()
    df=pd.DataFrame(rows,columns=headers)

    id_cols = ["Columna 1","Piso","Torre","Departamento"]
    for col in df.columns:
        if col not in id_cols:
            df[col]=df[col].map(normalize_route_pct)
    return df

def route_critical_display(path):
    df=route_critical_data(path).copy()
    id_cols=["Columna 1","Piso","Torre","Departamento"]
    for col in df.columns:
        if col not in id_cols:
            df[col]=df[col].map(lambda x: f"{x:.1f}%" if isinstance(x,(int,float)) else "")
    return df

def route_overall(path):
    df=route_critical_data(path)
    col="Av. Depto/Piso"
    if col not in df.columns or df.empty:
        return 0.0
    vals=pd.to_numeric(df[col],errors="coerce").dropna()
    return round(float(vals.mean()),1) if len(vals) else 0.0

def ensure_history_sheet(path):
    """Crea la hoja HISTORIAL_SEMANAL si aún no existe."""
    wb = load_workbook(path)
    if "HISTORIAL_SEMANAL" not in wb.sheetnames:
        ws = wb.create_sheet("HISTORIAL_SEMANAL")
        ws.append([
            "Fecha actualización",
            "Avance General",
            "Fase 1",
            "Fase 2",
            "Fase 3",
            "Fase 4",
            "Usuario"
        ])
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 14
        ws.column_dimensions["G"].width = 22
    wb.save(path)
    wb.close()

def register_weekly_snapshot(path, summaries, general, username, force=False):
    """
    Registra una sola actualización por fecha.
    La regla normal es viernes. Un Administrador puede forzar el registro.
    """
    ensure_history_sheet(path)
    now = datetime.now()
    is_friday = now.weekday() == 4

    if not is_friday and not force:
        return False, "La actualización semanal corresponde a los viernes."

    date_value = now.strftime("%Y-%m-%d")
    wb = load_workbook(path)
    ws = wb["HISTORIAL_SEMANAL"]

    # Evitar duplicar una actualización del mismo día.
    for r in range(2, ws.max_row + 1):
        existing = ws.cell(r, 1).value
        if existing is not None and str(existing)[:10] == date_value:
            wb.close()
            return False, f"Ya existe una actualización registrada para {date_value}."

    ws.append([
        date_value,
        float(general),
        float(summaries["FASE1"]),
        float(summaries["FASE2"]),
        float(summaries["FASE3"]),
        float(summaries["FASE4"]),
        username
    ])

    # Formato visible como porcentaje (los valores guardados ya son 0–100).
    for c in range(2, 7):
        ws.cell(ws.max_row, c).number_format = '0.0"%"'

    wb.save(path)
    wb.close()
    return True, f"Actualización semanal registrada: {date_value}."

@st.cache_data(show_spinner=False)
def load_weekly_history(path):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        if "HISTORIAL_SEMANAL" not in wb.sheetnames:
            wb.close()
            return pd.DataFrame(columns=[
                "Fecha actualización","Avance General","Fase 1","Fase 2","Fase 3","Fase 4","Usuario"
            ])
        ws = wb["HISTORIAL_SEMANAL"]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if len(rows) <= 1:
            return pd.DataFrame(columns=rows[0] if rows else [
                "Fecha actualización","Avance General","Fase 1","Fase 2","Fase 3","Fase 4","Usuario"
            ])
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df["Fecha actualización"] = pd.to_datetime(df["Fecha actualización"], errors="coerce")
        for c in ["Avance General","Fase 1","Fase 2","Fase 3","Fase 4"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        return df.dropna(subset=["Fecha actualización"]).sort_values("Fecha actualización")
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

with st.sidebar:
    if os.path.exists(LOGO):
        st.image(LOGO, use_container_width=True)
    st.markdown("## CONTROL DE OBRA")
    st.caption("EDIFICIO SAN FRANCISCO 211 · PAZ")
    st.markdown("---")

    page = st.radio(
        "Menú",
        [
            "📊 Dashboard",
            "📆 Comparación semanal",
            "📈 Gráficos por fase",
            "🧱 Actualizar avances",
            "🏢 Fases completas",
            "⚠️ Ruta crítica",
            "⬆️ Importar / Exportar",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.caption(f"👤 {st.session_state.get('user_name','Usuario')}")
    st.caption(f"Rol: {st.session_state.get('user_role','Usuario')}")
    if st.button("🚪 Cerrar sesión", use_container_width=True):
        st.session_state.authenticated=False
        st.session_state.pop("username",None)
        st.session_state.pop("user_name",None)
        st.session_state.pop("user_role",None)
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
general = round(sum(summaries.values())/4,1)

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

st.caption("El porcentaje de cada fase se calcula directamente desde las partidas visibles en «Fases completas», por departamento y luego por fase.")
st.divider()

if page == "📊 Dashboard":
    st.info("Fuente de los indicadores y gráficos: partidas actuales de «Fases completas». El avance se recalcula automáticamente, sin depender de fórmulas almacenadas en Excel.")
    c1,c2,c3,c4,c5=st.columns(5)
    cards=[
        (c1,"AVANCE GENERAL",f"{general:.1f}%","blue"),
        (c2,"FASE 1",f"{summaries['FASE1']:.1f}%","good"),
        (c3,"FASE 2",f"{summaries['FASE2']:.1f}%","warn"),
        (c4,"FASE 3",f"{summaries['FASE3']:.1f}%","warn"),
        (c5,"FASE 4",f"{summaries['FASE4']:.1f}%","bad"),
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
    selected_phase=st.selectbox("Fase para detalle",PHASES,format_func=lambda x:x.replace("FASE","Fase "))
    piso=phase_by_floor(st.session_state.workbook_path,selected_phase)
    if not piso.empty:
        st.bar_chart(piso.set_index("Piso"),height=330)

elif page == "📆 Comparación semanal":
    st.markdown('<div class="section">COMPARACIÓN DE AVANCE SEMANAL</div>', unsafe_allow_html=True)
    st.caption(
        "Esta sección guarda una fotografía del avance del proyecto cada viernes y "
        "permite comparar la evolución del Avance General y de las cuatro fases."
    )

    is_friday = datetime.now().weekday() == 4
    role = st.session_state.get("user_role", "Usuario")
    user = st.session_state.get("user_name", st.session_state.get("username", "Usuario"))

    cstatus, cbutton = st.columns([2,1])
    with cstatus:
        if is_friday:
            st.success("✅ Hoy es viernes: corresponde registrar la actualización semanal.")
        else:
            st.info("📅 " + next_friday_text())

    with cbutton:
        if st.button(
            "📌 Registrar avance de hoy",
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
        st.caption("Como Administrador puedes registrar manualmente una actualización fuera del viernes para pruebas o regularización.")

    hist = load_weekly_history(st.session_state.workbook_path)

    if hist.empty:
        st.warning(
            "Aún no hay semanas registradas. El primer punto aparecerá cuando registres "
            "una actualización semanal."
        )
    else:
        latest = hist.iloc[-1]
        k1,k2,k3,k4,k5 = st.columns(5)
        k1.metric("Avance General", f"{latest['Avance General']:.1f}%")
        k2.metric("Fase 1", f"{latest['Fase 1']:.1f}%")
        k3.metric("Fase 2", f"{latest['Fase 2']:.1f}%")
        k4.metric("Fase 3", f"{latest['Fase 3']:.1f}%")
        k5.metric("Fase 4", f"{latest['Fase 4']:.1f}%")

        chart_df = hist.copy()
        chart_df["Semana"] = chart_df["Fecha actualización"].dt.strftime("%d-%m-%Y")

        st.markdown("### Evolución del avance por semana")
        st.line_chart(
            chart_df.set_index("Semana")[
                ["Avance General","Fase 1","Fase 2","Fase 3","Fase 4"]
            ],
            height=430
        )

        if len(hist) >= 2:
            prev = hist.iloc[-2]
            st.markdown("### Variación respecto de la semana anterior")
            variation = pd.DataFrame({
                "Indicador":["Avance General","Fase 1","Fase 2","Fase 3","Fase 4"],
                "Semana anterior":[
                    prev["Avance General"],prev["Fase 1"],prev["Fase 2"],prev["Fase 3"],prev["Fase 4"]
                ],
                "Semana actual":[
                    latest["Avance General"],latest["Fase 1"],latest["Fase 2"],latest["Fase 3"],latest["Fase 4"]
                ],
            })
            variation["Variación (pp)"] = (
                variation["Semana actual"] - variation["Semana anterior"]
            ).round(1)
            st.dataframe(
                variation.style.format({
                    "Semana anterior":"{:.1f}%",
                    "Semana actual":"{:.1f}%",
                    "Variación (pp)":"{:+.1f}"
                }),
                use_container_width=True,
                hide_index=True
            )

        st.markdown("### Historial de actualizaciones")
        shown = hist.copy()
        shown["Fecha actualización"] = shown["Fecha actualización"].dt.strftime("%d-%m-%Y")
        for c in ["Avance General","Fase 1","Fase 2","Fase 3","Fase 4"]:
            shown[c] = shown[c].map(lambda x: f"{x:.1f}%")
        st.dataframe(shown, use_container_width=True, hide_index=True)

        st.caption(
            "Los puntos del gráfico corresponden a las fotografías de avance guardadas "
            "en la hoja HISTORIAL_SEMANAL del archivo de trabajo."
        )

elif page == "📈 Gráficos por fase":
    st.markdown('<div class="section">GRÁFICOS DE CADA FASE · FUENTE: FASES COMPLETAS</div>',unsafe_allow_html=True)
    st.caption("Todos los gráficos se generan directamente desde los datos de «Fases completas» y trabajan en escala 0%–100%.")

    for phase in PHASES:
        st.markdown(f"## {phase.replace('FASE','Fase ')}")
        st.metric("AVANCE GENERAL DE LA FASE", f"{summaries[phase]:.1f}%")
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
            st.bar_chart(act.set_index("Partida"),height=420)
        st.divider()

elif page == "🧱 Actualizar avances":
    st.markdown('<div class="section">ACTUALIZAR AVANCE DE DEPARTAMENTO</div>',unsafe_allow_html=True)
    st.success("Todos los controles de avance se muestran y editan de 0% a 100%.")

    phase=st.selectbox("Fase",PHASES,format_func=lambda x:x.replace("FASE","Fase "))
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

    if st.button("💾 GUARDAR ACTUALIZACIÓN",type="primary",use_container_width=True):
        save_department(st.session_state.workbook_path,phase,excel_row,edits)
        st.success(f"Departamento {depto} actualizado correctamente.")
        st.rerun()

elif page == "🏢 Fases completas":
    st.markdown('<div class="section">EDITAR FASE COMPLETA</div>', unsafe_allow_html=True)
    st.success("Ahora puedes editar directamente todas las partidas de cada Fase. Los valores se trabajan de 0% a 100%.")

    phase = st.selectbox(
        "Selecciona una fase",
        PHASES,
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
        if st.button("💾 GUARDAR CAMBIOS DE LA FASE", type="primary", use_container_width=True):
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
    st.markdown('<div class="section">RUTA CRÍTICA · AVANCE</div>', unsafe_allow_html=True)

    rdf = route_critical_data(st.session_state.workbook_path)
    rdisplay = route_critical_display(st.session_state.workbook_path)
    roverall = route_overall(st.session_state.workbook_path)

    c1,c2,c3 = st.columns(3)
    c1.metric("AVANCE PROMEDIO RUTA CRÍTICA", f"{roverall:.1f}%")

    if "Piso" in rdf.columns and "Av. Depto/Piso" in rdf.columns:
        piso_avg = (
            rdf.groupby("Piso", dropna=True)["Av. Depto/Piso"]
            .mean()
            .reset_index()
        )
        c2.metric("PISOS CON REGISTRO", str(piso_avg["Piso"].nunique()))
    else:
        piso_avg = pd.DataFrame()

    if "Departamento" in rdf.columns:
        dep_mask = rdf["Departamento"].astype(str).str.lower() != "todos"
        c3.metric("DEPARTAMENTOS CONTROLADOS", str(int(dep_mask.sum())))

    st.progress(min(max(roverall/100,0),1))
    st.success(
        "Ruta Crítica está leyendo y mostrando los avances. "
        "Los valores en escala decimal (por ejemplo 0,30) se convierten automáticamente a 30%."
    )

    left,right = st.columns(2)
    with left:
        st.markdown("### Avance de Ruta Crítica por piso")
        if not piso_avg.empty:
            piso_avg["Av. Depto/Piso"] = piso_avg["Av. Depto/Piso"].round(1)
            st.bar_chart(piso_avg.set_index("Piso"), height=340)

    with right:
        st.markdown("### Avance por torre")
        if "Torre" in rdf.columns and "Av. Depto/Piso" in rdf.columns:
            tower = (
                rdf.groupby("Torre", dropna=True)["Av. Depto/Piso"]
                .mean()
                .reset_index()
            )
            tower["Av. Depto/Piso"] = tower["Av. Depto/Piso"].round(1)
            st.bar_chart(tower.set_index("Torre"), height=340)

    st.markdown("### Detalle de Ruta Crítica")
    st.caption("Todas las partidas se muestran como porcentaje 0%–100%.")
    st.dataframe(rdisplay, use_container_width=True, height=650)

elif page == "⬆️ Importar / Exportar":
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

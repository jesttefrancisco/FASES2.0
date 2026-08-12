
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from io import BytesIO
import os, tempfile, shutil, re
from datetime import datetime

st.set_page_config(
    page_title="Control Fases SFCO211",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE = os.path.join(os.path.dirname(__file__), "CONTROL_FASES_SFCO211.xlsx")
LOGO = os.path.join(os.path.dirname(__file__), "logo_san_francisco.png")
PHASES = ["FASE1", "FASE2", "FASE3", "FASE4"]

st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f4f7fb;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#071f36 0%,#0a3b63 100%);}
[data-testid="stSidebar"] *{color:#fff;}
.block-container{padding-top:1.2rem;max-width:1650px;}
.sf-title{font-size:30px;font-weight:850;color:#10233d;letter-spacing:-.4px;margin-bottom:1px}
.sf-sub{font-size:14px;color:#708096;margin-bottom:22px}
.card{background:#fff;border:1px solid #e2e8f0;border-radius:15px;padding:18px;box-shadow:0 4px 16px rgba(15,35,60,.04)}
.klabel{font-size:11px;font-weight:800;color:#77849a;text-transform:uppercase}
.kvalue{font-size:30px;font-weight:850;color:#10233d;margin-top:5px}
.good{color:#179b59}.warn{color:#e79b13}.bad{color:#df4545}.blue{color:#1678e7}
.section{font-size:17px;font-weight:850;color:#1b2d47;margin:14px 0 10px}
.smallnote{font-size:12px;color:#758299}
.stButton>button,.stDownloadButton>button{border-radius:9px;font-weight:750}
div[data-testid="stDataFrame"],div[data-testid="stDataEditor"]{
 background:#fff;border-radius:12px;border:1px solid #e2e8f0;padding:5px;
}
</style>
""", unsafe_allow_html=True)

if "workbook_path" not in st.session_state:
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    shutil.copy2(BASE, p)
    st.session_state.workbook_path = p

def phase_scale(sheet_name):
    # En el archivo nuevo FASE1 trabaja en 0–100 y FASE2–4 en 0–1.
    return 100.0 if sheet_name == "FASE1" else 1.0

def to_percent(v, sheet_name):
    if v is None or v == "":
        return 0.0
    try:
        x = float(v)
        return x if phase_scale(sheet_name) == 100.0 else x * 100.0
    except:
        return 0.0

def from_percent(v, sheet_name):
    x = float(v)
    x = min(max(x, 0.0), 100.0)
    return x if phase_scale(sheet_name) == 100.0 else x / 100.0

@st.cache_data(show_spinner=False)
def get_sheet_names(path):
    wb = load_workbook(path, read_only=True, data_only=False)
    names = wb.sheetnames
    wb.close()
    return names

@st.cache_data(show_spinner=False)
def load_phase(path, sheet_name):
    wb = load_workbook(path, data_only=False, read_only=False)
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    rows = []
    excel_rows = []
    for r in range(2, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        # Omitir filas totalmente vacías
        if not any(v is not None for v in vals):
            continue
        rows.append(vals)
        excel_rows.append(r)
    wb.close()
    df = pd.DataFrame(rows, columns=headers)
    df["_excel_row"] = excel_rows
    return df

def recalc_department_progress(ws, row, sheet_name):
    # A:D = identificación. Última columna = avance real.
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

def save_department(path, sheet_name, excel_row, activity_values):
    wb = load_workbook(path)
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    header_to_col = {str(h): i+1 for i, h in enumerate(headers) if h is not None}

    for activity, pct in activity_values.items():
        col = header_to_col.get(activity)
        if col:
            ws.cell(excel_row, col).value = from_percent(pct, sheet_name)

    recalc_department_progress(ws, excel_row, sheet_name)
    wb.save(path)
    wb.close()
    load_phase.clear()
    get_sheet_names.clear()

def phase_summary(path, sheet_name):
    df = load_phase(path, sheet_name)
    progress_col = "% Avance Real Depto"
    if progress_col not in df.columns or len(df) == 0:
        return 0.0
    vals = pd.to_numeric(df[progress_col], errors="coerce").fillna(0)
    vals = vals if phase_scale(sheet_name) == 100.0 else vals * 100.0
    return round(float(vals.mean()), 1)

def all_phase_summaries(path):
    return {p: phase_summary(path, p) for p in PHASES}

def raw_sheet_df(path, sheet_name):
    wb = load_workbook(path, read_only=True, data_only=False)
    ws = wb[sheet_name]
    data = [[c.value for c in row] for row in ws.iter_rows()]
    wb.close()
    width = max((len(r) for r in data), default=1)
    data = [r + [None] * (width - len(r)) for r in data]
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
            "🧱 Actualizar avances",
            "🏢 Fases completas",
            "📈 Avance semanal",
            "⚠️ Ruta crítica",
            "📋 Resumen",
            "🏁 Terminaciones",
            "⬆️ Importar / Exportar",
        ],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption("Datos basados en CONTROL_FASES_SFCO211.xlsx")

hlogo, htitle = st.columns([1.1, 4.9], vertical_alignment="center")
with hlogo:
    if os.path.exists(LOGO):
        st.image(LOGO, use_container_width=True)
with htitle:
    st.markdown('<div class="sf-title">CONTROL FASES SAN FRANCISCO 211</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="sf-sub">Panel de control profesional · {datetime.now().strftime("%d-%m-%Y %H:%M")}</div>',
    unsafe_allow_html=True
)

summaries = all_phase_summaries(st.session_state.workbook_path)
general = round(sum(summaries.values()) / len(summaries), 1)

if page == "📊 Dashboard":
    c1,c2,c3,c4,c5 = st.columns(5)
    items = [
        (c1, "AVANCE GENERAL", f"{general:.1f}%", "blue"),
        (c2, "FASE 1", f"{summaries['FASE1']:.1f}%", "good"),
        (c3, "FASE 2", f"{summaries['FASE2']:.1f}%", "warn"),
        (c4, "FASE 3", f"{summaries['FASE3']:.1f}%", "warn"),
        (c5, "FASE 4", f"{summaries['FASE4']:.1f}%", "bad"),
    ]
    for c,lab,val,cl in items:
        c.markdown(f'<div class="card"><div class="klabel">{lab}</div><div class="kvalue {cl}">{val}</div></div>',unsafe_allow_html=True)

    st.write("")
    left,right = st.columns([1.25,1])
    with left:
        st.markdown('<div class="section">AVANCE POR FASE</div>',unsafe_allow_html=True)
        chart = pd.DataFrame({
            "Fase":["Fase 1","Fase 2","Fase 3","Fase 4"],
            "Avance (%)":[summaries[p] for p in PHASES]
        })
        st.bar_chart(chart.set_index("Fase"), height=330)

    with right:
        st.markdown('<div class="section">ESTADO DE FASES</div>',unsafe_allow_html=True)
        for p in PHASES:
            val = summaries[p]
            st.write(f"**{p.replace('FASE','Fase ')}**")
            st.progress(min(max(val/100,0),1), text=f"{val:.1f}%")

    st.markdown('<div class="section">AVANCE POR PISO Y FASE</div>', unsafe_allow_html=True)
    rows=[]
    for phase in PHASES:
        df=load_phase(st.session_state.workbook_path, phase)
        if "Piso" in df.columns and "% Avance Real Depto" in df.columns:
            tmp=df[["Piso","% Avance Real Depto"]].copy()
            tmp["% Avance Real Depto"]=pd.to_numeric(tmp["% Avance Real Depto"],errors="coerce").fillna(0)
            if phase_scale(phase)==1:
                tmp["% Avance Real Depto"]*=100
            grp=tmp.groupby("Piso",dropna=True)["% Avance Real Depto"].mean()
            for piso,val in grp.items():
                rows.append({"Fase":phase.replace("FASE","Fase "),"Piso":piso,"Avance (%)":round(float(val),1)})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True, height=310)

elif page == "🧱 Actualizar avances":
    st.markdown('<div class="section">ACTUALIZAR AVANCE DE DEPARTAMENTO</div>',unsafe_allow_html=True)
    st.info("Todos los avances se muestran de 0% a 100%, aunque el Excel internamente use escalas distintas.")

    phase = st.selectbox("Fase", PHASES, format_func=lambda x:x.replace("FASE","Fase "))
    df = load_phase(st.session_state.workbook_path, phase)

    towers = sorted([str(x) for x in df["Torre"].dropna().unique()])
    torre = st.selectbox("Torre", towers)

    f1 = df[df["Torre"].astype(str) == torre]
    pisos = sorted(f1["Piso"].dropna().unique(), key=lambda x: float(x) if str(x).replace(".","",1).isdigit() else str(x))
    piso = st.selectbox("Piso", pisos)

    f2 = f1[f1["Piso"] == piso]
    deps = sorted(f2["Departamento"].dropna().unique(), key=lambda x: str(x))
    depto = st.selectbox("Departamento", deps)

    rec = f2[f2["Departamento"] == depto].iloc[0]
    excel_row = int(rec["_excel_row"])

    id_cols = ["Fase","Piso","Torre","Departamento"]
    progress_col = "% Avance Real Depto"
    activity_cols = [c for c in df.columns if c not in id_cols + [progress_col, "_excel_row"]]

    current_progress = to_percent(rec[progress_col], phase)
    st.metric("Avance actual departamento", f"{current_progress:.1f}%")
    st.progress(min(max(current_progress/100,0),1))

    st.markdown('<div class="section">PARTIDAS</div>',unsafe_allow_html=True)

    edits={}
    cols=st.columns(2)
    for i, activity in enumerate(activity_cols):
        current = to_percent(rec[activity], phase)
        with cols[i % 2]:
            edits[activity] = st.number_input(
                activity,
                min_value=0.0,
                max_value=100.0,
                value=float(round(current,1)),
                step=5.0,
                format="%.1f",
                key=f"{phase}-{excel_row}-{i}"
            )

    calculated = round(sum(edits.values()) / len(edits), 1) if edits else 0.0
    st.markdown(f"### Nuevo avance calculado: **{calculated:.1f}%**")
    st.progress(min(max(calculated/100,0),1))

    if st.button("💾 GUARDAR ACTUALIZACIÓN", type="primary", use_container_width=True):
        save_department(st.session_state.workbook_path, phase, excel_row, edits)
        st.success(f"Departamento {depto} actualizado correctamente a {calculated:.1f}%.")
        st.rerun()

elif page == "🏢 Fases completas":
    phase = st.selectbox("Selecciona una fase", PHASES, format_func=lambda x:x.replace("FASE","Fase "))
    df = load_phase(st.session_state.workbook_path, phase).drop(columns=["_excel_row"])
    shown=df.copy()
    activity_cols=[c for c in shown.columns if c not in ["Fase","Piso","Torre","Departamento"]]
    for col in activity_cols:
        shown[col]=pd.to_numeric(shown[col], errors="coerce")
        if phase_scale(phase)==1:
            shown[col]=shown[col]*100
    st.caption("La vista muestra los avances como porcentaje 0–100%.")
    st.dataframe(shown, use_container_width=True, height=680)

elif page == "📈 Avance semanal":
    names=get_sheet_names(st.session_state.workbook_path)
    opts=[n for n in names if "AVANCE SEMANAL" in n.upper()]
    sh=opts[0] if opts else names[0]
    st.dataframe(raw_sheet_df(st.session_state.workbook_path, sh),use_container_width=True,height=700)

elif page == "⚠️ Ruta crítica":
    names=get_sheet_names(st.session_state.workbook_path)
    opts=[n for n in names if "CRITICA" in n.upper()]
    sh=opts[0] if opts else names[0]
    st.warning("Seguimiento de Ruta Crítica")
    st.dataframe(raw_sheet_df(st.session_state.workbook_path, sh),use_container_width=True,height=700)

elif page == "📋 Resumen":
    st.dataframe(raw_sheet_df(st.session_state.workbook_path, "RESUMEN"),use_container_width=True,height=700)

elif page == "🏁 Terminaciones":
    names=get_sheet_names(st.session_state.workbook_path)
    opts=[n for n in names if "TERMIN" in n.upper()]
    sh=opts[0] if opts else names[0]
    st.dataframe(raw_sheet_df(st.session_state.workbook_path, sh),use_container_width=True,height=700)

elif page == "⬆️ Importar / Exportar":
    st.markdown('<div class="section">IMPORTAR / EXPORTAR</div>',unsafe_allow_html=True)
    uploaded=st.file_uploader("Cargar una nueva versión de CONTROL_FASES_SFCO211.xlsx",type=["xlsx"])
    if uploaded is not None:
        with open(st.session_state.workbook_path,"wb") as f:
            f.write(uploaded.getbuffer())
        load_phase.clear(); get_sheet_names.clear()
        st.success("Archivo cargado correctamente.")
        st.rerun()

    with open(st.session_state.workbook_path,"rb") as f:
        data=f.read()

    st.download_button(
        "⬇️ Descargar Excel actualizado",
        data=data,
        file_name="CONTROL_FASES_SFCO211_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    if st.button("Restablecer archivo original"):
        shutil.copy2(BASE, st.session_state.workbook_path)
        load_phase.clear(); get_sheet_names.clear()
        st.rerun()


import streamlit as st
import openpyxl
from openpyxl import load_workbook
import pandas as pd
import os, re, tempfile, shutil
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Control Fases San Francisco", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")
BASE=os.path.join(os.path.dirname(__file__),"CONTROL_FASES_SAN_FRANCISCO.xlsx")

CSS = """
<style>
[data-testid="stAppViewContainer"]{background:#f5f7fb;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#061d34 0%,#073459 100%);}
[data-testid="stSidebar"] *{color:white;}
.block-container{padding-top:1.4rem;max-width:1600px;}
.sf-title{font-size:29px;font-weight:800;color:#10213a;margin-bottom:2px}
.sf-sub{color:#718096;margin-bottom:20px}
.card{background:white;border:1px solid #e3e8ef;border-radius:14px;padding:18px;box-shadow:0 2px 10px rgba(16,33,58,.04);min-height:126px}
.klabel{font-size:12px;font-weight:700;color:#667085;text-transform:uppercase}
.kvalue{font-size:31px;font-weight:800;color:#10213a;margin-top:6px}
.good{color:#17a05d}.warn{color:#e79a00}.bad{color:#df3e3e}.blue{color:#1677e8}
.section{font-size:17px;font-weight:800;color:#172b4d;margin:8px 0 12px}
div[data-testid="stDataFrame"]{background:white;border-radius:12px;border:1px solid #e3e8ef;padding:6px}
.stButton>button,.stDownloadButton>button{border-radius:9px;font-weight:700}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

if "path" not in st.session_state:
    fd,p=tempfile.mkstemp(suffix=".xlsx"); os.close(fd); shutil.copy2(BASE,p); st.session_state.path=p

@st.cache_data(show_spinner=False)
def sheet_names(path):
    w=load_workbook(path,read_only=True,data_only=False); n=w.sheetnames; w.close(); return n

@st.cache_data(show_spinner=False)
def read_raw(path,sheet):
    w=load_workbook(path,read_only=False,data_only=False); ws=w[sheet]
    vals=[[c.value for c in row] for row in ws.iter_rows()]
    w.close()
    width=max((len(r) for r in vals),default=1)
    vals=[r+[None]*(width-len(r)) for r in vals]
    return pd.DataFrame(vals,columns=[f"Columna {i+1}" for i in range(width)])

def pct_values(df):
    vals=[]
    for v in df.to_numpy().flatten():
        if isinstance(v,(int,float)) and not isinstance(v,bool):
            x=float(v)
            if 0<=x<=1: vals.append(x*100)
        elif isinstance(v,str):
            m=re.fullmatch(r"\s*(\d+(?:[.,]\d+)?)\s*%\s*",v)
            if m:
                x=float(m.group(1).replace(",","."))
                if 0<=x<=100: vals.append(x)
    return vals

def phase_metric(path,name):
    vals=pct_values(read_raw(path,name))
    return round(float(pd.Series(vals).median()),1) if vals else None

def save_sheet(path,sheet,df):
    w=load_workbook(path)
    ws=w[sheet]
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            v=df.iat[r,c]
            if pd.isna(v): v=None
            ws.cell(r+1,c+1).value=v
    w.save(path); w.close()
    read_raw.clear(); sheet_names.clear()

names=sheet_names(st.session_state.path)
phases=[n for n in names if n.strip().upper().startswith("FASE")]
pmetrics={p:phase_metric(st.session_state.path,p) for p in phases[:4]}
valid=[v for v in pmetrics.values() if v is not None]
general=round(sum(valid)/len(valid),1) if valid else 0

with st.sidebar:
    st.markdown("## 🏢 SAN FRANCISCO")
    st.caption("CONTROL DE OBRA")
    st.markdown("---")
    page=st.radio("NAVEGACIÓN",["📊 Dashboard","🧱 Fases","📈 Avance semanal","⚠️ Ruta crítica","📅 Gantt","🏁 Terminaciones","✏️ Editar planillas","⬆️ Importar / Exportar"],label_visibility="collapsed")
    st.markdown("---")
    st.caption("Proyecto San Francisco · Control de fases")

st.markdown('<div class="sf-title">CONTROL FASES SAN FRANCISCO</div>',unsafe_allow_html=True)
st.markdown(f'<div class="sf-sub">Panel de control de obra · Última sesión: {datetime.now().strftime("%d-%m-%Y %H:%M")}</div>',unsafe_allow_html=True)

if page=="📊 Dashboard":
    cols=st.columns(5)
    cards=[
        ("AVANCE GENERAL",f"{general:.1f}%","blue"),
        ("FASES",str(len(phases)),"good"),
        ("HOJAS DE CONTROL",str(len(names)),"warn"),
        ("RUTA CRÍTICA","Activa","bad"),
        ("ARCHIVO","En línea","good")
    ]
    for c,(lab,val,cl) in zip(cols,cards):
        c.markdown(f'<div class="card"><div class="klabel">{lab}</div><div class="kvalue {cl}">{val}</div></div>',unsafe_allow_html=True)

    st.write("")
    a,b=st.columns([1.25,1])
    with a:
        st.markdown('<div class="section">AVANCE POR FASE</div>',unsafe_allow_html=True)
        chart=pd.DataFrame({"Fase":[p.strip() for p in pmetrics],"Avance":[v or 0 for v in pmetrics.values()]})
        st.bar_chart(chart.set_index("Fase"),height=310)
    with b:
        st.markdown('<div class="section">ESTADO DEL PROYECTO</div>',unsafe_allow_html=True)
        for p,v in pmetrics.items():
            st.write(f"**{p.strip()}**")
            st.progress(min(max((v or 0)/100,0),1),text=f"{(v or 0):.1f}%")
        st.info("Los porcentajes son indicadores automáticos iniciales detectados desde las hojas de Fase. Podemos vincularlos después a las celdas oficiales exactas de tu planilla.")

    st.markdown('<div class="section">ACCESOS RÁPIDOS</div>',unsafe_allow_html=True)
    modules=[]
    for label,key in [("Avance semanal","AVANCE SEMANAL"),("Ruta crítica","CRITICA"),("Gantt","GANTT"),("Terminaciones","TERMIN")]:
        found=[n for n in names if key in n.upper()]
        modules.append({"Módulo":label,"Hoja vinculada":", ".join(found) if found else "No encontrada","Estado":"Disponible" if found else "Pendiente"})
    st.dataframe(pd.DataFrame(modules),use_container_width=True,hide_index=True)

elif page=="🧱 Fases":
    phase=st.selectbox("Selecciona una fase",phases)
    df=read_raw(st.session_state.path,phase)
    q=st.text_input("🔎 Buscar en la fase",placeholder="Partida, piso, departamento, responsable...")
    view=df
    if q:
        mask=df.astype(str).apply(lambda x:x.str.contains(q,case=False,na=False)).any(axis=1)
        view=df.loc[mask]
    st.dataframe(view,use_container_width=True,height=650)

elif page=="📈 Avance semanal":
    opts=[n for n in names if "AVANCE SEMANAL" in n.upper()]
    sh=st.selectbox("Hoja",opts)
    st.dataframe(read_raw(st.session_state.path,sh),use_container_width=True,height=680)

elif page=="⚠️ Ruta crítica":
    opts=[n for n in names if "CRITICA" in n.upper()]
    sh=opts[0] if opts else names[0]
    st.warning("Seguimiento de Ruta Crítica")
    st.dataframe(read_raw(st.session_state.path,sh),use_container_width=True,height=680)

elif page=="📅 Gantt":
    opts=[n for n in names if "GANTT" in n.upper()]
    sh=opts[0] if opts else names[0]
    st.dataframe(read_raw(st.session_state.path,sh),use_container_width=True,height=680)

elif page=="🏁 Terminaciones":
    opts=[n for n in names if "TERMIN" in n.upper()]
    sh=opts[0] if opts else names[0]
    st.dataframe(read_raw(st.session_state.path,sh),use_container_width=True,height=680)

elif page=="✏️ Editar planillas":
    st.markdown('<div class="section">EDITOR DE PLANILLAS</div>',unsafe_allow_html=True)
    sh=st.selectbox("Hoja a editar",names)
    df=read_raw(st.session_state.path,sh)
    edited=st.data_editor(df,use_container_width=True,height=650,num_rows="fixed")
    if st.button("💾 Guardar cambios",type="primary"):
        save_sheet(st.session_state.path,sh,edited)
        st.success("Cambios guardados en la copia de trabajo.")

elif page=="⬆️ Importar / Exportar":
    st.markdown('<div class="section">IMPORTAR / EXPORTAR EXCEL</div>',unsafe_allow_html=True)
    up=st.file_uploader("Cargar una nueva versión de la planilla",type=["xlsx"])
    if up is not None:
        with open(st.session_state.path,"wb") as f: f.write(up.getbuffer())
        read_raw.clear(); sheet_names.clear()
        st.success("Planilla cargada.")
    with open(st.session_state.path,"rb") as f: data=f.read()
    st.download_button("⬇️ Descargar Excel actualizado",data=data,file_name="CONTROL_FASES_SAN_FRANCISCO_ACTUALIZADO.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if st.button("Restablecer planilla original"):
        shutil.copy2(BASE,st.session_state.path); read_raw.clear(); sheet_names.clear(); st.rerun()

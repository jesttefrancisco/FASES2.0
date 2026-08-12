import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
import os
import tempfile
import shutil
import re

st.set_page_config(page_title="Control de Fases | San Francisco", page_icon="🏗️", layout="wide")
DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "CONTROL_FASES_SAN_FRANCISCO.xlsx")

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
div[data-testid="stMetric"] {border: 1px solid rgba(128,128,128,.25); padding: 12px; border-radius: 12px;}
h1, h2, h3 {letter-spacing: -0.02em;}
</style>
""", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def get_sheet_names(path):
    wb = load_workbook(path, read_only=True, data_only=False)
    names = wb.sheetnames
    wb.close()
    return names

@st.cache_data(show_spinner=False)
def read_sheet(path, sheet_name):
    wb = load_workbook(path, data_only=False, read_only=False)
    ws = wb[sheet_name]
    values = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        values.append([c.value if c.value is not None else "" for c in row])
    wb.close()
    columns = [f"Columna {i+1}" for i in range(len(values[0]) if values else 1)]
    return pd.DataFrame(values, columns=columns)

def update_sheet(base_path, sheet_name, df):
    wb = load_workbook(base_path)
    ws = wb[sheet_name]
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            value = df.iat[r, c]
            if pd.isna(value):
                value = None
            ws.cell(row=r + 1, column=c + 1).value = value
    bio = BytesIO()
    wb.save(bio)
    wb.close()
    bio.seek(0)
    return bio.getvalue()

def detect_percentages(df):
    vals = []
    for col in df.columns:
        for v in df[col]:
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                x = float(v)
                if 0 <= x <= 1:
                    vals.append(x * 100)
            elif isinstance(v, str):
                m = re.fullmatch(r"\s*(\d+(?:[\.,]\d+)?)\s*%\s*", v)
                if m:
                    x = float(m.group(1).replace(",", "."))
                    if 0 <= x <= 100:
                        vals.append(x)
    return vals

def phase_progress(path, sheet_name):
    try:
        vals = detect_percentages(read_sheet(path, sheet_name))
        if not vals:
            return None
        return float(pd.Series(vals).median())
    except Exception:
        return None

if "workbook_path" not in st.session_state:
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    shutil.copy2(DEFAULT_FILE, tmp)
    st.session_state.workbook_path = tmp

st.title("🏗️ CONTROL DE FASES – SAN FRANCISCO")
st.caption("Aplicación web para consultar, editar y exportar el control de obra.")

with st.sidebar:
    st.header("Archivo")
    uploaded = st.file_uploader("Cargar Excel", type=["xlsx"])
    if uploaded is not None:
        with open(st.session_state.workbook_path, "wb") as f:
            f.write(uploaded.getbuffer())
        get_sheet_names.clear(); read_sheet.clear()
        st.success("Archivo cargado")

    if st.button("Restablecer original", use_container_width=True):
        shutil.copy2(DEFAULT_FILE, st.session_state.workbook_path)
        get_sheet_names.clear(); read_sheet.clear()
        st.rerun()

    st.divider()
    st.caption("La aplicación trabaja con una copia, por lo que el Excel original incluido no se modifica.")

sheet_names = get_sheet_names(st.session_state.workbook_path)

tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "✏️ Editar planillas", "🧱 Seguimiento"])

with tab1:
    st.subheader("Resumen del proyecto")
    phases = [s for s in sheet_names if s.strip().upper().startswith("FASE")]
    if phases:
        cols = st.columns(min(4, len(phases)))
        results = []
        for i, phase in enumerate(phases[:4]):
            p = phase_progress(st.session_state.workbook_path, phase)
            results.append(p)
            cols[i].metric(phase.strip(), f"{p:.1f}%" if p is not None else "Sin cálculo")
        valid = [x for x in results if x is not None]
        if valid:
            general = sum(valid) / len(valid)
            st.progress(max(0.0, min(1.0, general / 100)))
            st.caption(f"Avance general referencial: {general:.1f}%")
            st.info("Este cálculo inicial detecta porcentajes presentes en las hojas. Después se puede vincular a las celdas oficiales de avance de tu planilla.")

    modules = {
        "Avance semanal": [s for s in sheet_names if "AVANCE SEMANAL" in s.upper()],
        "Ruta crítica": [s for s in sheet_names if "CRITICA" in s.upper()],
        "Gantt": [s for s in sheet_names if "GANTT" in s.upper()],
        "Terminaciones": [s for s in sheet_names if "TERMIN" in s.upper()],
        "Gráficos": [s for s in sheet_names if "GRAF" in s.upper()],
    }
    st.dataframe(pd.DataFrame([{"Módulo": k, "Hojas": ", ".join(v) or "—"} for k, v in modules.items()]), use_container_width=True, hide_index=True)

with tab2:
    st.subheader("Editor de planillas")
    selected = st.selectbox("Hoja", sheet_names)
    df = read_sheet(st.session_state.workbook_path, selected)
    search = st.text_input("Buscar", placeholder="Ej.: yeso, pintura, departamento…")
    if search:
        mask = df.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        st.caption(f"Coincidencias: {int(mask.sum())} filas")
        st.dataframe(df.loc[mask], use_container_width=True, height=220)

    edited = st.data_editor(df, use_container_width=True, num_rows="fixed", height=560, key=f"ed_{selected}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Guardar cambios", type="primary", use_container_width=True):
            content = update_sheet(st.session_state.workbook_path, selected, edited)
            with open(st.session_state.workbook_path, "wb") as f:
                f.write(content)
            read_sheet.clear(); get_sheet_names.clear()
            st.success(f"Cambios guardados en {selected}")
    with c2:
        with open(st.session_state.workbook_path, "rb") as f:
            download = f.read()
        st.download_button("⬇️ Descargar Excel actualizado", data=download,
                           file_name="CONTROL_FASES_SAN_FRANCISCO_ACTUALIZADO.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

with tab3:
    st.subheader("Seguimiento de obra")
    candidates = [s for s in sheet_names if any(k in s.upper() for k in ["FASE", "AVANCE", "CRITICA", "TERMIN", "GANTT"])]
    selected_module = st.selectbox("Módulo de seguimiento", candidates if candidates else sheet_names)
    st.dataframe(read_sheet(st.session_state.workbook_path, selected_module), use_container_width=True, height=620)
    st.caption("Próxima mejora: formulario simplificado por partida, responsable, estado, fechas y porcentaje de avance.")

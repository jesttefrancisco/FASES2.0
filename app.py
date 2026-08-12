
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os, re, tempfile, shutil
from datetime import datetime

st.set_page_config(
    page_title="Control Fases San Francisco",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = os.path.join(os.path.dirname(__file__), "CONTROL_FASES_SAN_FRANCISCO.xlsx")

# Celdas de Avance General acordadas para cada fase.
OFFICIAL_PROGRESS = {
    "FASE 1": "W3",
    "FASE 2": "R3",
    "FASE 3": "X3",
    "FASE 4": "O3",
}

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
div[data-testid="stDataFrame"],div[data-testid="stDataEditor"]{background:white;border-radius:12px;border:1px solid #e3e8ef;padding:6px}
.stButton>button,.stDownloadButton>button{border-radius:9px;font-weight:700}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

if "path" not in st.session_state:
    fd, p = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    shutil.copy2(BASE, p)
    st.session_state.path = p

def norm_phase_name(name):
    return re.sub(r"\s+", " ", name.strip()).upper()

def is_percent_format(fmt):
    return isinstance(fmt, str) and "%" in fmt

def percent_text(value):
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return f"{value * 100:.0f}%" if abs(value * 100 - round(value * 100)) < 1e-8 else f"{value * 100:.1f}%"
    return "" if value is None else str(value)

def parse_percent(text):
    if text is None:
        return None
    if isinstance(text, (int, float)) and not isinstance(text, bool):
        x = float(text)
        return x / 100 if x > 1 else x

    s = str(text).strip().replace(",", ".")
    if not s:
        return None
    s = s.replace("%", "").strip()
    x = float(s)
    if x < 0 or x > 100:
        raise ValueError("El porcentaje debe estar entre 0 y 100.")
    return x / 100.0

@st.cache_data(show_spinner=False)
def sheet_names(path):
    wb = load_workbook(path, read_only=True, data_only=False)
    names = wb.sheetnames
    wb.close()
    return names

@st.cache_data(show_spinner=False)
def read_editor_data(path, sheet_name):
    """Devuelve texto editable para evitar columnas bloqueadas por tipos mixtos."""
    wb = load_workbook(path, read_only=False, data_only=False)
    ws = wb[sheet_name]

    rows = []
    formats = []
    raw_values = []

    for row in ws.iter_rows():
        display_row = []
        fmt_row = []
        raw_row = []
        for cell in row:
            raw = cell.value
            fmt = cell.number_format or "General"

            if is_percent_format(fmt) and isinstance(raw, (int, float)) and not isinstance(raw, bool):
                display = percent_text(raw)
            else:
                display = "" if raw is None else str(raw)

            display_row.append(display)
            fmt_row.append(fmt)
            raw_row.append(raw)
        rows.append(display_row)
        formats.append(fmt_row)
        raw_values.append(raw_row)

    wb.close()

    width = max((len(r) for r in rows), default=1)
    for collection in (rows, formats, raw_values):
        for r in collection:
            r.extend([""] * (width - len(r)))

    columns = [f"Columna {i+1}" for i in range(width)]
    return pd.DataFrame(rows, columns=columns), formats, raw_values

@st.cache_data(show_spinner=False)
def read_display_data(path, sheet_name):
    """Vista de lectura mostrando las celdas con formato % como porcentaje."""
    df, _, _ = read_editor_data(path, sheet_name)
    return df

def get_official_progress(path, sheet_name):
    normalized = norm_phase_name(sheet_name)
    target = OFFICIAL_PROGRESS.get(normalized)
    if not target:
        return None

    # data_only=True usa el último resultado calculado guardado por Excel.
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    value = ws[target].value
    wb.close()

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 0 <= value <= 1:
            return round(value * 100, 1)
        if 0 <= value <= 100:
            return round(value, 1)
    return None

def save_edited_sheet(path, sheet_name, edited_df, original_df, formats, raw_values):
    wb = load_workbook(path)
    ws = wb[sheet_name]
    changed = 0

    for r in range(edited_df.shape[0]):
        for c in range(edited_df.shape[1]):
            new_text = edited_df.iat[r, c]
            old_text = original_df.iat[r, c]

            # Solo tocar celdas realmente modificadas para preservar fórmulas y tipos.
            if str(new_text) == str(old_text):
                continue

            cell = ws.cell(row=r + 1, column=c + 1)
            fmt = formats[r][c] if r < len(formats) and c < len(formats[r]) else cell.number_format
            raw = raw_values[r][c] if r < len(raw_values) and c < len(raw_values[r]) else cell.value

            try:
                if is_percent_format(fmt):
                    cell.value = parse_percent(new_text)
                    # Mantener formato porcentaje del Excel.
                    if "%" not in (cell.number_format or ""):
                        cell.number_format = "0%"
                else:
                    s = "" if new_text is None else str(new_text).strip()

                    if s == "":
                        cell.value = None
                    elif s.startswith("="):
                        cell.value = s
                    elif isinstance(raw, bool):
                        cell.value = s.lower() in ("true", "1", "sí", "si")
                    elif isinstance(raw, int) and not isinstance(raw, bool):
                        try:
                            cell.value = int(float(s.replace(",", ".")))
                        except:
                            cell.value = s
                    elif isinstance(raw, float):
                        try:
                            cell.value = float(s.replace(",", "."))
                        except:
                            cell.value = s
                    else:
                        cell.value = s
                changed += 1
            except ValueError as e:
                wb.close()
                raise ValueError(f"Error en fila {r+1}, columna {c+1}: {e}")

    wb.save(path)
    wb.close()
    read_editor_data.clear()
    read_display_data.clear()
    sheet_names.clear()
    return changed

names = sheet_names(st.session_state.path)
phases = [n for n in names if norm_phase_name(n).startswith("FASE")]
pmetrics = {p: get_official_progress(st.session_state.path, p) for p in phases[:4]}
valid = [v for v in pmetrics.values() if v is not None]
general = round(sum(valid) / len(valid), 1) if valid else 0.0

with st.sidebar:
    st.markdown("## 🏢 SAN FRANCISCO")
    st.caption("CONTROL DE OBRA")
    st.markdown("---")
    page = st.radio(
        "NAVEGACIÓN",
        ["📊 Dashboard","🧱 Fases","📈 Avance semanal","⚠️ Ruta crítica",
         "📅 Gantt","🏁 Terminaciones","✏️ Editar planillas","⬆️ Importar / Exportar"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption("Proyecto San Francisco · Control de fases")

st.markdown('<div class="sf-title">CONTROL FASES SAN FRANCISCO</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="sf-sub">Panel de control de obra · Última sesión: {datetime.now().strftime("%d-%m-%Y %H:%M")}</div>',
    unsafe_allow_html=True
)

if page == "📊 Dashboard":
    cols = st.columns(5)
    cards = [
        ("AVANCE GENERAL", f"{general:.1f}%", "blue"),
        ("FASES", str(len(phases)), "good"),
        ("HOJAS DE CONTROL", str(len(names)), "warn"),
        ("RUTA CRÍTICA", "Activa", "bad"),
        ("ARCHIVO", "En línea", "good"),
    ]
    for c, (lab, val, cl) in zip(cols, cards):
        c.markdown(
            f'<div class="card"><div class="klabel">{lab}</div><div class="kvalue {cl}">{val}</div></div>',
            unsafe_allow_html=True
        )

    st.write("")
    a, b = st.columns([1.25, 1])

    with a:
        st.markdown('<div class="section">AVANCE POR FASE</div>', unsafe_allow_html=True)
        chart = pd.DataFrame({
            "Fase": [p.strip() for p in pmetrics],
            "Avance (%)": [v or 0 for v in pmetrics.values()],
        })
        st.bar_chart(chart.set_index("Fase"), height=310)

    with b:
        st.markdown('<div class="section">ESTADO DEL PROYECTO</div>', unsafe_allow_html=True)
        for p, v in pmetrics.items():
            st.write(f"**{p.strip()}**")
            if v is None:
                st.caption("Sin valor calculado guardado en el Excel")
            else:
                st.progress(min(max(v / 100, 0), 1), text=f"{v:.1f}%")

        st.success("Los indicadores se leen desde las celdas de Avance General definidas para cada Fase.")

elif page == "🧱 Fases":
    phase = st.selectbox("Selecciona una fase", phases)
    df = read_display_data(st.session_state.path, phase)
    q = st.text_input("🔎 Buscar en la fase", placeholder="Partida, piso, departamento, responsable...")
    view = df
    if q:
        mask = df.astype(str).apply(lambda x: x.str.contains(q, case=False, na=False)).any(axis=1)
        view = df.loc[mask]
    st.dataframe(view, use_container_width=True, height=650)

elif page == "📈 Avance semanal":
    opts = [n for n in names if "AVANCE SEMANAL" in n.upper()]
    sh = st.selectbox("Hoja", opts)
    st.dataframe(read_display_data(st.session_state.path, sh), use_container_width=True, height=680)

elif page == "⚠️ Ruta crítica":
    opts = [n for n in names if "CRITICA" in n.upper()]
    sh = opts[0] if opts else names[0]
    st.warning("Seguimiento de Ruta Crítica")
    st.dataframe(read_display_data(st.session_state.path, sh), use_container_width=True, height=680)

elif page == "📅 Gantt":
    opts = [n for n in names if "GANTT" in n.upper()]
    sh = opts[0] if opts else names[0]
    st.dataframe(read_display_data(st.session_state.path, sh), use_container_width=True, height=680)

elif page == "🏁 Terminaciones":
    opts = [n for n in names if "TERMIN" in n.upper()]
    sh = opts[0] if opts else names[0]
    st.dataframe(read_display_data(st.session_state.path, sh), use_container_width=True, height=680)

elif page == "✏️ Editar planillas":
    st.markdown('<div class="section">EDITOR DE PLANILLAS</div>', unsafe_allow_html=True)
    st.info("Las celdas con formato porcentaje se muestran como 0%–100%. Puedes escribir, por ejemplo, 75% o 75.")

    sh = st.selectbox("Hoja a editar", names)
    original_df, formats, raw_values = read_editor_data(st.session_state.path, sh)

    # Todo se presenta como texto para evitar que Streamlit bloquee columnas de tipos mixtos.
    edited = st.data_editor(
        original_df,
        use_container_width=True,
        height=650,
        num_rows="fixed",
        key=f"editor_{sh}",
    )

    if st.button("💾 Guardar cambios", type="primary"):
        try:
            n = save_edited_sheet(
                st.session_state.path, sh, edited, original_df, formats, raw_values
            )
            st.success(f"Guardado correctamente. Se modificaron {n} celdas.")
            st.rerun()
        except ValueError as e:
            st.error(str(e))

elif page == "⬆️ Importar / Exportar":
    st.markdown('<div class="section">IMPORTAR / EXPORTAR EXCEL</div>', unsafe_allow_html=True)
    up = st.file_uploader("Cargar una nueva versión de la planilla", type=["xlsx"])

    if up is not None:
        with open(st.session_state.path, "wb") as f:
            f.write(up.getbuffer())
        read_editor_data.clear()
        read_display_data.clear()
        sheet_names.clear()
        st.success("Planilla cargada.")

    with open(st.session_state.path, "rb") as f:
        data = f.read()

    st.download_button(
        "⬇️ Descargar Excel actualizado",
        data=data,
        file_name="CONTROL_FASES_SAN_FRANCISCO_ACTUALIZADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if st.button("Restablecer planilla original"):
        shutil.copy2(BASE, st.session_state.path)
        read_editor_data.clear()
        read_display_data.clear()
        sheet_names.clear()
        st.rerun()

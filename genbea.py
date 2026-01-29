import streamlit as st
import pandas as pd
from io import BytesIO
import os
import plotly.express as px
import plotly.graph_objects as go
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors

# ---------- CONTRASEÑA ----------
password_input = st.text_input("Introduce la contraseña:", type="password")
if password_input != st.secrets["PASSWORD"]:
    st.warning("❌ Contraseña incorrecta")
    st.stop()

# ---------- FUNCIONES ----------
@st.cache_data(ttl=3600)
def load_excel(file_path):
    return pd.read_excel(file_path, sheet_name=None)

def combinar_hojas(ficheros_list):
    combined_sheets = {}
    for fichero in ficheros_list:
        for hoja, df in fichero.items():
            if hoja not in combined_sheets:
                combined_sheets[hoja] = df.copy()
            else:
                combined_sheets[hoja] = pd.concat([combined_sheets[hoja], df], ignore_index=True)
    return combined_sheets

# ---------- DETECCIÓN DE ARCHIVOS ----------
data_folder = "datos"
archivos = sorted([f for f in os.listdir(data_folder) if f.endswith(".xlsx")])

# Extraer años y trimestres
trimestres_dict = {}
for f in archivos:
    base = f.replace(".xlsx", "")
    if "-" in base:
        anio, trimestre = base.replace("genbea", "").split("-")
        trimestres_dict.setdefault(anio, []).append(f)
    else:
        anio = base.replace("genbea", "")
        trimestres_dict.setdefault(anio, []).append(f)

# ---------- SELECCIÓN EN LA BARRA LATERAL ----------
st.sidebar.header("📂 Selección de archivo")
anio_sel = st.sidebar.selectbox("Selecciona el año", sorted(trimestres_dict.keys()))
checkbox_anual = st.sidebar.checkbox("📊 Resumen anual (combinar todos los trimestres del año)", value=False)

if checkbox_anual:
    archivos_a_cargar = trimestres_dict[anio_sel]
    st.sidebar.write(f"Se combinarán {len(archivos_a_cargar)} trimestres del año {anio_sel}")
else:
    trimestre_sel = st.sidebar.selectbox("Selecciona el trimestre", sorted(trimestres_dict[anio_sel]))
    archivos_a_cargar = [trimestre_sel]

# ---------- CARGA DE DATOS ----------
ficheros_cargados = [load_excel(os.path.join(data_folder, f)) for f in archivos_a_cargar]
ficheros = combinar_hojas(ficheros_cargados) if checkbox_anual else ficheros_cargados[0]

# ---------- COLUMNAS Y FILTROS ----------
columnas = ["Extracción ADN", "PCRs", "Secuenciación", "Proyecto", "Organismo"]
df_muestras = ficheros["Estado_cepas"]
df_muestras.columns = df_muestras.columns.str.strip()
df_muestras[columnas] = df_muestras[columnas].fillna("No definido")
df_seg = df_muestras.copy()

st.title(f"GENBEA {anio_sel}")
st.write("Estado de las muestras procesadas en el departamento de genética del Banco Español de Algas")

# Filtros por columna
filtros = {}
st.sidebar.header("Filtros por columna")
for col in columnas:
    if col in df_muestras.columns:
        opciones = df_muestras[col].dropna().unique()
        seleccion = st.sidebar.multiselect(col, options=opciones)
        if not seleccion:
            seleccion = opciones
        filtros[col] = seleccion
    else:
        st.sidebar.warning(f"⚠️ Columna no encontrada: {col}")

# Filtro parcial por ID
nombre = st.text_input("Búsqueda por identificador")

# Aplicar filtros
df_filtro = df_muestras.copy()
for col, sel in filtros.items():
    df_filtro = df_filtro[df_filtro[col].isin(sel)]
if nombre:
    df_filtro = df_filtro[df_filtro["Codigo"].astype(str).str.contains(nombre, case=False, na=False)]

# ---------- ESTADÍSTICAS LATERALES ----------
st.sidebar.divider()
st.sidebar.subheader("📊 Número de muestras")
n_muestras = len(df_seg)
n_muestras_filt = len(df_filtro)
n_indef = (df_seg[columnas] == "No definido").any(axis=1).sum()
st.sidebar.metric(label="Nº total de muestras", value=n_muestras)
st.sidebar.metric(label="Nº muestras filtradas", value=n_muestras_filt, delta=n_muestras-n_muestras_filt)
st.sidebar.metric(label="Nº muestras incompletas", value=n_indef)

# ---------- MUESTRA DE DATOS FILTRADOS ----------
st.subheader("Muestras seleccionadas")
st.dataframe(df_filtro, use_container_width=True)

# ---------- FILTRO EN EL RESTO DE HOJAS ----------
muestras = df_filtro["Codigo"].astype(str).tolist()
codigo_prefix = [c[:8] for c in muestras]

filtered_sheets = {"Estado_cepas": df_filtro}
for hoja, df in ficheros.items():
    if hoja != "Estado_cepas":
        st.subheader(hoja)
        if "Codigo" in df.columns:
            df_temp = df.copy()
            df_temp["Codigo_prefix"] = df_temp["Codigo"].astype(str).str[:8]
            df_filtrado = df_temp[df_temp["Codigo_prefix"].isin(codigo_prefix)]
            df_filtrado = df_filtrado.drop(columns=["Codigo_prefix"])
        else:
            df_filtrado = df
        filtered_sheets[hoja] = df_filtrado
        st.dataframe(df_filtrado, use_container_width=True)

img_abs = None
img_adn = None
img_trimestres = None

# ---------- GRÁFICO RESUMEN POR TRIMESTRES ----------
if checkbox_anual:
    resumen_list = []
    for f in archivos_a_cargar:
        df_temp = load_excel(os.path.join(data_folder, f))["Estado_cepas"]
        df_temp.columns = df_temp.columns.str.strip()
        n_total = len(df_temp)
        n_incompletas = (df_temp[columnas] == "No definido").any(axis=1).sum()
        n_completas = n_total - n_incompletas
        resumen_list.append({
            "Trimestre": f.replace(".xlsx", ""),
            "Completas": n_completas,
            "Incompletas": n_incompletas,
            "Total": n_total
        })

    df_resumen = pd.DataFrame(resumen_list)
    st.subheader("📊 Resumen por trimestres")
    st.dataframe(df_resumen, use_container_width=True)

    fig_trimestres = px.bar(
        df_resumen,
        x="Trimestre",
        y=["Completas", "Incompletas"],
        title=f"Estado de las muestras por trimestre ({anio_sel})",
        barmode="group",
        labels={"value": "Número de muestras", "variable": "Estado"}
    )
    st.plotly_chart(fig_trimestres, use_container_width=True)

    img_trimestres = fig_trimestres.to_image(format="png")
    st.download_button(
        label="📥 Descargar gráfico resumen por trimestres (PNG)",
        data=img_trimestres,
        file_name="resumen_trimestres.png",
        mime="image/png"
    )

# ---------- GRÁFICAS DE MÉTRICAS (Absorbancia y ADN) ----------
if "Extraídas" in ficheros:
    df_extraidas = ficheros["Extraídas"]
    df_extraidas.columns = df_extraidas.columns.str.strip()
    df_extra_filtrado = df_extraidas[df_extraidas["Codigo"].astype(str).isin(muestras)]

    if not df_extra_filtrado.empty:
        with st.expander("📊 Mostrar gráficas de métricas"):
            # --- Gráfico de absorbancias ---
            if all(col in df_extra_filtrado.columns for col in ["DNA 260/230", "DNA 260/280"]):
                df_abs = df_extra_filtrado.melt(
                    id_vars=["Codigo"],
                    value_vars=["DNA 260/230", "DNA 260/280"],
                    var_name="Tipo",
                    value_name="Valor")
                df_abs["Valor"] = pd.to_numeric(df_abs["Valor"], errors="coerce")

                # Filtros por calidad
                solo_ideal = st.checkbox("🟢 Mostrar valores óptimos")
                solo_acept = st.checkbox("🟠 Mostrar valores aceptados")
                solo_bajos = st.checkbox("🔴 Mostrar valores no deseables")
                filtro_abs = df_abs.copy()

                cond_ideal = filtro_abs["Valor"].between(1.8, 2.0, inclusive="both")
                cond_acept = filtro_abs["Valor"].between(1.5, 1.8, inclusive="left") | filtro_abs["Valor"].between(2, 2.3, inclusive="right")
                cond_baja = (filtro_abs["Valor"] < 1.5) | (filtro_abs["Valor"] > 2.3)

                filtro_abs_final = pd.Series([False]*len(filtro_abs))
                if solo_ideal: filtro_abs_final |= cond_ideal
                if solo_acept: filtro_abs_final |= cond_acept
                if solo_bajos: filtro_abs_final |= cond_baja
                if not (solo_ideal or solo_acept or solo_bajos):
                    filtro_abs_final = pd.Series([True]*len(filtro_abs))
                filtro_abs = filtro_abs[filtro_abs_final]

                fig_abs = px.bar(filtro_abs, x="Codigo", y="Valor", color="Tipo", barmode="group", title="Absorbancias 260/280 y 260/230 por muestra")
                rangos_abs = st.checkbox("🎨 Mostrar rangos de calidad", value=False)
                if rangos_abs:
                    max_abs = filtro_abs["Valor"].max(skipna=True)
                    fig_abs.add_hrect(y0 = 0.05, y1 = max_abs, fillcolor="red", opacity=0.15, line_width=0.5)
                    fig_abs.add_hrect(y0 = 1.5, y1 = 2.3, fillcolor="orange", opacity=0.15, line_width=0.5)
                    fig_abs.add_hrect(y0 = 1.8, y1 = 2.0, fillcolor="green", opacity=0.15, line_width=0.5)
                n_muestras_abs = filtro_abs["Codigo"].nunique()
                st.caption(f"🧬 Nº muestras representadas: **{n_muestras_abs}**")
                st.plotly_chart(fig_abs, use_container_width=True)
                img_abs = fig_abs.to_image(format="png")
                st.download_button("📥 Descargar gráfica de absorbancias (PNG)", data=img_abs, file_name="absorbancias.png", mime="image/png")

            # --- Gráfico de cantidad de ADN ---
            if "DNA_(ng/uL)" in df_extra_filtrado.columns:
                df_adn_filtrado = df_extra_filtrado.copy()
                df_adn_filtrado["DNA_(ng/uL)"] = pd.to_numeric(df_adn_filtrado["DNA_(ng/uL)"], errors="coerce")

                solo_bajo = st.checkbox("🟡 Mostrar valores < 20ng/uL")
                solo_medio = st.checkbox("🔵 Mostrar valores 20-50ng/uL")
                solo_alto = st.checkbox("🟣 Mostrar valores > 50ng/uL")
                filtro_adn = df_adn_filtrado.copy()

                cond_bajo = (filtro_adn["DNA_(ng/uL)"] < 20)
                cond_medio = filtro_adn["DNA_(ng/uL)"].between(20, 50, inclusive="both")
                cond_alto = (filtro_adn["DNA_(ng/uL)"] > 50)

                filtro_adn_final = pd.Series(False, index=filtro_adn.index)
                if solo_bajo: filtro_adn_final |= cond_bajo
                if solo_medio: filtro_adn_final |= cond_medio
                if solo_alto: filtro_adn_final |= cond_alto
                if not (solo_bajo or solo_medio or solo_alto):
                    filtro_adn_final[:] = True

                filtro_adn = filtro_adn[filtro_adn_final.values]
                fig_adn = px.bar(filtro_adn, x="Codigo", y="DNA_(ng/uL)", title="Cantidad de ADN por muestra")
                rangos_adn = st.checkbox("🎨 Mostrar rangos de cantidad", value=False)
                if rangos_adn:
                    max_adn = filtro_adn["DNA_(ng/uL)"].max(skipna=True)
                    fig_adn.add_hrect(y0=0.05, y1=20, fillcolor="yellow", opacity=0.15, line_width=0.5)
                    fig_adn.add_hrect(y0=20, y1=50, fillcolor="cyan", opacity=0.15, line_width=0.5)
                    fig_adn.add_hrect(y0=50, y1=max_adn, fillcolor="purple", opacity=0.15, line_width=0.5)
                n_muestras_adn = filtro_adn["Codigo"].nunique()
                st.caption(f"🧬 Nº muestras representadas: **{n_muestras_adn}**")
                st.plotly_chart(fig_adn, use_container_width=True)
                img_adn = fig_adn.to_image(format="png")
                st.download_button("📥 Descargar gráfica de ADN (PNG)", data=img_adn, file_name="cantidad_adn.png", mime="image/png")

# ---------- FUNCION PDF ----------
def generar_pdf(filtered_sheets, img_abs=None, img_adn=None, img_trimestres=None):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Informe GENBEA – Resultados filtrados", styles["Title"]))
    elements.append(Spacer(1, 0.7*cm))

    for sheet_name, df in filtered_sheets.items():
        if df.empty: continue
        elements.append(Paragraph(sheet_name, styles["Heading1"]))
        elements.append(Spacer(1, 0.3*cm))
        tabla_data = [df.columns.tolist()] + df.astype(str).values.tolist()
        num_cols = len(df.columns)
        page_width = A4[0]-2*cm
        col_widths = [page_width/num_cols]*num_cols
        tabla = Table(tabla_data, colWidths=col_widths, repeatRows=1)
        tabla.setStyle([("GRID",(0,0),(-1,-1),0.5,colors.grey), ("BACKGROUND",(0,0),(-1,0),colors.lightgrey), ("FONTSIZE",(0,0),(-1,-1),7), ("ALIGN",(0,0),(-1,-1),"CENTER"), ("VALIGN",(0,0),(-1,-1),"MIDDLE")])
        elements.append(tabla)
        elements.append(Spacer(1,0.5*cm))
        elements.append(PageBreak())

    # Gráficas
    if img_trimestres:
        elements.append(Paragraph("Resumen por trimestres", styles["Heading1"]))
        elements.append(Image(BytesIO(img_trimestres), width=16*cm, height=9*cm))
        elements.append(Spacer(1,0.5*cm))
    if img_abs:
        elements.append(Paragraph("Absorbancias", styles["Heading1"]))
        elements.append(Image(BytesIO(img_abs), width=16*cm, height=9*cm))
        elements.append(Spacer(1,0.5*cm))
    if img_adn:
        elements.append(Paragraph("Cantidad de ADN", styles["Heading1"]))
        elements.append(Image(BytesIO(img_adn), width=16*cm, height=9*cm))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ---------- EXPORTAR INFORME ----------
st.divider()
st.subheader("📄 Exportar informe")

# Excel filtrado
output = BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    for sheet_name, df_sheet in filtered_sheets.items():
        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
output.seek(0)

# PDF
pdf_buffer = generar_pdf(filtered_sheets, img_abs=img_abs, img_adn=img_adn, img_trimestres=img_trimestres)

st.download_button("📄 Descargar informe PDF", data=pdf_buffer, file_name="informe_genbea.pdf", mime="application/pdf")
st.download_button("📥 Descargar datos filtrados", data=output, file_name="muestras_filtradas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

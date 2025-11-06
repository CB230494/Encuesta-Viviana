# app.py
# -*- coding: utf-8 -*-
# ================================================================
# Encuesta MPGP (Sí/No) → Excel con tablas y gráficos
# Requisitos: streamlit, pandas, numpy, xlsxwriter
# pip install streamlit pandas numpy xlsxwriter
# ================================================================
import io
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Encuesta MPGP (Sí/No)", page_icon="✅", layout="wide")

# ----------------------- Preguntas (16) -------------------------
# Ajusta los textos según las páginas indicadas del manual.
PREGUNTAS_SI_NO = [
    "¿Conoce el Nuevo Modelo Preventivo de Gestión Policial (MPGP)?",
    "¿Ha leído el Manual Operativo del MPGP en el último año?",
    "¿Reconoce las dos estrategias del MPGP (Prevención Comunitaria y Operativa)?",
    "¿Ha utilizado formularios oficiales de recolección de información del MPGP?",
    "¿Registra sistemáticamente la información recolectada en su jurisdicción?",
    "¿Conoce los anexos de formularios del Manual Operativo del MPGP?",
    "¿Ha recibido capacitación reciente sobre el uso de dichos formularios?",
    "¿Comparte los resultados de los formularios con su jefatura o equipo?",
    "¿Utiliza medios digitales (Google Forms / Forms / Outlook) para aplicar encuestas?",
    "¿Valida la calidad de los datos antes de analizarlos?",
    "¿Elabora gráficos o tablas para socializar los hallazgos de las encuestas?",
    "¿Los datos recolectados influyen en la toma de decisiones locales?",
    "¿Conoce el procedimiento para resguardar la confidencialidad de la información?",
    "¿Sabe a qué población objetivo se debe aplicar cada formulario del MPGP?",
    "¿Ha aplicado formularios al menos una vez en los últimos 3 meses?",
    "¿Considera suficientes los recursos disponibles para aplicar los formularios?"
]

# ----------------------- Estado (memoria) -----------------------
if "respuestas" not in st.session_state:
    st.session_state.respuestas = []  # cada item es un dict con Delegación, Fecha y Q01..Q16

# ----------------------- Encabezado -----------------------------
st.title("✅ Encuesta MPGP (Sí/No)")
st.caption(
    "Diagnóstico de **Uso de Formularios de Recolección de Información** — "
    "Modelo Preventivo de Gestión Policial (**MPGP**). "
    "Complete los datos generales y responda 16 preguntas con **Sí** o **No**."
)

# ------------------------ Formulario ----------------------------
st.markdown("### 📝 Formulario de Encuesta")

with st.form("form_encuesta"):
    delegacion = st.text_input("Delegación o Jurisdicción", placeholder="Ejemplo: D48-Guadalupe")
    fecha = st.date_input("Fecha de aplicación", value=datetime.today())

    st.markdown("---")
    st.subheader("Preguntas (responda Sí o No)")

    respuestas = {}
    for i, q in enumerate(PREGUNTAS_SI_NO, start=1):
        respuestas[f"Q{i:02d} - {q}"] = st.radio(
            f"{i}. {q}",
            options=["Sí", "No"],
            horizontal=True,
            index=1,
            key=f"q_{i}",
        )

    col_btn = st.columns([1, 1, 6])
    with col_btn[0]:
        enviado = st.form_submit_button("➕ Agregar encuesta")
    with col_btn[1]:
        limpiar = st.form_submit_button("🧹 Limpiar selección")

    if enviado:
        if not delegacion:
            st.warning("Por favor indique la **Delegación o Jurisdicción** antes de guardar la encuesta.")
        else:
            fila = {"Delegación": delegacion, "Fecha": fecha.strftime("%Y-%m-%d")}
            fila.update(respuestas)
            st.session_state.respuestas.append(fila)
            st.success("✅ Encuesta registrada correctamente.")

# ---------------------- Tabla de respuestas ---------------------
st.markdown("### 📋 Respuestas capturadas")
if len(st.session_state.respuestas) == 0:
    st.info("Aún no hay encuestas registradas.")
else:
    df = pd.DataFrame(st.session_state.respuestas)
    st.dataframe(df, use_container_width=True, hide_index=True)

    # ------------------- Resumen Sí/No por pregunta --------------
    solo_pregs = [c for c in df.columns if c.startswith("Q")]
    # conteo Sí/No por columna
    resumen = (
        df[solo_pregs]
        .apply(lambda s: s.value_counts())
        .T.fillna(0)
        .astype(int)
        .rename(columns={"Sí": "Si", "No": "No"})  # nombres sin tilde para Excel
    )
    if "Si" not in resumen.columns:
        resumen["Si"] = 0
    if "No" not in resumen.columns:
        resumen["No"] = 0

    resumen["Total"] = resumen["Si"] + resumen["No"]
    resumen["%Si"] = np.where(resumen["Total"] > 0, (resumen["Si"] / resumen["Total"] * 100).round(1), 0.0)
    resumen["%No"] = 100 - resumen["%Si"]

    st.markdown("### 🧮 Resumen por pregunta")
    st.dataframe(resumen, use_container_width=True)

    # ---------------------- Descarga a Excel ---------------------
    def crear_excel_con_graficos(df_respuestas: pd.DataFrame, resumen_p: pd.DataFrame) -> bytes:
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
            # Hoja 1: Respuestas
            df_respuestas.to_excel(writer, sheet_name="Respuestas", index=False)
            ws_resp = writer.sheets["Respuestas"]
            ws_resp.freeze_panes(1, 0)

            # Hoja 2: Resumen
            resumen_excel = resumen_p.reset_index().rename(columns={"index": "Pregunta"})
            resumen_excel.to_excel(writer, sheet_name="Resumen", index=False)
            ws = writer.sheets["Resumen"]
            ws.freeze_panes(1, 1)

            # Rango de datos base
            nrows = len(resumen_excel)

            # --------- Gráfico 1: Barras apiladas Si/No por pregunta
            chart1 = writer.book.add_chart({"type": "column", "subtype": "stacked"})
            chart1.add_series({
                "name": "Sí",
                "categories": ["Resumen", 1, 0, nrows, 0],  # Preguntas
                "values": ["Resumen", 1, 1, nrows, 1],      # Si
            })
            chart1.add_series({
                "name": "No",
                "categories": ["Resumen", 1, 0, nrows, 0],
                "values": ["Resumen", 1, 2, nrows, 2],      # No
            })
            chart1.set_title({"name": "Respuestas por Pregunta (Si/No)"})
            chart1.set_x_axis({"name": "Pregunta", "label_position": "low"})
            chart1.set_y_axis({"name": "Cantidad"})
            chart1.set_legend({"position": "bottom"})
            ws.insert_chart("H2", chart1, {"x_scale": 1.6, "y_scale": 1.6})

            # --------- Gráfico 2: Torta global Si vs No
            total_si = int(resumen_excel["Si"].sum())
            total_no = int(resumen_excel["No"].sum())
            global_df = pd.DataFrame({"Respuesta": ["Sí", "No"], "Cantidad": [total_si, total_no]})
            # Guardamos la tablita auxiliar para la torta en columnas R:S
            base_row, base_col = 1, 18
            global_df.to_excel(writer, sheet_name="Resumen", index=False, startrow=base_row, startcol=base_col, header=False)

            chart2 = writer.book.add_chart({"type": "pie"})
            chart2.add_series({
                "name": "Distribución Global",
                "categories": ["Resumen", base_row, base_col, base_row + 1, base_col],
                "values": ["Resumen", base_row, base_col + 1, base_row + 1, base_col + 1],
                "data_labels": {"percentage": True, "leader_lines": True}
            })
            chart2.set_title({"name": "Global: Sí vs No"})
            ws.insert_chart("R2", chart2, {"x_scale": 1.1, "y_scale": 1.1})

            # --------- Gráfico 3: % de Sí por pregunta
            chart3 = writer.book.add_chart({"type": "column"})
            chart3.add_series({
                "name": "% Sí",
                "categories": ["Resumen", 1, 0, nrows, 0],
                "values": ["Resumen", 1, 4, nrows, 4],  # %Si
                "data_labels": {"value": True}
            })
            chart3.set_title({"name": "% de Sí por Pregunta"})
            chart3.set_x_axis({"name": "Pregunta", "label_position": "low"})
            chart3.set_y_axis({"name": "%", "major_unit": 10, "min": 0, "max": 100})
            chart3.set_legend({"none": True})
            ws.insert_chart("H28", chart3, {"x_scale": 1.6, "y_scale": 1.4})

        return bio.getvalue()

    excel_bytes = crear_excel_con_graficos(df, resumen)
    st.download_button(
        "⬇️ Descargar Excel con tablas y gráficos",
        data=excel_bytes,
        file_name=f"Encuesta_MPGP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ------------------------- Pie ---------------------------
st.markdown("---")
st.caption("Edita la lista de 16 preguntas en la variable `PREGUNTAS_SI_NO` para ajustarlas exactamente al manual.")


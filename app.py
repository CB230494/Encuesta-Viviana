# app.py
# -*- coding: utf-8 -*-
import io
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Encuesta MPGP (Sí/No)", page_icon="✅", layout="wide")

# ----------------------- Preguntas (16) -------------------------
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
    st.session_state.respuestas = []

# ----------------------- Encabezado -----------------------------
st.title("✅ Encuesta MPGP (Sí/No)")
st.caption(
    "Diagnóstico de **Uso de Formularios de Recolección de Información** — "
    "Modelo Preventivo de Gestión Policial (**MPGP**). Complete los datos y responda 16 preguntas con **Sí** o **No**."
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
            f"{i}. {q}", options=["Sí", "No"], horizontal=True, index=1, key=f"q_{i}"
        )

    col_btn = st.columns([1, 1, 6])
    with col_btn[0]:
        enviado = st.form_submit_button("➕ Agregar encuesta")
    with col_btn[1]:
        st.form_submit_button("🧹 Limpiar selección")

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
    resumen = (
        df[solo_pregs]
        .apply(lambda s: s.value_counts())
        .T.fillna(0).astype(int)
        .rename(columns={"Sí": "Si", "No": "No"})
    )
    if "Si" not in resumen.columns: resumen["Si"] = 0
    if "No" not in resumen.columns: resumen["No"] = 0
    resumen["Total"] = resumen["Si"] + resumen["No"]
    resumen["%Si"] = np.where(resumen["Total"] > 0, (resumen["Si"]/resumen["Total"]*100).round(1), 0.0)
    resumen["%No"] = 100 - resumen["%Si"]

    st.markdown("### 🧮 Resumen por pregunta")
    st.dataframe(resumen, use_container_width=True)

    # ---------------------- Descarga a Excel ---------------------
    def crear_excel_con_graficos(df_respuestas: pd.DataFrame, resumen_p: pd.DataFrame) -> bytes:
        bio = io.BytesIO()
        with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
            # ==== Hoja 1: Respuestas
            df_respuestas.to_excel(writer, sheet_name="Respuestas", index=False)
            ws_resp = writer.sheets["Respuestas"]
            ws_resp.freeze_panes(1, 0)
            ws_resp.set_column(0, max(0, len(df_respuestas.columns)-1), 22)

            # ==== Hoja 2: Resumen (con Código corto)
            resumen_excel = resumen_p.reset_index().rename(columns={"index": "Pregunta"})
            resumen_excel.insert(0, "Código", [f"Q{i:02d}" for i in range(1, len(resumen_excel)+1)])
            resumen_excel.to_excel(writer, sheet_name="Resumen", index=False)
            ws = writer.sheets["Resumen"]
            ws.freeze_panes(1, 2)
            ws.set_column("A:A", 8)    # Código
            ws.set_column("B:B", 70)   # Pregunta
            ws.set_column("C:G", 12)   # Si, No, Total, %Si, %No

            # ==== Hoja 3: Gráficos (sin choques)
            ws_g = writer.book.add_worksheet("Gráficos")
            ws_g.set_column("A:A", 2)
            ws_g.set_column("B:M", 12)
            ws_g.set_column("N:Z", 12)
            ws_g.write("B1", "Gráficos de Resultados (MPGP)")

            nrows = len(resumen_excel)

            # -- Gráfico 1: Barras apiladas (Si/No) por Código
            chart1 = writer.book.add_chart({"type": "column", "subtype": "stacked"})
            chart1.add_series({
                "name": "Sí",
                "categories": ["Resumen", 1, 0, nrows, 0],  # Código
                "values":     ["Resumen", 1, 2, nrows, 2],  # Si
            })
            chart1.add_series({
                "name": "No",
                "categories": ["Resumen", 1, 0, nrows, 0],
                "values":     ["Resumen", 1, 3, nrows, 3],  # No
            })
            chart1.set_title({"name": "Respuestas por Pregunta (Si/No)"})
            chart1.set_x_axis({"name": "Código"})
            chart1.set_y_axis({"name": "Cantidad"})
            chart1.set_legend({"position": "bottom"})
            ws_g.insert_chart("B3", chart1, {"x_scale": 1.3, "y_scale": 1.2})

            # -- Gráfico 2: % de Sí por Código
            chart3 = writer.book.add_chart({"type": "column"})
            chart3.add_series({
                "name": "% Sí",
                "categories": ["Resumen", 1, 0, nrows, 0],   # Código
                "values":     ["Resumen", 1, 5, nrows, 5],   # %Si
                "data_labels": {"value": True}
            })
            chart3.set_title({"name": "% de Sí por Pregunta"})
            chart3.set_x_axis({"name": "Código"})
            chart3.set_y_axis({"name": "%", "major_unit": 10, "min": 0, "max": 100})
            chart3.set_legend({"none": True})
            ws_g.insert_chart("B28", chart3, {"x_scale": 1.3, "y_scale": 1.0})

            # -- Gráfico 3: Torta global
            total_si = int(resumen_excel["Si"].sum())
            total_no = int(resumen_excel["No"].sum())
            global_df = pd.DataFrame({"Respuesta": ["Sí", "No"], "Cantidad": [total_si, total_no]})
            base_row, base_col = 2, 14   # Guardamos tabla auxiliar en N/O (col=14→O en 0-based xlsxwriter)
            global_df.to_excel(writer, sheet_name="Gráficos", index=False, startrow=base_row, startcol=base_col)

            chart2 = writer.book.add_chart({"type": "pie"})
            chart2.add_series({
                "name": "Global: Sí vs No",
                "categories": ["Gráficos", base_row+1, base_col, base_row+2, base_col],
                "values":     ["Gráficos", base_row+1, base_col+1, base_row+2, base_col+1],
                "data_labels": {"percentage": True, "leader_lines": True}
            })
            chart2.set_title({"name": "Global: Sí vs No"})
            ws_g.insert_chart("N10", chart2, {"x_scale": 1.1, "y_scale": 1.1})

            # -- Leyenda Código → Pregunta (para leer fácil los códigos)
            ws_g.write("N28", "Leyenda (Código → Pregunta)")
            leyenda = resumen_excel[["Código", "Pregunta"]]
            leyenda.to_excel(writer, sheet_name="Gráficos", startrow=29, startcol=13, index=False)
            ws_g.set_column("N:N", 10)   # Código
            ws_g.set_column("O:Z", 70)   # Pregunta

        return bio.getvalue()

    excel_bytes = crear_excel_con_graficos(df, resumen)
    st.download_button(
        "⬇️ Descargar Excel con tablas y gráficos (ordenado)",
        data=excel_bytes,
        file_name=f"Encuesta_MPGP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")
st.caption("Los gráficos ahora están en una hoja separada (**Gráficos**) y las categorías usan códigos Q01–Q16 para evitar choques.")


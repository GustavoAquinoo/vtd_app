import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import os
import io
import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# 1. Configuración del Dashboard
st.set_page_config(page_title="Visualizador de Curva Cupón Cero (VTD)", layout="wide")
st.title("Dashboard de Vector de Tasa de Descuento Ajustado")

archivos_fuente = {
    "Soles (CSS)": "curva_historica_soles.xlsx",
    "Dólares (CDG)": "curva_historica_dolares.xlsx",
    "Soles VAC (CSS VAC)": "curva_historica_soles_vac.xlsx"
}

# Valores VOLA predeterminados por moneda
VOLA_VALORES = {
    "Soles (CSS)": 0.001502, # 0.1502%
    "Soles VAC (CSS VAC)": 0.002471, # 0.2471%
    "Dólares (CDG)": 0.002175 # 0.2175%
}

# --- NUEVO: Inicialización de variables en memoria para ambas pestañas ---
if 'tablas_mostrar' not in st.session_state:
    st.session_state.tablas_mostrar = {}
if 'datos_grafico' not in st.session_state:
    st.session_state.datos_grafico = []
if 'comp_resultados' not in st.session_state:
    st.session_state.comp_resultados = {}
if 'comp_df_list' not in st.session_state:
    st.session_state.comp_df_list = []

# --- 2A. CARGA EN MEMORIA ---
@st.cache_data
def cargar_datos_en_memoria(ruta_archivo):
    """Lee el archivo crudo y lo deja pivoteado y listo en la memoria RAM."""
    try:
        if ruta_archivo.endswith('.parquet'):
            df_pivot = pd.read_parquet(ruta_archivo)
        else:
            df_raw = pd.read_excel(ruta_archivo, header=1)
            df_raw['Fecha de Proceso'] = pd.to_datetime(df_raw['Fecha de Proceso'], format='%d/%m/%Y')
            df_pivot = df_raw.pivot(index='Fecha de Proceso', columns='Plazo (DIAS)', values='Tasas (%)')

        plazos_completos = list(range(0, 14401, 90))
        df_pivot = df_pivot.reindex(columns=plazos_completos).ffill(axis=1)

        return df_pivot
    except Exception as e:
        return None

# Cargar bases de datos
bases_de_datos = {}
for moneda, ruta in archivos_fuente.items():
    if os.path.exists(ruta):
        bases_de_datos[moneda] = cargar_datos_en_memoria(ruta)

# --- 2B. MOTOR DE CÁLCULO ---
def calcular_curva_cupon_cero(df_pivot, fecha_inicio, fecha_fin):
    """Toma el DataFrame pre-cargado en RAM y calcula la curva."""
    f_inicio = pd.to_datetime(fecha_inicio)
    f_fin = pd.to_datetime(fecha_fin)
    df_filtrado = df_pivot.loc[(df_pivot.index >= f_inicio) & (df_pivot.index <= f_fin)]

    if df_filtrado.empty:
        return None

    tasas_promedio = df_filtrado.mean().values

    meses = np.arange(481)
    df_curva = pd.DataFrame({'Mes': meses})
    df_curva['Año'] = df_curva['Mes'] // 12

    block_idx = df_curva['Mes'] // 3
    idx_inicio = np.clip(block_idx, 0, len(tasas_promedio) - 1)
    idx_fin = np.clip(block_idx + 1, 0, len(tasas_promedio) - 1)

    df_curva['Inicio'] = tasas_promedio[idx_inicio]
    df_curva['Fin'] = tasas_promedio[idx_fin]
    df_curva['Flag Inicio'] = block_idx * 3
    df_curva['Flag Fin'] = (block_idx + 1) * 3

    df_curva['Tasa anual (%)'] = (
        df_curva['Fin'] - ((df_curva['Fin'] - df_curva['Inicio']) * (df_curva['Flag Fin'] - df_curva['Mes']) / 3)
    ) / 100

    df_curva['Tasa mensual (%)'] = (1 + df_curva['Tasa anual (%)']) ** (1/12) - 1

    return df_curva[['Año', 'Tasa anual (%)', 'Mes', 'Tasa mensual (%)']]

# --- 2C. GENERADOR DE EXCEL FORMATO VTD + VOLA ---
def generar_excel_vtd_vola(dict_tasas, fecha_inicio, fecha_fin, valores_vola):
    """Genera el archivo Excel en memoria con valores estáticos (sin fórmulas)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VTD + VOLA"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Times New Roman", size=20, bold=True)
    font_subtitle = Font(name="Times New Roman", size=14, bold=True)
    font_header_group = Font(name="Times New Roman", size=11, bold=True)
    font_header = Font(name="Times New Roman", size=10, bold=True)
    font_data = Font(name="Arial Narrow", size=10)
    font_vola_val = Font(name="Arial Narrow", size=10)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    border_thin_bottom = Border(bottom=Side(style='thin', color='000000'))
    border_last_row = Border(bottom=Side(style='medium', color='000000'))

    ws.merge_cells("B2:P2")
    ws["B2"] = "Vector de tasa de descuento (VTD) y ajuste por volatilidad (VOLA)"
    ws["B2"].font = font_title
    ws["B2"].alignment = align_center

    ws.merge_cells("B3:P3")
    ws["B3"] = "(tasas anuales)"
    ws["B3"].font = font_subtitle
    ws["B3"].alignment = align_center

    ws.row_dimensions[2].height = 25.8
    ws.row_dimensions[6].height = 34

    f_ini_str = fecha_inicio.strftime("%b-%Y")
    f_fin_str = fecha_fin.strftime("%b-%Y")
    texto_periodo = f"VTD aplicable a {f_fin_str}\n(Promedio {f_ini_str} a {f_fin_str})"

    ws.merge_cells("B6:F6")
    ws["B6"] = texto_periodo
    ws["B6"].font = font_header_group
    ws["B6"].alignment = align_center

    for col in range(2, 7):
        ws.cell(row=6, column=col).border = border_thin_bottom

    headers_t1 = ["Mes", "Año", "Soles", "Soles VAC", "Dólares"]
    cols_t1 = ["B", "C", "D", "E", "F"]
    for col_letter, header in zip(cols_t1, headers_t1):
        cell = ws[f"{col_letter}8"]
        cell.value = header
        cell.font = font_header
        cell.alignment = align_center
        cell.border = border_thin_bottom

    ws.merge_cells("H6:L6")
    ws["H6"] = "VTD incluyendo VOLA"
    ws["H6"].font = font_header_group
    ws["H6"].alignment = align_center

    for col in range(8, 13):
        ws.cell(row=6, column=col).border = border_thin_bottom

    headers_t2 = ["Mes", "Año", "Soles", "Soles VAC", "Dólares"]
    cols_t2 = ["H", "I", "J", "K", "L"]
    for col_letter, header in zip(cols_t2, headers_t2):
        cell = ws[f"{col_letter}8"]
        cell.value = header
        cell.font = font_header
        cell.alignment = align_center
        cell.border = border_thin_bottom

    ws.merge_cells("N6:P6")
    ws["N6"] = "VOLA"
    ws["N6"].font = font_header_group
    ws["N6"].alignment = align_center

    for col in range(14, 17):
        ws.cell(row=6, column=col).border = border_thin_bottom

    headers_t3 = ["Soles", "Soles VAC", "Dólares"]
    cols_t3 = ["N", "O", "P"]
    for col_letter, header in zip(cols_t3, headers_t3):
        cell = ws[f"{col_letter}8"]
        cell.value = header
        cell.font = font_header
        cell.alignment = align_center
        cell.border = border_thin_bottom

    vola_soles = valores_vola.get("Soles (CSS)", 0.001502)
    vola_vac = valores_vola.get("Soles VAC (CSS VAC)", 0.002471)
    vola_usd = valores_vola.get("Dólares (CDG)", 0.002175)

    ws["N10"] = vola_soles
    ws["O10"] = vola_vac
    ws["P10"] = vola_usd

    for col_letter in cols_t3:
        cell = ws[f"{col_letter}10"]
        cell.font = font_vola_val
        cell.alignment = align_right
        cell.number_format = '0.0000%'
        cell.border = border_thin_bottom

    df_soles = dict_tasas.get("Soles (CSS)")
    df_vac = dict_tasas.get("Soles VAC (CSS VAC)")
    df_usd = dict_tasas.get("Dólares (CDG)")

    row_start = 10
    total_meses_excel = 1321

    val_soles_curr = 0
    val_vac_curr = 0
    val_usd_curr = 0

    for i in range(total_meses_excel):
        curr_row = row_start + i

        ws[f"B{curr_row}"] = i
        ws[f"C{curr_row}"] = i // 12
        ws[f"H{curr_row}"] = i
        ws[f"I{curr_row}"] = i // 12

        ws[f"B{curr_row}"].font = font_data
        ws[f"C{curr_row}"].font = font_data
        ws[f"H{curr_row}"].font = font_data
        ws[f"I{curr_row}"].font = font_data
        ws[f"B{curr_row}"].alignment = align_center
        ws[f"C{curr_row}"].alignment = align_center
        ws[f"H{curr_row}"].alignment = align_center
        ws[f"I{curr_row}"].alignment = align_center

        if i <= 480:
            val_soles_curr = df_soles['Tasa anual (%)'].iloc[i] if df_soles is not None else 0
            val_vac_curr = df_vac['Tasa anual (%)'].iloc[i] if df_vac is not None else 0
            val_usd_curr = df_usd['Tasa anual (%)'].iloc[i] if df_usd is not None else 0

        ws[f"D{curr_row}"] = val_soles_curr
        ws[f"E{curr_row}"] = val_vac_curr
        ws[f"F{curr_row}"] = val_usd_curr

        for col_letter in ["D", "E", "F"]:
            cell = ws[f"{col_letter}{curr_row}"]
            cell.font = font_data
            cell.alignment = align_right
            cell.number_format = '0.000%'

        ws[f"J{curr_row}"] = val_soles_curr + vola_soles
        ws[f"K{curr_row}"] = val_vac_curr + vola_vac
        ws[f"L{curr_row}"] = val_usd_curr + vola_usd

        for col_letter in ["J", "K", "L"]:
            cell = ws[f"{col_letter}{curr_row}"]
            cell.font = font_data
            cell.alignment = align_right
            cell.number_format = '0.000%'

        if i == total_meses_excel - 1:
            for col_letter in ["B", "C", "D", "E", "F", "H", "I", "J", "K", "L"]:
                ws[f"{col_letter}{curr_row}"].border = border_last_row

    column_widths = {
        'A': 3, 'B': 8, 'C': 8, 'D': 14, 'E': 14, 'F': 14, 'G': 4,
        'H': 8, 'I': 8, 'J': 14, 'K': 14, 'L': 14, 'M': 4,
        'N': 14, 'O': 14, 'P': 14
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()

# ==========================================
# CREACIÓN DE PESTAÑAS (TABS)
# ==========================================
tab_general, tab_comparativa = st.tabs(["📈 Análisis General", "⚖️ Comparativa de Periodos"])

# ==========================================
# PESTAÑA 1: ANÁLISIS GENERAL
# ==========================================
with tab_general:
    st.sidebar.header("Parámetros Análisis General")

    fecha_inicio = st.sidebar.date_input(
        "Fecha Inicio",
        value=datetime.date(2025, 1, 1),
        min_value=datetime.date(2015, 1, 1),
        max_value=datetime.date(2030, 12, 31)
    )

    fecha_fin = st.sidebar.date_input(
        "Fecha Fin",
        value=datetime.date(2025, 12, 31),
        min_value=datetime.date(2015, 1, 1),
        max_value=datetime.date(2030, 12, 31)
    )

    monedas_seleccionadas = st.sidebar.multiselect(
        "Seleccione las curvas a graficar:",
        options=list(archivos_fuente.keys()),
        default=["Soles (CSS)"]
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Ajuste por Volatilidad (VOLA)")
    vola_soles_in = st.sidebar.number_input("VOLA Soles (%)", value=0.1502, format="%.4f") / 100
    vola_vac_in = st.sidebar.number_input("VOLA Soles VAC (%)", value=0.2471, format="%.4f") / 100
    vola_usd_in = st.sidebar.number_input("VOLA Dólares (%)", value=0.2175, format="%.4f") / 100

    volas_usuario = {
        "Soles (CSS)": vola_soles_in,
        "Soles VAC (CSS VAC)": vola_vac_in,
        "Dólares (CDG)": vola_usd_in
    }

    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 Disponibilidad de Datos")

    for moneda, df_pivot in bases_de_datos.items():
        if df_pivot is not None:
            f_min, f_max = df_pivot.index.min(), df_pivot.index.max()
            st.sidebar.info(f"**{moneda}**\n\nDesde: {f_min.strftime('%d/%m/%Y')}\nHasta: {f_max.strftime('%d/%m/%Y')}")
        else:
            st.sidebar.error(f"**{moneda}:** Error al cargar datos")

    if st.button("Generar curvas de la VTD"):
        st.session_state.tablas_mostrar = {}
        st.session_state.datos_grafico = []

        if not monedas_seleccionadas:
            st.warning("Por favor, seleccione al menos una moneda.")
        else:
            for moneda in monedas_seleccionadas:
                if moneda in bases_de_datos and bases_de_datos[moneda] is not None:
                    df_resultado = calcular_curva_cupon_cero(bases_de_datos[moneda], fecha_inicio, fecha_fin)

                    if df_resultado is not None:
                        vola_aplicable = volas_usuario.get(moneda, 0)
                        df_resultado['Tasa anual + VOLA (%)'] = df_resultado['Tasa anual (%)'] + vola_aplicable
                        st.session_state.tablas_mostrar[moneda] = df_resultado

                        df_plot = df_resultado[['Año', 'Mes', 'Tasa anual + VOLA (%)']].copy()
                        df_plot['Moneda'] = moneda
                        st.session_state.datos_grafico.append(df_plot)
                else:
                    st.warning(f"No hay datos en {moneda} para el rango de fechas seleccionado.")

    if st.session_state.datos_grafico:
        df_consolidado = pd.concat(st.session_state.datos_grafico)
        df_consolidado['Tasa Graficar (%)'] = df_consolidado['Tasa anual + VOLA (%)'] * 100

        # --- OPCIÓN DE VISUALIZACIÓN DEL EJE X ---
        vista_eje_x = st.radio(
            "Seleccione la escala del gráfico:",
            ["Mensual", "Anual"],
            horizontal=True,
            key="vista_eje_gen"
        )

        if vista_eje_x == "Mensual":
            df_plot_final = df_consolidado[df_consolidado['Mes'] <= 240]
            col_eje_x = "Mes"
            etiqueta_x = "Plazo (Meses)"
        else:
            df_plot_final = df_consolidado[(df_consolidado['Mes'] % 12 == 0) & (df_consolidado['Año'] <= 20)]
            col_eje_x = "Año"
            etiqueta_x = "Plazo (Años)"

        fig = px.line(
            df_plot_final,
            x=col_eje_x,
            y="Tasa Graficar (%)",
            color="Moneda",
            title="Curva VTDA (Tasa anual + VOLA)",
            labels={"Tasa Graficar (%)": "Tasa Anual + VOLA (%)", col_eje_x: etiqueta_x},
            template="plotly_white",
            markers=True if col_eje_x == "Año" else False
        )
        fig.update_layout(yaxis_tickformat='.4f')

        if col_eje_x == "Año":
            fig.update_xaxes(dtick=1, range=[0, 20])
        else:
            fig.update_xaxes(dtick=6, range=[0, 240])

        st.plotly_chart(fig, use_container_width=True)

        # --- SECCIÓN DE DESCARGA EXCEL ---
        st.subheader("📥 Descarga de reportes Excel")
        str_inicio = fecha_inicio.strftime('%d-%m-%Y')
        str_fin = fecha_fin.strftime('%d-%m-%Y')

        dict_3_monedas = {}
        for m_key in archivos_fuente.keys():
            if m_key in bases_de_datos and bases_de_datos[m_key] is not None:
                dict_3_monedas[m_key] = calcular_curva_cupon_cero(bases_de_datos[m_key], fecha_inicio, fecha_fin)

        col_down1, col_down2 = st.columns(2)

        with col_down1:
            nombre_excel_orig = f"Reporte_VTD_{str_inicio}_al_{str_fin}.xlsx"
            buffer_orig = io.BytesIO()
            with pd.ExcelWriter(buffer_orig, engine='openpyxl') as writer:
                df_resumen = pd.DataFrame({
                    'Parámetro': ['Fecha Inicio de Consulta', 'Fecha Fin de Consulta', 'Curvas Generadas'],
                    'Valor': [str_inicio, str_fin, ", ".join(st.session_state.tablas_mostrar.keys())]
                })
                df_resumen.to_excel(writer, sheet_name='Resumen_Filtros', index=False)

                ws_resumen = writer.sheets['Resumen_Filtros']
                ws_resumen.column_dimensions['A'].width = 25
                ws_resumen.column_dimensions['B'].width = 40

                for moneda, df in st.session_state.tablas_mostrar.items():
                    sheet_name = moneda[:31]
                    df_export = df.copy()
                    df_export.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    ws = writer.sheets[sheet_name]
                    for col_idx, col_name in enumerate(df_export.columns, 1):
                        col_letter = get_column_letter(col_idx)
                        if 'Tasa' in col_name:
                            ws.column_dimensions[col_letter].width = 22
                            for row in range(2, len(df_export) + 2):
                                ws[f'{col_letter}{row}'].number_format = '0.000%'
                        else:
                            ws.column_dimensions[col_letter].width = 10

            st.download_button(
                label=f"📄 Descargar tasas mensuales (según filtro elegido)",
                data=buffer_orig.getvalue(),
                file_name=nombre_excel_orig,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_descarga_individual"
            )

        with col_down2:
            excel_vtd_vola_bytes = generar_excel_vtd_vola(dict_3_monedas, fecha_inicio, fecha_fin, volas_usuario)
            nombre_vtd_vola = f"VTD_y_VOLA_Consolidado_{str_inicio}_al_{str_fin}.xlsx"

            st.download_button(
                label=f"📊 Descargar formato oficial SBS (VTD + VOLA)",
                data=excel_vtd_vola_bytes,
                file_name=nombre_vtd_vola,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_descarga_vtd_vola"
            )

        st.markdown("---")

        cols = st.columns(len(st.session_state.tablas_mostrar))
        for i, (moneda, df) in enumerate(st.session_state.tablas_mostrar.items()):
            with cols[i]:
                st.markdown(f"**{moneda}**")
                df_display = df[['Año', 'Mes', 'Tasa anual + VOLA (%)']]
                st.dataframe(
                    df_display.style.format({'Tasa anual + VOLA (%)': '{:.4%}'}),
                    hide_index=True,
                    height=400
                )

# ==========================================
# PESTAÑA 2: COMPARATIVA DE PERIODOS
# ==========================================
with tab_comparativa:
    st.subheader("Comparar Curva de una Moneda en Múltiples Periodos")

    moneda_comp = st.selectbox("Seleccione la moneda a evaluar:", list(archivos_fuente.keys()), key="moneda_comp")
    vola_default_comp = VOLA_VALORES.get(moneda_comp, 0.001502) * 100

    col1, col2, col3 = st.columns(3)
    periodos_config = []

    with col1:
        st.markdown("##### 📅 Periodo 1")
        p1_inicio = st.date_input("Inicio P1", value=datetime.date(2024, 1, 1), min_value=datetime.date(2015, 1, 1), max_value=datetime.date(2030, 12, 31), key="p1_ini")
        p1_fin = st.date_input("Fin P1", value=datetime.date(2024, 12, 31), min_value=datetime.date(2015, 1, 1), max_value=datetime.date(2030, 12, 31), key="p1_fin")
        p1_vola = st.number_input("VOLA P1 (%)", value=vola_default_comp, format="%.4f", key="vola_p1") / 100
        periodos_config.append(("Periodo 1", p1_inicio, p1_fin, p1_vola))

    with col2:
        st.markdown("##### 📅 Periodo 2")
        p2_act = st.checkbox("Habilitar Periodo 2", value=True)
        p2_inicio = st.date_input("Inicio P2", value=datetime.date(2025, 1, 1), min_value=datetime.date(2015, 1, 1), max_value=datetime.date(2030, 12, 31), key="p2_ini")
        p2_fin = st.date_input("Fin P2", value=datetime.date(2025, 12, 31), min_value=datetime.date(2015, 1, 1), max_value=datetime.date(2030, 12, 31), key="p2_fin")
        p2_vola = st.number_input("VOLA P2 (%)", value=vola_default_comp, format="%.4f", key="vola_p2") / 100
        if p2_act:
            periodos_config.append(("Periodo 2", p2_inicio, p2_fin, p2_vola))

    with col3:
        st.markdown("##### 📅 Periodo 3")
        p3_act = st.checkbox("Habilitar Periodo 3", value=False)
        p3_inicio = st.date_input("Inicio P3", value=datetime.date(2025, 6, 1), min_value=datetime.date(2015, 1, 1), max_value=datetime.date(2030, 12, 31), key="p3_ini")
        p3_fin = st.date_input("Fin P3", value=datetime.date(2025, 12, 31), min_value=datetime.date(2015, 1, 1), max_value=datetime.date(2030, 12, 31), key="p3_fin")
        p3_vola = st.number_input("VOLA P3 (%)", value=vola_default_comp, format="%.4f", key="vola_p3") / 100
        if p3_act:
            periodos_config.append(("Periodo 3", p3_inicio, p3_fin, p3_vola))

    st.markdown("---")

    # --- MODIFICADO: Guardamos la lógica en session_state y dibujamos fuera del botón ---
    if st.button("Generar Comparativa", type="primary"):
        # Limpiamos los estados previos de la pestaña comparativa
        st.session_state.comp_df_list = []
        st.session_state.comp_resultados = {}
        
        if moneda_comp in bases_de_datos and bases_de_datos[moneda_comp] is not None:
            for etiqueta, p_ini, p_fin, p_vola in periodos_config:
                df_p = calcular_curva_cupon_cero(bases_de_datos[moneda_comp], p_ini, p_fin)
                if df_p is not None:
                    df_p['Tasa anual + VOLA (%)'] = df_p['Tasa anual (%)'] + p_vola
                    nombre_periodo = f"{etiqueta} ({p_ini.strftime('%d/%m/%y')} - {p_fin.strftime('%d/%m/%y')})"
                    df_p['Periodo'] = nombre_periodo

                    st.session_state.comp_df_list.append(df_p)
                    st.session_state.comp_resultados[etiqueta] = df_p

            if not st.session_state.comp_df_list:
                st.warning("⚠️ No se encontraron datos suficientes para los periodos seleccionados.")
        else:
            st.error("Los datos de esta moneda no están disponibles o no se pudieron cargar.")

    # Renderizamos SIEMPRE que haya datos en la caché de esta pestaña
    if st.session_state.comp_df_list:
        df_comparativo = pd.concat(st.session_state.comp_df_list)
        df_comparativo['Tasa Graficar (%)'] = df_comparativo['Tasa anual + VOLA (%)'] * 100

        # El st.radio ahora está FUERA del bloque st.button, por lo que no reiniciará el cálculo
        vista_eje_comp = st.radio(
            "Seleccione la escala del gráfico:",
            ["Mensual", "Anual"],
            horizontal=True,
            key="vista_eje_comp"
        )

        if vista_eje_comp == "Mensual":
            df_comp_final = df_comparativo[df_comparativo['Mes'] <= 240]
            col_eje_comp = "Mes"
            etiqueta_x_comp = "Plazo (Meses)"
        else:
            df_comp_final = df_comparativo[(df_comparativo['Mes'] % 12 == 0) & (df_comparativo['Año'] <= 20)]
            col_eje_comp = "Año"
            etiqueta_x_comp = "Plazo (Años)"

        fig_comp = px.line(
            df_comp_final,
            x=col_eje_comp,
            y="Tasa Graficar (%)",
            color="Periodo",
            title=f"Comparativa Curva Cupón Cero - {moneda_comp} (Tasa Anual + VOLA)",
            labels={"Tasa Graficar (%)": "Tasa Anual + VOLA (%)", col_eje_comp: etiqueta_x_comp},
            template="plotly_white",
            color_discrete_sequence=px.colors.qualitative.Set1,
            markers=True if col_eje_comp == "Año" else False
        )
        fig_comp.update_layout(yaxis_tickformat='.4f', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))

        if col_eje_comp == "Año":
            fig_comp.update_xaxes(dtick=1, range=[0, 20])
        else:
            fig_comp.update_xaxes(dtick=6, range=[0, 240])
            
        st.plotly_chart(fig_comp, use_container_width=True)

        cols_tables = st.columns(len(st.session_state.comp_resultados))
        for idx, (etiqueta, df_res) in enumerate(st.session_state.comp_resultados.items()):
            with cols_tables[idx]:
                nombre_periodo = df_res['Periodo'].iloc[0]
                st.markdown(f"**{nombre_periodo}**")
                df_show = df_res[['Año', 'Mes', 'Tasa anual + VOLA (%)']]
                st.dataframe(df_show.style.format({'Tasa anual + VOLA (%)': '{:.4%}'}), hide_index=True, height=250)

        st.markdown("---")
        nombre_moneda_corta = moneda_comp.split()[0]
        nombre_excel_comp = f"Comparativa_VTD_{nombre_moneda_corta}.xlsx"

        buffer_comp = io.BytesIO()
        with pd.ExcelWriter(buffer_comp, engine='openpyxl') as writer:
            df_resumen_data = {
                'Parámetro': ['Moneda Evaluada'] + [f'Periodo {i+1}' for i in range(len(st.session_state.comp_resultados))],
                'Valor': [moneda_comp] + [df['Periodo'].iloc[0] for df in st.session_state.comp_resultados.values()]
            }
            df_resumen_comp = pd.DataFrame(df_resumen_data)
            df_resumen_comp.to_excel(writer, sheet_name='Resumen_Filtros', index=False)

            ws_resumen = writer.sheets['Resumen_Filtros']
            ws_resumen.column_dimensions['A'].width = 20
            ws_resumen.column_dimensions['B'].width = 55

            for etiqueta, df_res in st.session_state.comp_resultados.items():
                df_export = df_res.copy()
                if 'Periodo' in df_export.columns:
                    df_export.drop(columns=['Periodo'], inplace=True)
                    
                df_export.to_excel(writer, sheet_name=etiqueta.replace(" ", "_"), index=False)
                
                ws = writer.sheets[etiqueta.replace(" ", "_")]
                for col_idx, col_name in enumerate(df_export.columns, 1):
                    col_letter = get_column_letter(col_idx)
                    if 'Tasa' in col_name:
                        ws.column_dimensions[col_letter].width = 22
                        for row in range(2, len(df_export) + 2):
                            ws[f'{col_letter}{row}'].number_format = '0.000%'
                    else:
                        ws.column_dimensions[col_letter].width = 10

        st.download_button(
            label="📥 Descargar Comparativa en formato Excel",
            data=buffer_comp.getvalue(),
            file_name=nombre_excel_comp,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_descarga_comparativa"
        )
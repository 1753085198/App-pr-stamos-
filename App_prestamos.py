import streamlit as st
import pandas as pd
import io

def generar_excel(nombre, cedula, capital, meses, tasa_anual):
    # 1. Cálculos de amortización
    tasa_mensual = (tasa_anual / 100) / 12
    if tasa_mensual > 0:
        cuota = capital * (tasa_mensual * (1 + tasa_mensual)**meses) / ((1 + tasa_mensual)**meses - 1)
    else:
        cuota = capital / meses
        
    saldo_restante = capital
    datos_tabla = []
    
    for mes in range(1, meses + 1):
        interes_pagado = saldo_restante * tasa_mensual
        capital_pagado = cuota - interes_pagado
        saldo_restante -= capital_pagado
        
        if saldo_restante < 0.01:
            saldo_restante = 0
            
        datos_tabla.append({
            "Mes": mes,
            "Cuota Fija": round(cuota, 2),
            "Abono a Capital": round(capital_pagado, 2),
            "Interés Pagado": round(interes_pagado, 2),
            "Saldo Restante": round(saldo_restante, 2)
        })
        
    df = pd.DataFrame(datos_tabla)

    # 2. Crear el archivo Excel con formato y encabezado
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Escribimos la tabla a partir de la fila 8 de Excel (startrow=7)
        df.to_excel(writer, index=False, startrow=7, sheet_name="Amortizacion")
        
        # Obtenemos la hoja de trabajo para escribir los datos del cliente arriba
        workbook = writer.book
        worksheet = writer.sheets["Amortizacion"]
        
        # Escribir el encabezado con los datos
        worksheet["A1"] = "REPORTE DE PRÉSTAMO"
        worksheet["A2"] = "Nombre del Cliente:"
        worksheet["B2"] = nombre
        worksheet["A3"] = "Cédula de Identidad:"
        worksheet["B3"] = cedula
        worksheet["A4"] = "Monto del Préstamo:"
        worksheet["B4"] = f"${capital:,.2f}"
        worksheet["A5"] = "Plazo (meses):"
        worksheet["B5"] = meses
        worksheet["A6"] = "Tasa de Interés Anual:"
        worksheet["B6"] = f"{tasa_anual}%"

        # Ajustar el ancho de las columnas para que se lea todo bien
        worksheet.column_dimensions['A'].width = 22
        worksheet.column_dimensions['B'].width = 15
        worksheet.column_dimensions['C'].width = 18
        worksheet.column_dimensions['D'].width = 18
        worksheet.column_dimensions['E'].width = 18
        
    return buffer, df

# --- INTERFAZ GRÁFICA ---
st.set_page_config(page_title="Calculadora de Préstamos", page_icon="💰")

st.title("💰 Calculadora de Amortización")
st.write("Ingresa los datos para generar el Excel completo.")

# Nuevos campos de entrada
nombre = st.text_input("👤 Nombre del cliente:")
cedula = st.text_input("🪪 Número de Cédula:")

col1, col2 = st.columns(2)

with col1:
    capital = st.number_input("💵 Cantidad del préstamo:", min_value=0.0, value=1000.0, step=100.0)
    meses = st.number_input("📅 Número de meses:", min_value=1, value=12, step=1)

with col2:
    tasa_anual = st.number_input("📈 Tasa de interés anual (%):", min_value=0.0, value=15.0, step=1.0)

# Botón de acción
if st.button("Generar Archivo Excel"):
    if nombre.strip() == "" or cedula.strip() == "":
        st.warning("Por favor, ingresa el nombre y la cédula del cliente.")
    else:
        # Ejecutamos la función
        buffer, df = generar_excel(nombre, cedula, capital, meses, tasa_anual)
        
        st.success("¡Documento generado con éxito!")
        
        # Vista previa en la web
        st.write("### Vista previa de los pagos:")
        st.dataframe(df, use_container_width=True)
        
        # Botón de descarga con el nombre o cédula
        nombre_archivo = f"Prestamo_{nombre.replace(' ', '_')}.xlsx"
        st.download_button(
            label="📥 Descargar archivo Excel",
            data=buffer.getvalue(),
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

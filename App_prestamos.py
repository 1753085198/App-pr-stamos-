import streamlit as st
import pandas as pd
import io

def generar_tabla(capital, meses, tasa_anual):
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
        
    return pd.DataFrame(datos_tabla)

# --- INTERFAZ GRÁFICA ---
st.set_page_config(page_title="Calculadora de Préstamos", page_icon="💰")

st.title("💰 Calculadora de Amortización")
st.write("Ingresa los datos para generar la tabla de pagos.")

# Campos de entrada
nombre = st.text_input("👤 Nombre del cliente:")
col1, col2 = st.columns(2)

with col1:
    capital = st.number_input("💵 Cantidad del préstamo:", min_value=0.0, value=1000.0, step=100.0)
    meses = st.number_input("📅 Número de meses:", min_value=1, value=12, step=1)

with col2:
    tasa_anual = st.number_input("📈 Tasa de interés anual (%):", min_value=0.0, value=15.0, step=1.0)

# Botón de acción
if st.button("Generar Tabla"):
    if nombre.strip() == "":
        st.warning("Por favor, ingresa el nombre de la persona.")
    else:
        # Generar datos
        df = generar_tabla(capital, meses, tasa_anual)
        
        st.success("¡Tabla generada con éxito!")
        
        # Mostrar la tabla en la pantalla para que la vea antes de descargar
        st.dataframe(df, use_container_width=True)
        
        # Preparar el archivo Excel en la memoria (para que se pueda descargar en web)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        # Crear botón de descarga
        nombre_archivo = f"Amortizacion_{nombre.replace(' ', '_')}.xlsx"
        st.download_button(
            label="📥 Descargar archivo Excel",
            data=buffer.getvalue(),
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

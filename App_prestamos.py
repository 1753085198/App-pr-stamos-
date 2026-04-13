import streamlit as st
import pandas as pd
import os
import uuid
import io
import glob
from datetime import datetime
from openpyxl.drawing.image import Image as xlImage

ARCHIVO_BD = "base_prestamos.csv"
CARPETA_COMPROBANTES = "comprobantes"

# Crear carpeta para guardar los comprobantes
if not os.path.exists(CARPETA_COMPROBANTES):
    os.makedirs(CARPETA_COMPROBANTES)

# --- FUNCIONES DE BASE DE DATOS Y CÁLCULO ---
def cargar_bd():
    if os.path.exists(ARCHIVO_BD):
        df = pd.read_csv(ARCHIVO_BD, dtype={"Cedula": str, "ID": str})
        if "Tasa" not in df.columns:
            df["Tasa"] = 15.0
        return df
    else:
        columnas = ["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado", "Tasa"]
        df = pd.DataFrame(columns=columnas)
        df.to_csv(ARCHIVO_BD, index=False)
        return df

def guardar_bd(df):
    df.to_csv(ARCHIVO_BD, index=False)

def calcular_cuota(capital, meses, tasa_anual):
    tasa_mensual = (tasa_anual / 100) / 12
    if tasa_mensual > 0:
        return capital * (tasa_mensual * (1 + tasa_mensual)**meses) / ((1 + tasa_mensual)**meses - 1)
    return capital / meses

def generar_tabla_completa(capital, meses, tasa_anual):
    tasa_mensual = (tasa_anual / 100) / 12
    cuota = calcular_cuota(capital, meses, tasa_anual)
    saldo_restante = capital
    datos_tabla = []
    
    for mes in range(1, int(meses) + 1):
        interes_pagado = saldo_restante * tasa_mensual
        capital_pagado = cuota - interes_pagado
        saldo_restante -= capital_pagado
        if saldo_restante < 0.01: saldo_restante = 0
            
        datos_tabla.append({
            "Mes": mes,
            "Cuota Fija": round(cuota, 2),
            "Abono a Capital": round(capital_pagado, 2),
            "Interés Pagado": round(interes_pagado, 2),
            "Saldo Restante": round(saldo_restante, 2)
        })
    return pd.DataFrame(datos_tabla)

def generar_excel(nombre, cedula, capital, meses, tasa_anual, df_tabla, id_cliente):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_tabla.to_excel(writer, index=False, startrow=7, sheet_name="Amortizacion")
        workbook = writer.book
        worksheet = writer.sheets["Amortizacion"]
        
        worksheet["A1"] = "REPORTE DE PRÉSTAMO"
        worksheet["A2"] = "Nombre del Cliente:"
        worksheet["B2"] = nombre
        worksheet["A3"] = "Cédula:"
        worksheet["B3"] = cedula
        worksheet["A4"] = "Monto:"
        worksheet["B4"] = f"${capital:,.2f}"
        worksheet["A5"] = "Plazo:"
        worksheet["B5"] = meses
        worksheet["A6"] = "Tasa Anual:"
        worksheet["B6"] = f"{tasa_anual}%"

        worksheet.column_dimensions['A'].width = 22
        worksheet.column_dimensions['B'].width = 15
        worksheet.column_dimensions['C'].width = 18
        worksheet.column_dimensions['D'].width = 18
        worksheet.column_dimensions['E'].width = 18

        hoja_comp = workbook.create_sheet("Comprobantes")
        hoja_comp["A1"] = f"COMPROBANTES DE PAGO - {nombre}"
        hoja_comp.column_dimensions['A'].width = 40
        
        patron = os.path.join(CARPETA_COMPROBANTES, f"*_{id_cliente}_*.*")
        archivos_imagenes = glob.glob(patron)
        archivos_imagenes.sort()
        
        fila_actual = 3
        if not archivos_imagenes: 
            hoja_comp["A3"] = "No hay comprobantes registrados aún."
        else:
            for ruta_img in archivos_imagenes:
                try:
                    nombre_archivo = os.path.basename(ruta_img)
                    mes_texto = nombre_archivo.split("_")[1] 
                    hoja_comp[f"A{fila_actual}"] = f"Comprobante - Cuota {mes_texto}"
                    fila_actual += 1
                    
                    img = xlImage(ruta_img)
                    img.width = 300
                    img.height = 400
                    hoja_comp.add_image(img, f"A{fila_actual}")
                    fila_actual += 22 
                except Exception as e:
                    hoja_comp[f"A{fila_actual}"] = f"Error al cargar imagen: {ruta_img}"
                    fila_actual += 2

    return buffer

def resaltar_pagados(row, pagos_hechos):
    if row['Mes'] <= pagos_hechos:
        return ['background-color: #d4edda; color: #155724'] * len(row) 
    return [''] * len(row)

# --- INTERFAZ GRÁFICA ---
st.set_page_config(page_title="Sistema de Préstamos", page_icon="🏦")
st.title("🏦 Sistema Automático de Préstamos")

df_prestamos = cargar_bd()

tab1, tab2, tab3, tab4 = st.tabs(["➕ Nuevo", "📋 Lista", "🔍 Detalles/Excel", "💵 Pagar"])

# PESTAÑA 1: NUEVO
with tab1:
    st.subheader("Crear un nuevo préstamo")
    nombre = st.text_input("👤 Nombre del cliente:")
    cedula = st.text_input("🪪 Número de Cédula:")
    col1, col2 = st.columns(2)
    with col1:
        capital = st.number_input("💵 Monto a prestar:", min_value=0.0, value=1000.0, step=100.0)
        meses = st.number_input("📅 Plazo en meses:", min_value=1, value=12, step=1)
    with col2:
        tasa_anual = st.number_input("📈 Tasa de interés anual (%):", min_value=0.0, value=15.0, step=1.0)
        
    if st.button("Guardar Préstamo"):
        if nombre.strip() and cedula.strip():
            cuota = calcular_cuota(capital, meses, tasa_anual)
            nuevo_registro = {
                "ID": str(uuid.uuid4())[:8], "Fecha": datetime.now().strftime("%Y-%m-%d"),
                "Nombre": nombre, "Cedula": cedula, "Monto_Inicial": capital, "Saldo_Restante": capital,
                "Cuota_Mensual": round(cuota, 2), "Meses_Totales": meses, "Pagos_Realizados": 0,
                "Estado": "ACTIVO", "Tasa": tasa_anual
            }
            df_nuevo = pd.DataFrame([nuevo_registro])
            df_prestamos = pd.concat([df_prestamos, df_nuevo], ignore_index=True)
            guardar_bd(df_prestamos)
            st.success(f"¡Préstamo de {nombre} guardado!")
            st.rerun()

# PESTAÑA 2: LISTA (AHORA DIVIDIDA EN ACTIVOS E HISTORIAL)
with tab2:
    st.subheader("Base de datos de clientes")
    if not df_prestamos.empty:
        activos_lista = df_prestamos[df_prestamos["Estado"] == "ACTIVO"]
        pagados_lista = df_prestamos[df_prestamos["Estado"] == "PAGADO"]
        
        st.write(f"### 🟢 Préstamos Activos ({len(activos_lista)})")
        if not activos_lista.empty:
            st.dataframe(activos_lista[["Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Pagos_Realizados"]], use_container_width=True)
        else:
            st.info("No tienes préstamos activos en este momento.")
            
        st.write(f"### ⚪ Historial - Pagados ({len(pagados_lista)})")
        if not pagados_lista.empty:
            st.dataframe(pagados_lista[["Nombre", "Cedula", "Monto_Inicial", "Pagos_Realizados", "Estado"]], use_container_width=True)
        else:
            st.info("Aún no tienes préstamos completamente pagados.")
    else:
        st.info("La base de datos está vacía.")

# PESTAÑA 3: EXCEL CON IMÁGENES Y BOTÓN DE ELIMINAR
with tab3:
    st.subheader("Ver tabla, exportar o eliminar")
    if not df_prestamos.empty:
        opciones_ver = df_prestamos["Nombre"].astype(str) + " - C.I: " + df_prestamos["Cedula"].astype(str) + " (ID: " + df_prestamos["ID"].astype(str) + ")"
        seleccion_ver = st.selectbox("Selecciona un cliente:", opciones_ver, key="ver_cliente")
        
        if seleccion_ver:
            id_ver = seleccion_ver.split("ID: ")[1].replace(")", "")
            cliente_ver = df_prestamos[df_prestamos["ID"] == id_ver].iloc[0]
            
            df_tabla = generar_tabla_completa(cliente_ver["Monto_Inicial"], cliente_ver["Meses_Totales"], cliente_ver["Tasa"])
            st.dataframe(df_tabla.style.apply(resaltar_pagados, pagos_hechos=cliente_ver["Pagos_Realizados"], axis=1), use_container_width=True)
            
            buffer_excel = generar_excel(
                cliente_ver["Nombre"], cliente_ver["Cedula"], cliente_ver["Monto_Inicial"], 
                cliente_ver["Meses_Totales"], cliente_ver["Tasa"], df_tabla, id_ver
            )
            st.download_button(
                label="📥 Descargar Excel (Tabla + Comprobantes)",
                data=buffer_excel.getvalue(),
                file_name=f"Prestamo_{cliente_ver['Nombre'].replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # --- NUEVO: BOTÓN DE ELIMINAR ---
            st.divider() # Pone una línea separadora
            st.write("#### ⚠️ Zona de peligro")
            if st.button("🗑️ Eliminar este préstamo permanentemente"):
                # Filtramos la base de datos para quedarnos con todos EXCEPTO este ID
                df_prestamos = df_prestamos[df_prestamos["ID"] != id_ver]
                guardar_bd(df_prestamos)
                st.success("El registro ha sido eliminado del sistema.")
                st.rerun()

# PESTAÑA 4: PAGAR
with tab4:
    st.subheader("Registrar el pago")
    activos = df_prestamos[df_prestamos["Estado"] == "ACTIVO"]
    
    if not activos.empty:
        opciones_pago = activos["Nombre"].astype(str) + " - C.I: " + activos["Cedula"].astype(str) + " (ID: " + activos["ID"].astype(str) + ")"
        seleccion_pago = st.selectbox("Selecciona el cliente:", opciones_pago, key="pagar_cliente")
        
        if seleccion_pago:
            id_seleccionado = seleccion_pago.split("ID: ")[1].replace(")", "")
            cliente = activos[activos["ID"] == id_seleccionado].iloc[0]
            
            st.write(f"**Cuota:** ${cliente['Cuota_Mensual']} | **Progreso:** {cliente['Pagos_Realizados']} de {cliente['Meses_Totales']}")
            
            comprobante = st.file_uploader("Sube la foto del comprobante (JPG o PNG):", type=["jpg", "jpeg", "png"])
            
            if st.button("Confirmar Pago"):
                idx = df_prestamos[df_prestamos["ID"] == id_seleccionado].index[0]
                cuota_pagada = cliente['Pagos_Realizados'] + 1
                
                if comprobante is not None:
                    ext = comprobante.name.split('.')[-1]
                    nombre_archivo = f"Mes_{cuota_pagada}_{id_seleccionado}_recibo.{ext}"
                    ruta_guardado = os.path.join(CARPETA_COMPROBANTES, nombre_archivo)
                    with open(ruta_guardado, "wb") as f:
                        f.write(comprobante.getbuffer())
                    st.success("Foto guardada y lista para el Excel.")
                
                df_prestamos.at[idx, "Pagos_Realizados"] = cuota_pagada
                abono_capital = cliente['Monto_Inicial'] / cliente['Meses_Totales'] 
                nuevo_saldo = df_prestamos.at[idx, "Saldo_Restante"] - abono_capital
                
                if nuevo_saldo <= 1: 
                    df_prestamos.at[idx, "Saldo_Restante"] = 0
                    df_prestamos.at[idx, "Estado"] = "PAGADO"
                else:
                    df_prestamos.at[idx, "Saldo_Restante"] = round(nuevo_saldo, 2)
                
                guardar_bd(df_prestamos)
                st.success("¡Pago registrado!")
                st.rerun()

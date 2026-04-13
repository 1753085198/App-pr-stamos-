import streamlit as st
import pandas as pd
import os
import uuid
from datetime import datetime

# Nombre del archivo que servirá como nuestra Base de Datos
ARCHIVO_BD = "base_prestamos.csv"

# --- FUNCIONES DE BASE DE DATOS ---
def cargar_bd():
    if os.path.exists(ARCHIVO_BD):
        return pd.read_csv(ARCHIVO_BD)
    else:
        # Si no existe, creamos la estructura
        columnas = ["ID", "Fecha", "Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Meses_Totales", "Pagos_Realizados", "Estado"]
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

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Sistema de Préstamos", page_icon="🏦")
st.title("🏦 Sistema Automático de Préstamos")

# Cargar los datos actuales
df_prestamos = cargar_bd()

# Crear pestañas de navegación
tab1, tab2, tab3 = st.tabs(["➕ Nuevo Préstamo", "📋 Lista de Préstamos", "💵 Registrar Pago"])

# ==========================================
# PESTAÑA 1: NUEVO PRÉSTAMO
# ==========================================
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
            
            # Crear el nuevo registro
            nuevo_registro = {
                "ID": str(uuid.uuid4())[:8], # Genera un código único corto
                "Fecha": datetime.now().strftime("%Y-%m-%d"),
                "Nombre": nombre,
                "Cedula": cedula,
                "Monto_Inicial": capital,
                "Saldo_Restante": capital,
                "Cuota_Mensual": round(cuota, 2),
                "Meses_Totales": meses,
                "Pagos_Realizados": 0,
                "Estado": "ACTIVO"
            }
            
            # Agregar a la base de datos y guardar
            df_nuevo = pd.DataFrame([nuevo_registro])
            df_prestamos = pd.concat([df_prestamos, df_nuevo], ignore_index=True)
            guardar_bd(df_prestamos)
            
            st.success(f"¡Préstamo de {nombre} guardado exitosamente!")
            st.rerun() # Recarga la página para actualizar los datos
        else:
            st.warning("Faltan el nombre o la cédula.")

# ==========================================
# PESTAÑA 2: LISTA DE PRÉSTAMOS
# ==========================================
with tab2:
    st.subheader("Base de datos de clientes")
    if df_prestamos.empty:
        st.info("Aún no hay préstamos registrados.")
    else:
        # Mostrar cuántos clientes hay (ej. los 50 que mencionaste)
        st.write(f"**Total de préstamos registrados:** {len(df_prestamos)}")
        
        # Mostrar la tabla formateada
        st.dataframe(
            df_prestamos[["Nombre", "Cedula", "Monto_Inicial", "Saldo_Restante", "Cuota_Mensual", "Pagos_Realizados", "Estado"]], 
            use_container_width=True
        )

# ==========================================
# PESTAÑA 3: REGISTRAR UN PAGO
# ==========================================
with tab3:
    st.subheader("Registrar el pago de una cuota")
    
    # Filtrar solo los préstamos activos
    activos = df_prestamos[df_prestamos["Estado"] == "ACTIVO"]
    
    if activos.empty:
        st.warning("No hay préstamos activos para registrar pagos.")
    else:
        # Crear una lista de opciones para seleccionar al cliente
        opciones = activos["Nombre"] + " - C.I: " + activos["Cedula"] + " (ID: " + activos["ID"] + ")"
        seleccion = st.selectbox("Selecciona el cliente que va a pagar:", opciones)
        
        if seleccion:
            # Extraer el ID del cliente seleccionado
            id_seleccionado = seleccion.split("ID: ")[1].replace(")", "")
            
            # Obtener los datos de ese cliente
            cliente = activos[activos["ID"] == id_seleccionado].iloc[0]
            
            st.write(f"**Cuota mensual a pagar:** ${cliente['Cuota_Mensual']}")
            st.write(f"**Saldo restante actual:** ${cliente['Saldo_Restante']}")
            st.write(f"**Pagos realizados hasta ahora:** {cliente['Pagos_Realizados']} de {cliente['Meses_Totales']}")
            
            if st.button("Confirmar Pago de Cuota"):
                # Encontrar el índice del cliente en la base de datos original
                idx = df_prestamos[df_prestamos["ID"] == id_seleccionado].index[0]
                
                # Actualizar los valores: Sumar 1 al pago, restar al saldo
                df_prestamos.at[idx, "Pagos_Realizados"] += 1
                
                # Cálculo simplificado para restar el saldo (asumiendo que abona al capital su parte proporcional)
                # En un sistema francés exacto esto varía mes a mes, pero aquí hacemos una aproximación para el control general
                abono_capital = cliente['Monto_Inicial'] / cliente['Meses_Totales'] 
                nuevo_saldo = df_prestamos.at[idx, "Saldo_Restante"] - abono_capital
                
                if nuevo_saldo <= 1: # Si el saldo es casi 0
                    df_prestamos.at[idx, "Saldo_Restante"] = 0
                    df_prestamos.at[idx, "Estado"] = "PAGADO"
                else:
                    df_prestamos.at[idx, "Saldo_Restante"] = round(nuevo_saldo, 2)
                
                guardar_bd(df_prestamos)
                st.success("¡Pago registrado correctamente!")
                st.rerun()

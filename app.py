import streamlit as st
import requests
import pandas as pd
import io

# Configuración básica de la pestaña del navegador
st.set_page_config(page_title="Herramientas SIGOF", layout="centered")

# Inicializar la sesión persistente para guardar las cookies de SIGOF
if 'sesion_sigof' not in st.session_state:
    st.session_state.sesion_sigof = requests.Session()

# Variable de control para saber si ya ingresamos
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def iniciar_sesion(usuario, clave):
    url_login = "http://sigof.distriluz.com.pe/plus/usuario/login"
    payload = {
        "data[Usuario][usuario]": usuario,
        "data[Usuario][pass]": clave
    }
    try:
        respuesta = st.session_state.sesion_sigof.post(url_login, data=payload)
        if "dashboard/modulos" in respuesta.url or respuesta.status_code == 200:
            if "Módulo de Lectura" in respuesta.text or "Módulo de Entrega" in respuesta.text:
                 return True
            elif "dashboard/modulos" in respuesta.url:
                 return True
            else:
                 return False
        else:
            return False
    except Exception as e:
        st.error(f"Error de conexión con el servidor SIGOF: {e}")
        return False

# Interfaz gráfica principal
st.title("Sistema de Automatización")

if not st.session_state.autenticado:
    st.subheader("Acceso a SIGOF")
    
    usuario_input = st.text_input("Usuario")
    clave_input = st.text_input("Contraseña", type="password")
    
    if st.button("Ingresar"):
        if usuario_input and clave_input:
            if iniciar_sesion(usuario_input, clave_input):
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Credenciales incorrectas o el sistema no respondió como se esperaba.")
        else:
            st.warning("Por favor ingresa tu usuario y contraseña.")

else:
    # Esta sección solo se muestra si el login fue exitoso
    st.success("Sesión iniciada correctamente en SIGOF. Conexión establecida.")
    
    # Botón para salir y limpiar la sesión
    if st.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.sesion_sigof.cookies.clear()
        st.rerun()
        
    st.divider()
    
    # --- MÓDULO DE ASIGNACIÓN MASIVA ---
    st.subheader("Módulo: Asignación Masiva de Rutas")
    
    st.write("Para evitar errores en el procesamiento, utiliza la estructura oficial del sistema.")
    
    # Crear la plantilla en memoria
    columnas_requeridas = ['Ciclo', 'SECTOR', 'RUTA REP.', 'NOMBRE COMPLETO', 'ID']
    df_plantilla = pd.DataFrame(columns=columnas_requeridas)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_plantilla.to_excel(writer, index=False, sheet_name='Cronograma')
    
    st.download_button(
        label="Descargar Plantilla Excel",
        data=buffer.getvalue(),
        file_name="Plantilla_Asignacion_Rutas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.write("Una vez llenada la plantilla, súbela a continuación:")
    archivo_excel = st.file_uploader("Sube el cronograma (.xlsx)", type=["xlsx"])

    if archivo_excel is not None:
        df = pd.read_excel(archivo_excel)
        
        # Validar que existan todas las columnas
        columnas_excel = [str(col).strip().upper() for col in df.columns]
        columnas_req_upper = [col.upper() for col in columnas_requeridas]
        
        if all(col in columnas_excel for col in columnas_req_upper):
            st.success("Archivo leído correctamente. Columnas validadas.")
            
            if st.button("Iniciar Asignación Masiva"):
                barra_progreso = st.progress(0)
                log_resultados = []
                
                url_leer = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listacreatelibro2"
                url_guardar = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_guardarlecturistalibro"
                
                total_filas = len(df)
                
                for index, fila in df.iterrows():
                    try:
                        ciclo = str(int(float(fila['Ciclo']))) if pd.notna(fila['Ciclo']) else ""
                        sector = str(int(float(fila['SECTOR']))) if pd.notna(fila['SECTOR']) else ""
                        ruta = str(int(float(fila['RUTA REP.']))) if pd.notna(fila['RUTA REP.']) else ""
                        nombre = str(fila['NOMBRE COMPLETO'])
                        
                        # Extraer el ID
                        if pd.notna(fila['ID']):
                            id_lecturista = str(int(float(fila['ID'])))
                        else:
                            id_lecturista = ""
                    except Exception as e:
                        log_resultados.append(f"[ERROR] Fila {index+1}: Datos inválidos o vacíos. Detalle: {e}")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    if not id_lecturista or id_lecturista == "":
                        log_resultados.append(f"[ERROR] Fila {index+1}: ID vacío para {nombre}")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    # --- PASO 1: LEER LÍMITES DE LA RUTA ---
                    payload_leer = {
                        "ciclo": ciclo,
                        "sector": sector,
                        "ruta": ruta,
                        "id_tipo_reparto": "N"
                    }
                    
                    try:
                        res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                        datos = res_leer.json()
                        
                        if "aaData" in datos and len(datos["aaData"]) > 0:
                            filas_ruta = datos["aaData"]
                            
                            # Validar si ya tiene un lecturista asignado
                            lecturista_actual = str(filas_ruta[0][5]).strip()
                            if lecturista_actual and lecturista_actual != "None" and lecturista_actual != "":
                                log_resultados.append(f"[OMITIDO] Ruta {ruta} ya está asignada a: {lecturista_actual}")
                            else:
                                suministro_inicio = filas_ruta[0][2]
                                suministro_fin = filas_ruta[-1][2]
                                
                                # --- PASO 2: GUARDAR ASIGNACIÓN ---
                                payload_guardar = {
                                    "repartidor": id_lecturista,
                                    "ciclo": ciclo,
                                    "sector": sector,
                                    "ruta": ruta,
                                    "negocio": "82",
                                    "suministro_inicio": suministro_inicio,
                                    "suministro_fin": suministro_fin
                                }
                                
                                res_guardar = st.session_state.sesion_sigof.post(url_guardar, data=payload_guardar)
                                
                                if res_guardar.status_code == 200:
                                    log_resultados.append(f"[EXITO] Ruta {ruta} asignada correctamente a ID {id_lecturista} ({nombre})")
                                else:
                                    log_resultados.append(f"[ERROR] Fallo al guardar Ruta {ruta} en el servidor (Status: {res_guardar.status_code}).")
                        else:
                            log_resultados.append(f"[ALERTA] Ruta {ruta} en Sector {sector} no fue encontrada en SIGOF o no tiene ítems.")
                            
                    except Exception as e:
                        log_resultados.append(f"[ERROR] Fallo de conexión en Ruta {ruta}: {e}")
                    
                    # Actualizar barra de progreso
                    barra_progreso.progress((index + 1) / total_filas)
                
                # Mostrar resumen
                st.divider()
                st.subheader("Reporte de Asignación")
                for log in log_resultados:
                    if "[EXITO]" in log:
                        st.success(log)
                    elif "[ERROR]" in log:
                        st.error(log)
                    elif "[ALERTA]" in log or "[OMITIDO]" in log:
                        st.warning(log)
                    else:
                        st.write(log)
                    
        else:
            st.error(f"El Excel debe contener exactamente estas columnas: {columnas_requeridas}. Revisa que los nombres estén bien escritos.")
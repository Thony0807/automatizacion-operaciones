import streamlit as st
import requests
import pandas as pd
import io

# Configuracion basica de la pestana del navegador
st.set_page_config(page_title="Herramientas SIGOF", layout="centered")

# Inicializar la sesion persistente para guardar las cookies de SIGOF
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
        st.error(f"Error de conexion con el servidor SIGOF: {e}")
        return False

# Interfaz grafica principal
st.title("Sistema de Automatizacion")

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
                st.error("Credenciales incorrectas o el sistema no respondio como se esperaba.")
        else:
            st.warning("Por favor ingresa tu usuario y contraseña.")

else:
    st.success("Sesion iniciada correctamente en SIGOF. Conexion establecida.")
    
    if st.button("Cerrar Sesion"):
        st.session_state.autenticado = False
        st.session_state.sesion_sigof.cookies.clear()
        st.rerun()
        
    st.divider()
    
    # --- MODULO DE ASIGNACION MASIVA ---
    st.subheader("Modulo: Asignacion Masiva de Rutas")
    st.write("Utiliza la plantilla oficial para procesar las rutas.")
    
    # Crear la plantilla en memoria
    columnas_requeridas = ['Ciclo', 'SECTOR', 'RUTA REP.', 'ID']
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

    archivo_excel = st.file_uploader("Sube el cronograma (.xlsx)", type=["xlsx"])

    if archivo_excel is not None:
        df = pd.read_excel(archivo_excel)
        columnas_excel = [str(col).strip().upper() for col in df.columns]
        columnas_req_upper = [col.upper() for col in columnas_requeridas]
        
        if all(col in columnas_excel for col in columnas_req_upper):
            st.success("Archivo leido correctamente. Elige una accion:")
            
            # Crear dos columnas para los botones
            col1, col2 = st.columns(2)
            btn_analizar = col1.button("Analizar Estado de Rutas")
            btn_asignar = col2.button("Iniciar Asignacion Masiva")
            
            url_leer = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listacreatelibro2"
            url_guardar = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_guardarlecturistalibro"
            total_filas = len(df)
            
            # ---------------------------------------------------------
            # LOGICA PARA EL BOTON "ANALIZAR"
            # ---------------------------------------------------------
            if btn_analizar:
                st.info("Analizando rutas en la base de datos de SIGOF...")
                barra_progreso = st.progress(0)
                reporte_analisis = []
                
                for index, fila in df.iterrows():
                    try:
                        ciclo = str(int(float(fila['Ciclo']))) if pd.notna(fila['Ciclo']) else ""
                        sector = str(int(float(fila['SECTOR']))) if pd.notna(fila['SECTOR']) else ""
                        ruta = str(int(float(fila['RUTA REP.']))) if pd.notna(fila['RUTA REP.']) else ""
                        id_lec = str(int(float(fila['ID']))) if pd.notna(fila['ID']) else "FALTA ID"
                        
                        if not ciclo or not sector or not ruta:
                            reporte_analisis.append({"Ruta": "Desconocida", "Sector": sector, "ID Excel": id_lec, "Estado": "Faltan datos (Ciclo, Sector o Ruta)"})
                            continue
                            
                        payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                        res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                        datos = res_leer.json()
                        
                        if "aaData" in datos and len(datos["aaData"]) > 0:
                            lecturista_actual = str(datos["aaData"][0][5]).strip()
                            if lecturista_actual and lecturista_actual != "None":
                                estado_ruta = f"Ocupada (Asignada a: {lecturista_actual})"
                            else:
                                estado_ruta = "Libre para asignar"
                        else:
                            estado_ruta = "No encontrada / Sin items en SIGOF"
                            
                        reporte_analisis.append({"Ruta": ruta, "Sector": sector, "ID Excel": id_lec, "Estado": estado_ruta})
                        
                    except Exception as e:
                        reporte_analisis.append({"Ruta": ruta, "Sector": sector, "ID Excel": "N/A", "Estado": "Error de lectura"})
                        
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.subheader("Reporte Previo de Analisis")
                st.dataframe(pd.DataFrame(reporte_analisis), use_container_width=True)


            # ---------------------------------------------------------
            # LOGICA PARA EL BOTON "ASIGNAR"
            # ---------------------------------------------------------
            if btn_asignar:
                st.info("Iniciando asignacion masiva...")
                barra_progreso = st.progress(0)
                log_resultados = []
                
                for index, fila in df.iterrows():
                    try:
                        ciclo = str(int(float(fila['Ciclo']))) if pd.notna(fila['Ciclo']) else ""
                        sector = str(int(float(fila['SECTOR']))) if pd.notna(fila['SECTOR']) else ""
                        ruta = str(int(float(fila['RUTA REP.']))) if pd.notna(fila['RUTA REP.']) else ""
                        id_lecturista = str(int(float(fila['ID']))) if pd.notna(fila['ID']) else ""
                        
                    except Exception:
                        log_resultados.append(f"[OMITIDO] Fila {index+1}: Motivo: Formato incorrecto o celdas vacias en el Excel.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    if not id_lecturista:
                        log_resultados.append(f"[OMITIDO] Ruta {ruta}: Motivo: No se ingreso el ID del lecturista en el Excel.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                        
                    if not ciclo or not sector or not ruta:
                        log_resultados.append(f"[OMITIDO] Fila {index+1}: Motivo: Faltan datos clave de Ciclo, Sector o Ruta.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                    
                    try:
                        res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                        datos = res_leer.json()
                        
                        if "aaData" in datos and len(datos["aaData"]) > 0:
                            filas_ruta = datos["aaData"]
                            lecturista_actual = str(filas_ruta[0][5]).strip()
                            
                            if lecturista_actual and lecturista_actual != "None":
                                log_resultados.append(f"[OMITIDO] Ruta {ruta}: Motivo: Ya se encuentra asignada previamente a {lecturista_actual}.")
                            else:
                                suministro_inicio = filas_ruta[0][2]
                                suministro_fin = filas_ruta[-1][2]
                                
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
                                    log_resultados.append(f"[EXITO] Ruta {ruta} asignada correctamente al ID {id_lecturista}.")
                                else:
                                    log_resultados.append(f"[ERROR] Ruta {ruta}: Motivo: El servidor de SIGOF rechazo la asignacion (ID de lecturista invalido o problema interno).")
                        else:
                            log_resultados.append(f"[OMITIDO] Ruta {ruta}: Motivo: No fue encontrada en SIGOF bajo ese Ciclo/Sector o no contiene items.")
                            
                    except Exception as e:
                        log_resultados.append(f"[ERROR] Ruta {ruta}: Motivo: Fallo de comunicacion con el servidor durante el proceso.")
                    
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.divider()
                st.subheader("Reporte de Asignacion")
                for log in log_resultados:
                    if "[EXITO]" in log:
                        st.success(log)
                    elif "[ERROR]" in log:
                        st.error(log)
                    elif "[OMITIDO]" in log:
                        st.warning(log)
                    else:
                        st.write(log)
                    
        else:
            st.error(f"El Excel debe contener exactamente estas columnas: {columnas_requeridas}.")
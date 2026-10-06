import streamlit as st
import requests
import pandas as pd
import io

st.set_page_config(page_title="Herramientas SIGOF", layout="wide")

if 'sesion_sigof' not in st.session_state:
    st.session_state.sesion_sigof = requests.Session()
    st.session_state.sesion_sigof.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept-Language": "es-ES,es;q=0.9"
    })

if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def iniciar_sesion(usuario, clave):
    url_login = "http://sigof.distriluz.com.pe/plus/usuario/login"
    payload = {"data[Usuario][usuario]": usuario, "data[Usuario][pass]": clave}
    try:
        respuesta = st.session_state.sesion_sigof.post(url_login, data=payload)
        if "dashboard/modulos" in respuesta.url or respuesta.status_code == 200:
            if "Módulo de Lectura" in respuesta.text or "Módulo de Entrega" in respuesta.text:
                 return True
            elif "dashboard/modulos" in respuesta.url:
                 return True
        return False
    except Exception as e:
        st.error(f"Error de conexion: {e}")
        return False

def limpiar_dato(valor):
    if pd.isna(valor): 
        return ""
    v = str(valor).strip()
    return v[:-2] if v.endswith(".0") else v

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
                st.error("Credenciales incorrectas o servidor inactivo.")
        else:
            st.warning("Por favor ingresa usuario y contraseña.")

else:
    st.success("Sesion iniciada en SIGOF. Conexion lista.")
    
    if st.button("Cerrar Sesion"):
        st.session_state.autenticado = False
        st.session_state.sesion_sigof.cookies.clear()
        st.rerun()
        
    st.divider()
    
    st.subheader("Modulo: Asignacion Masiva de Rutas")
    
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

    archivo_excel = st.file_uploader("Sube el archivo Excel", type=["xlsx"])

    if archivo_excel is not None:
        df = pd.read_excel(archivo_excel)
        columnas_excel = [str(col).strip().upper() for col in df.columns]
        
        has_ciclo = any("CICLO" in col for col in columnas_excel)
        has_sector = "SECTOR" in columnas_excel
        has_ruta = "RUTA REP." in columnas_excel
        has_id = "ID" in columnas_excel
        
        if has_ciclo and has_sector and has_ruta and has_id:
            col1, col2 = st.columns(2)
            btn_analizar = col1.button("Analizar Estado de Rutas")
            btn_asignar = col2.button("Iniciar Asignacion Masiva")
            
            url_leer = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listacreatelibro2"
            url_guardar = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_guardarlecturistalibro"
            
            headers_ajax = {
                "X-Requested-With": "XMLHttpRequest",
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "Referer": "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion",
                "Origin": "http://sigof.distriluz.com.pe"
            }
            
            total_filas = len(df)
            
            def extraer_datos_fila(fila):
                c, s, r, i = "", "", "", ""
                for col_name in df.columns:
                    col_upper = str(col_name).strip().upper()
                    val = limpiar_dato(fila[col_name])
                    if "CICLO" in col_upper: c = val
                    elif col_upper == "SECTOR": s = val
                    elif col_upper == "RUTA REP.": r = val
                    elif col_upper == "ID": i = val
                return c, s, r, i

            if btn_analizar:
                st.info("Revisando el estado real de las rutas en el sistema...")
                barra_progreso = st.progress(0)
                reporte_analisis = []
                registro_consola = []
                
                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    
                    if not ciclo or not sector or not ruta:
                        estado = "[ RECHAZADO ] Faltan datos."
                        registro_consola.append(f"Fila {index+1}: Celdas vacias en el Excel.")
                    else:
                        payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                        try:
                            res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer, headers=headers_ajax)
                            
                            if "login" in res_leer.url:
                                estado = "[ SESION CADUCADA ]"
                                registro_consola.append(f"Fila {index+1} (Ruta {ruta}): SIGOF redirecciono al Login.")
                            else:
                                try:
                                    datos = res_leer.json()
                                    if "aaData" in datos and len(datos["aaData"]) > 0:
                                        lecturista_actual = str(datos["aaData"][0][5]).strip()
                                        if lecturista_actual and lecturista_actual != "None" and lecturista_actual != "":
                                            estado = f"[ OCUPADA ] Asignada a: {lecturista_actual}."
                                        elif not id_lec:
                                            estado = "[ FALTA LECTURISTA ]"
                                        else:
                                            estado = "[ LISTA ] Ruta libre y lista."
                                    else:
                                        estado = f"[ COMBINACION INVALIDA ] SIGOF devolvio datos vacios."
                                        registro_consola.append(f"Fila {index+1} (Ruta {ruta}) | Status: {res_leer.status_code} | Enviado: {payload_leer} | Respuesta de SIGOF: {res_leer.text}")
                                except Exception as json_err:
                                    estado = "[ ERROR DE FORMATO ] El servidor no devolvio JSON."
                                    registro_consola.append(f"Fila {index+1} (Ruta {ruta}) | Status: {res_leer.status_code} | Enviado: {payload_leer} | Respuesta cruda de SIGOF: {res_leer.text}")
                        except Exception as req_err:
                            estado = "[ ERROR SERVIDOR ] Fallo la conexion."
                            registro_consola.append(f"Fila {index+1} (Ruta {ruta}): Error de red -> {req_err}")
                    
                    reporte_analisis.append({"Ciclo": ciclo, "Sector": sector, "Ruta": ruta, "ID Excel": id_lec, "Diagnostico": estado})
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.subheader("Resultado del Analisis")
                st.dataframe(pd.DataFrame(reporte_analisis), use_container_width=True)
                
                # Desplegable de la consola de registro
                with st.expander("Consola de Registro Técnico (Clic para ver detalles de los errores)"):
                    if registro_consola:
                        st.text_area("Log de Errores (Respuesta cruda del servidor):", "\n\n".join(registro_consola), height=300)
                    else:
                        st.write("No se detectaron errores de comunicación con el servidor.")

            if btn_asignar:
                st.info("Iniciando la asignacion real en el sistema...")
                barra_progreso = st.progress(0)
                log_resultados = []
                registro_consola = []
                
                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    
                    if not ciclo or not sector or not ruta:
                        log_resultados.append(f"Fila {index+1}: SALTADA. Falta Ciclo, Sector o Ruta en el Excel.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                        
                    if not id_lec:
                        log_resultados.append(f"Ruta {ruta}: SALTADA. Falta el ID del lecturista en el Excel.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                    
                    try:
                        res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer, headers=headers_ajax)
                        
                        if "login" in res_leer.url:
                            log_resultados.append(f"Ruta {ruta}: ERROR CRITICO. La sesion caduco, vuelve a ingresar.")
                            break
                        
                        try:    
                            datos = res_leer.json()
                            if "aaData" in datos and len(datos["aaData"]) > 0:
                                lecturista_actual = str(datos["aaData"][0][5]).strip()
                                
                                if lecturista_actual and lecturista_actual != "None" and lecturista_actual != "":
                                    log_resultados.append(f"Ruta {ruta}: IGNORADA. Ya esta asignada a {lecturista_actual}.")
                                else:
                                    suministro_inicio = datos["aaData"][0][2]
                                    suministro_fin = datos["aaData"][-1][2]
                                    
                                    payload_guardar = {
                                        "repartidor": id_lec,
                                        "ciclo": ciclo,
                                        "sector": sector,
                                        "ruta": ruta,
                                        "negocio": "82",
                                        "suministro_inicio": suministro_inicio,
                                        "suministro_fin": suministro_fin
                                    }
                                    
                                    res_guardar = st.session_state.sesion_sigof.post(url_guardar, data=payload_guardar, headers=headers_ajax)
                                    
                                    if res_guardar.status_code == 200:
                                        log_resultados.append(f"Ruta {ruta}: ASIGNADA con exito al ID {id_lec}.")
                                    else:
                                        log_resultados.append(f"Ruta {ruta}: ERROR. El servidor rechazo los datos (Status {res_guardar.status_code}).")
                                        registro_consola.append(f"Asignacion Ruta {ruta} | Payload Guardar: {payload_guardar} | Respuesta: {res_guardar.text}")
                            else:
                                log_resultados.append(f"Ruta {ruta}: IGNORADA. Combinacion Ciclo/Sector/Ruta invalida en el sistema.")
                                registro_consola.append(f"Fila {index+1} (Ruta {ruta}) | Status: {res_leer.status_code} | Enviado: {payload_leer} | Respuesta de SIGOF: {res_leer.text}")
                        except Exception as json_err:
                            log_resultados.append(f"Ruta {ruta}: ERROR DE FORMATO.")
                            registro_consola.append(f"Fila {index+1} (Ruta {ruta}) | Respuesta cruda de SIGOF: {res_leer.text}")
                            
                    except Exception as e:
                        log_resultados.append(f"Ruta {ruta}: ERROR DE RED. Fallo la conexion.")
                        registro_consola.append(f"Fila {index+1} (Ruta {ruta}): Error de red -> {e}")
                        
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.divider()
                st.subheader("Resumen Final de Asignacion")
                for log in log_resultados:
                    if "ASIGNADA" in log:
                        st.success(log)
                    elif "ERROR" in log:
                        st.error(log)
                    else:
                        st.warning(log)
                        
                with st.expander("Consola de Registro Técnico (Detalles de errores de asignación)"):
                    if registro_consola:
                        st.text_area("Log de Errores (Respuesta cruda del servidor):", "\n\n".join(registro_consola), height=300)
                    else:
                        st.write("No se detectaron errores de comunicación con el servidor durante la asignación.")
                    
        else:
            st.error("Tu Excel debe tener las columnas 'Ciclo', 'SECTOR', 'RUTA REP.' e 'ID'.")
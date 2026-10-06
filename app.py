import streamlit as st
import requests
import pandas as pd
import io

st.set_page_config(page_title="Herramientas SIGOF", layout="wide")

if 'sesion_sigof' not in st.session_state:
    st.session_state.sesion_sigof = requests.Session()
    st.session_state.sesion_sigof.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": "http://sigof.distriluz.com.pe",
        "Referer": "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion"
    })

if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def iniciar_sesion(usuario, clave):
    url_login = "http://sigof.distriluz.com.pe/plus/usuario/login"
    payload = {"data[Usuario][usuario]": usuario, "data[Usuario][pass]": clave}
    try:
        # 1. Petición de login
        respuesta_login = st.session_state.sesion_sigof.post(url_login, data=payload)
        
        # 2. Verificación de redirección y carga de módulo inicial
        if "dashboard/modulos" in respuesta_login.url or respuesta_login.status_code == 200:
            if "Módulo de Lectura" in respuesta_login.text or "Módulo de Entrega" in respuesta_login.text:
                 # 3. Inicialización crucial: Visitar la página del módulo para establecer el contexto de la sesión
                 st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion")
                 return True
            elif "dashboard/modulos" in respuesta_login.url:
                 # 3. Inicialización crucial: Visitar la página del módulo para establecer el contexto de la sesión
                 st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion")
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
            
            url_listar = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listar_ordenes_pendientes_asignacion"
            url_leer = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listacreatelibro2"
            url_guardar = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_guardarlecturistalibro"
            
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
                st.info("Analizando rutas...")
                barra_progreso = st.progress(0)
                reporte_analisis = []
                registro_consola = []
                
                ultimo_ciclo = None
                ultimo_sector = None
                
                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    
                    if not ciclo or not sector or not ruta:
                        estado = "[ RECHAZADO ] Faltan datos clave."
                    else:
                        try:
                            # 1. Sincronización del contexto del servidor antes de cada consulta
                            if ciclo != ultimo_ciclo or sector != ultimo_sector:
                                payload_previa = {"ciclo": ciclo, "sector": sector, "id_tipo_reparto": "N"}
                                st.session_state.sesion_sigof.post(url_listar, data=payload_previa)
                                ultimo_ciclo = ciclo
                                ultimo_sector = sector
                            
                            # 2. Petición para recuperar datos de la ruta
                            payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                            res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                            
                            if "login" in res_leer.url:
                                estado = "[ SESION CADUCADA ]"
                                break
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
                                            estado = "[ LISTA ] Ruta libre."
                                    else:
                                        estado = f"[ COMBINACION INVALIDA ] Ruta sin items en este ciclo."
                                        registro_consola.append(f"Fila {index+1} (Ruta {ruta}) | Respuesta de SIGOF: {res_leer.text}")
                                except Exception as json_err:
                                     estado = "[ ERROR LECTURA JSON ]"
                                     registro_consola.append(f"Fila {index+1} (Ruta {ruta}): {json_err} | Respuesta de SIGOF: {res_leer.text}")
                        except Exception as e:
                            estado = "[ ERROR ] Fallo la peticion."
                            registro_consola.append(f"Fila {index+1} (Ruta {ruta}): {str(e)}")
                    
                    reporte_analisis.append({"Ciclo": ciclo, "Sector": sector, "Ruta": ruta, "ID Excel": id_lec, "Diagnostico": estado})
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.subheader("Resultado del Analisis")
                st.dataframe(pd.DataFrame(reporte_analisis), use_container_width=True)
                
                with st.expander("Consola de Registro Técnico (Detalles)"):
                    if registro_consola:
                        st.text_area("Log:", "\n\n".join(registro_consola), height=300)
                    else:
                        st.write("Analisis limpio.")

            if btn_asignar:
                st.info("Iniciando asignacion masiva...")
                barra_progreso = st.progress(0)
                log_resultados = []
                
                ultimo_ciclo = None
                ultimo_sector = None
                
                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    
                    if not ciclo or not sector or not ruta or not id_lec:
                        log_resultados.append(f"Fila {index+1}: SALTADA por falta de datos.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    try:
                        # 1. Sincronización del contexto del servidor antes de cada consulta
                        if ciclo != ultimo_ciclo or sector != ultimo_sector:
                            payload_previa = {"ciclo": ciclo, "sector": sector, "id_tipo_reparto": "N"}
                            st.session_state.sesion_sigof.post(url_listar, data=payload_previa)
                            ultimo_ciclo = ciclo
                            ultimo_sector = sector
                            
                        # 2. Leer limites
                        payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                        res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                        
                        if "login" in res_leer.url:
                            log_resultados.append(f"Ruta {ruta}: ERROR CRITICO. La sesion caduco.")
                            break
                        
                        try:
                            datos = res_leer.json()
                            if "aaData" in datos and len(datos["aaData"]) > 0:
                                lecturista_actual = str(datos["aaData"][0][5]).strip()
                                
                                if lecturista_actual and lecturista_actual != "None" and lecturista_actual != "":
                                    log_resultados.append(f"Ruta {ruta}: IGNORADA. Ya asignada a {lecturista_actual}.")
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
                                    
                                    res_guardar = st.session_state.sesion_sigof.post(url_guardar, data=payload_guardar)
                                    
                                    if res_guardar.status_code == 200:
                                        log_resultados.append(f"Ruta {ruta}: ASIGNADA con exito al ID {id_lec}.")
                                    else:
                                        log_resultados.append(f"Ruta {ruta}: ERROR. Rechazada al guardar.")
                            else:
                                log_resultados.append(f"Ruta {ruta}: IGNORADA. Sin datos en este ciclo.")
                        except Exception:
                            log_resultados.append(f"Ruta {ruta}: ERROR DE LECTURA JSON.")
                    except Exception:
                        log_resultados.append(f"Ruta {ruta}: ERROR DE RED.")
                        
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
                        
        else:
            st.error("Tu Excel debe tener las columnas 'Ciclo', 'SECTOR', 'RUTA REP.' e 'ID'.")
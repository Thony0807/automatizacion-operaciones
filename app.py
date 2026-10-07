import streamlit as st
import requests
import pandas as pd
import io
import re

st.set_page_config(page_title="Herramientas SIGOF", layout="wide")

if 'sesion_sigof' not in st.session_state:
    st.session_state.sesion_sigof = requests.Session()
    st.session_state.sesion_sigof.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "X-Requested-With": "XMLHttpRequest"
    })

if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def iniciar_sesion(usuario, clave):
    url_login = "http://sigof.distriluz.com.pe/plus/usuario/login"
    payload = {"data[Usuario][usuario]": usuario, "data[Usuario][pass]": clave}
    try:
        # 1. Cargar la página inicial para obtener las cookies
        st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/dashboard/init")
        
        # 2. Iniciar Sesión
        respuesta = st.session_state.sesion_sigof.post(url_login, data=payload)
        
        # 3. PASO CLAVE DESCUBIERTO EN EL HAR: Fijar la sesión en Huánuco (UUNN 82, Empresa 4)
        url_cambiar_sesion = "http://sigof.distriluz.com.pe/plus/usuario/ajax_cambiar_sesion"
        st.session_state.sesion_sigof.post(url_cambiar_sesion, data={'idempresa': '4', 'iduunn': '82'})
        
        if "dashboard" in respuesta.url or respuesta.status_code == 200:
            return True
        return False
    except Exception as e:
        st.error(f"Error de red: {e}")
        return False

def limpiar_dato(valor):
    if pd.isna(valor): 
        return ""
    v = str(valor).strip()
    return v[:-2] if v.endswith(".0") else v

def limpiar_html(raw_html):
    # Elimina las etiquetas HTML ocultas que envuelve SIGOF para sacar solo los números
    cleanr = re.compile('<.*?>')
    return re.sub(cleanr, '', str(raw_html)).strip()

st.title("Sistema de Automatizacion Masiva")

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
    st.success("Sesion iniciada. Conexion fijada en la sede Huánuco.")
    
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

            def crear_payload_lectura(c, s, r):
                return {
                    "ciclo": str(c), "sector": str(s), "ruta": str(r), "id_tipo_reparto": "N",
                    "sEcho": "1", "iColumns": "8", "sColumns": "",
                    "iDisplayStart": "0", "iDisplayLength": "5000",
                    "mDataProp_0": "0", "mDataProp_1": "1", "mDataProp_2": "2", "mDataProp_3": "3",
                    "mDataProp_4": "4", "mDataProp_5": "5", "mDataProp_6": "6", "mDataProp_7": "7"
                }

            if btn_analizar:
                st.info("Analizando rutas en Huánuco...")
                barra_progreso = st.progress(0)
                reporte_analisis = []
                
                try:
                    st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion")
                except:
                    pass

                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    
                    faltantes = []
                    if not ciclo: faltantes.append("Ciclo")
                    if not sector: faltantes.append("Sector")
                    if not ruta: faltantes.append("Ruta")
                    if not id_lec: faltantes.append("ID Lecturista")
                    
                    if len(faltantes) > 0:
                        estado = f"[ DATOS FALTANTES ] Falta en el Excel: {', '.join(faltantes)}."
                    else:
                        payload_leer = crear_payload_lectura(ciclo, sector, ruta)
                        try:
                            res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                            
                            if "login" in res_leer.url:
                                estado = "[ SESION CADUCADA ] Vuelve a ingresar."
                                break
                            elif not res_leer.text.strip():
                                estado = "[ SIN RESPUESTA ] Servidor vacío."
                            else:
                                try:
                                    datos = res_leer.json()
                                    if "aaData" in datos and len(datos["aaData"]) > 0:
                                        lecturista_actual = limpiar_html(datos["aaData"][0][5])
                                        if lecturista_actual and lecturista_actual.lower() != "none" and lecturista_actual != "":
                                            estado = f"[ OCUPADA ] Ya asignada a: {lecturista_actual}."
                                        else:
                                            estado = "[ LISTA ] Ruta libre para asignar."
                                    else:
                                        estado = "[ VACÍA ] No hay recibos para esta ruta."
                                except ValueError:
                                    estado = "[ ERROR INTERNO ] SIGOF devolvió datos ilegibles."
                        except Exception:
                            estado = "[ ERROR DE RED ] Falló la conexión."
                    
                    reporte_analisis.append({"Ciclo": ciclo, "Sector": sector, "Ruta": ruta, "ID Excel": id_lec, "Diagnostico": estado})
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.subheader("Resultado del Analisis")
                st.dataframe(pd.DataFrame(reporte_analisis), use_container_width=True)

            if btn_asignar:
                st.info("Iniciando asignacion masiva...")
                barra_progreso = st.progress(0)
                log_resultados = []
                
                try:
                    st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion")
                except:
                    pass

                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    
                    faltantes = []
                    if not ciclo: faltantes.append("Ciclo")
                    if not sector: faltantes.append("Sector")
                    if not ruta: faltantes.append("Ruta")
                    if not id_lec: faltantes.append("ID Lecturista")
                    
                    if len(faltantes) > 0:
                        log_resultados.append(f"Fila {index+1}: SALTADA. Falta: {', '.join(faltantes)}.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    payload_leer = crear_payload_lectura(ciclo, sector, ruta)
                    
                    try:
                        res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                        
                        if "login" in res_leer.url:
                            log_resultados.append("ERROR CRÍTICO: Tu sesión caducó.")
                            break
                        
                        try:
                            datos = res_leer.json()
                            if "aaData" in datos and len(datos["aaData"]) > 0:
                                lecturista_actual = limpiar_html(datos["aaData"][0][5])
                                
                                if lecturista_actual and lecturista_actual.lower() != "none" and lecturista_actual != "":
                                    log_resultados.append(f"Ruta {ruta}: IGNORADA. Ya asignada a {lecturista_actual}.")
                                else:
                                    item_inicio = limpiar_html(datos["aaData"][0][0])
                                    item_fin = limpiar_html(datos["aaData"][-1][0])
                                    
                                    payload_guardar = {
                                        "negocio": "82",
                                        "ciclo": str(ciclo),
                                        "sector": str(sector),
                                        "rutas": str(ruta),
                                        "desde": str(item_inicio),
                                        "hasta": str(item_fin),
                                        "lecturista": str(id_lec)
                                    }
                                    
                                    res_guardar = st.session_state.sesion_sigof.post(url_guardar, data=payload_guardar)
                                    
                                    if res_guardar.status_code == 200:
                                        log_resultados.append(f"Ruta {ruta}: ASIGNADA con éxito al ID {id_lec}.")
                                    else:
                                        log_resultados.append(f"Ruta {ruta}: ERROR al intentar guardar en SIGOF.")
                            else:
                                log_resultados.append(f"Ruta {ruta}: IGNORADA. No existen recibos cargados.")
                        except ValueError:
                            log_resultados.append(f"Ruta {ruta}: ERROR INTERNO de SIGOF al leer la tabla.")
                    except Exception:
                        log_resultados.append(f"Ruta {ruta}: ERROR DE RED. Revisa tu conexión.")
                        
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.divider()
                st.subheader("Resumen Final de Asignacion")
                for log in log_resultados:
                    if "ASIGNADA" in log:
                        st.success(log)
                    elif "ERROR" in log or "SALTADA" in log:
                        st.error(log)
                    else:
                        st.warning(log)
                        
        else:
            st.error("Tu Excel debe tener las columnas 'Ciclo', 'SECTOR', 'RUTA REP.' e 'ID'.")
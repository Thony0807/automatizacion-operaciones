import streamlit as st
import requests
import pandas as pd
import io
import re
import time

# ==========================================
# CONFIGURACION GLOBAL Y VARIABLES DE SESION
# ==========================================
st.set_page_config(page_title="Herramientas de Automatizacion", layout="wide")

# Sesion SIGOF
if 'sesion_sigof' not in st.session_state:
    st.session_state.sesion_sigof = requests.Session()
    st.session_state.sesion_sigof.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Origin": "http://sigof.distriluz.com.pe",
        "Referer": "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion"
    })

if 'autenticado_sigof' not in st.session_state:
    st.session_state.autenticado_sigof = False

# Sesion Field Service
if 'sesion_field' not in st.session_state:
    st.session_state.sesion_field = requests.Session()
    
if 'autenticado_field' not in st.session_state:
    st.session_state.autenticado_field = False

# ==========================================
# FUNCIONES DE UTILIDAD Y CONEXION
# ==========================================
def iniciar_sesion_sigof(usuario, clave):
    url_login = "http://sigof.distriluz.com.pe/plus/usuario/login"
    payload = {"data[Usuario][usuario]": usuario, "data[Usuario][pass]": clave}
    try:
        st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/dashboard/init")
        respuesta = st.session_state.sesion_sigof.post(url_login, data=payload)
        
        url_cambiar_sesion = "http://sigof.distriluz.com.pe/plus/usuario/ajax_cambiar_sesion"
        st.session_state.sesion_sigof.post(url_cambiar_sesion, data={'idempresa': '4', 'iduunn': '82'})
        
        if "dashboard" in respuesta.url or respuesta.status_code == 200:
            return True
        return False
    except Exception as e:
        st.error(f"Error de red SIGOF: {e}")
        return False

def iniciar_sesion_field(usuario, clave):
    # PLANTILLA: Aqui ira el codigo exacto de login para Field Service
    try:
        time.sleep(1) 
        if usuario and clave:
            return True
        return False
    except Exception as e:
        st.error(f"Error de red Field Service: {e}")
        return False

def extraer_numero(valor):
    if pd.isna(valor): return ""
    nums = re.findall(r'\d+', str(valor))
    return nums[0] if nums else "0"

def limpiar_html(raw_html):
    cleanr = re.compile('<.*?>')
    texto_limpio = re.sub(cleanr, '', str(raw_html)).strip()
    return extraer_numero(texto_limpio)

def crear_payload_lectura(c, s, r):
    return {
        "ciclo": str(c), "sector": str(s), "ruta": str(r), "id_tipo_reparto": "N",
        "sEcho": "1", "iColumns": "8", "sColumns": "",
        "iDisplayStart": "0", "iDisplayLength": "5000",
        "mDataProp_0": "0", "mDataProp_1": "1", "mDataProp_2": "2", "mDataProp_3": "3",
        "mDataProp_4": "4", "mDataProp_5": "5", "mDataProp_6": "6", "mDataProp_7": "7"
    }

# ==========================================
# MODULOS DE TRABAJO
# ==========================================

def modulo_asignacion_masiva_sigof():
    st.header("Asignacion Masiva de Rutas (SIGOF)")
    
    # --- BLOQUE DE AUTENTICACION DEL MODULO ---
    if not st.session_state.autenticado_sigof:
        st.subheader("Acceso a SIGOF requerido")
        st.info("Para utilizar este modulo, por favor inicia sesion con tus credenciales de SIGOF.")
        usu_sigof = st.text_input("Usuario SIGOF", key="login_usu_sigof")
        cla_sigof = st.text_input("Contrasena SIGOF", type="password", key="login_cla_sigof")
        
        if st.button("Ingresar a SIGOF"):
            if usu_sigof and cla_sigof:
                if iniciar_sesion_sigof(usu_sigof, cla_sigof):
                    st.session_state.autenticado_sigof = True
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas o servidor inactivo.")
            else:
                st.warning("Por favor ingresa usuario y contrasena.")
        return # Detiene la ejecucion del resto del modulo si no esta logueado

    # --- INTERFAZ DEL MODULO (Si esta logueado) ---
    col_estado, col_btn = st.columns([3, 1])
    with col_estado:
        st.success("[ OK ] Sesion activa en SIGOF (Sede Huanuco).")
    with col_btn:
        if st.button("Cerrar Sesion SIGOF", use_container_width=True):
            st.session_state.autenticado_sigof = False
            st.session_state.sesion_sigof.cookies.clear()
            st.rerun()
            
    st.divider()

    columnas_requeridas = ['Ciclo', 'SECTOR', 'RUTA REP.', 'ID']
    df_plantilla = pd.DataFrame(columns=columnas_requeridas)
    
    buffer_plantilla = io.BytesIO()
    with pd.ExcelWriter(buffer_plantilla, engine='openpyxl') as writer:
        df_plantilla.to_excel(writer, index=False, sheet_name='Cronograma')
    
    st.download_button(
        label="Descargar Plantilla Excel",
        data=buffer_plantilla.getvalue(),
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
                    val = extraer_numero(fila[col_name])
                    if "CICLO" in col_upper: c = val
                    elif col_upper == "SECTOR": s = val
                    elif col_upper == "RUTA REP.": r = val
                    elif col_upper == "ID": i = val
                return c, s, r, i

            if btn_analizar:
                st.info("Analizando rutas en SIGOF...")
                barra_progreso = st.progress(0)
                reporte_analisis = []
                
                try:
                    st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion")
                except:
                    pass

                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    
                    faltantes = []
                    if not ciclo or ciclo == "0": faltantes.append("Ciclo")
                    if not sector or sector == "0": faltantes.append("Sector")
                    if not ruta or ruta == "0": faltantes.append("Ruta")
                    if not id_lec or id_lec == "0": faltantes.append("ID Lecturista")
                    
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
                                estado = "[ SIN RESPUESTA ] Servidor vacio."
                            else:
                                try:
                                    datos = res_leer.json()
                                    if "aaData" in datos and len(datos["aaData"]) > 0:
                                        cleanr = re.compile('<.*?>')
                                        lecturista_actual = re.sub(cleanr, '', str(datos["aaData"][0][5])).strip()
                                        if lecturista_actual and lecturista_actual.lower() != "none" and lecturista_actual != "":
                                            estado = f"[ OCUPADA ] Ya asignada a: {lecturista_actual}."
                                        else:
                                            estado = "[ LISTA ] Ruta libre para asignar."
                                    else:
                                        estado = "[ VACIA ] No hay recibos para esta ruta."
                                except ValueError:
                                    estado = "[ ERROR INTERNO ] SIGOF devolvio datos ilegibles."
                        except Exception:
                            estado = "[ ERROR DE RED ] Fallo la conexion."
                    
                    reporte_analisis.append({"Ciclo": ciclo, "Sector": sector, "Ruta": ruta, "ID Excel": id_lec, "Diagnostico": estado})
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.subheader("Resultado del Analisis")
                st.dataframe(pd.DataFrame(reporte_analisis), use_container_width=True)

            if btn_asignar:
                st.info("Iniciando asignacion masiva en SIGOF...")
                barra_progreso = st.progress(0)
                reporte_asignacion = []
                
                try:
                    st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/listar_asignacion")
                except:
                    pass

                for index, fila in df.iterrows():
                    ciclo, sector, ruta, id_lec = extraer_datos_fila(fila)
                    estado_final = ""
                    detalle_final = ""
                    
                    faltantes = []
                    if not ciclo or ciclo == "0": faltantes.append("Ciclo")
                    if not sector or sector == "0": faltantes.append("Sector")
                    if not ruta or ruta == "0": faltantes.append("Ruta")
                    if not id_lec or id_lec == "0": faltantes.append("ID Lecturista")
                    
                    if len(faltantes) > 0:
                        estado_final = "NO ASIGNADO"
                        detalle_final = f"Datos faltantes en Excel: {', '.join(faltantes)}."
                    else:
                        payload_leer = crear_payload_lectura(ciclo, sector, ruta)
                        
                        try:
                            res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                            
                            if "login" in res_leer.url:
                                estado_final = "ERROR CRITICO"
                                detalle_final = "La sesion en SIGOF ha caducado."
                                reporte_asignacion.append({"Ciclo": ciclo, "SECTOR": sector, "RUTA REP.": ruta, "ID": id_lec, "Estado Final": estado_final, "Detalle": detalle_final})
                                break
                            
                            try:
                                datos = res_leer.json()
                                if "aaData" in datos and len(datos["aaData"]) > 0:
                                    cleanr = re.compile('<.*?>')
                                    lecturista_actual = re.sub(cleanr, '', str(datos["aaData"][0][5])).strip()
                                    
                                    if lecturista_actual and lecturista_actual.lower() != "none" and lecturista_actual != "":
                                        estado_final = "IGNORADO"
                                        detalle_final = f"Ruta ya asignada a: {lecturista_actual}."
                                    else:
                                        suministro_inicio = limpiar_html(datos["aaData"][0][2])
                                        suministro_fin = limpiar_html(datos["aaData"][-1][2])
                                        
                                        payload_guardar = {
                                            "repartidor": str(id_lec),
                                            "ciclo": str(ciclo),
                                            "sector": str(sector),
                                            "ruta": str(ruta),
                                            "negocio": "82",
                                            "suministro_inicio": str(suministro_inicio),
                                            "suministro_fin": str(suministro_fin)
                                        }
                                        
                                        res_guardar = st.session_state.sesion_sigof.post(url_guardar, data=payload_guardar)
                                        
                                        if res_guardar.status_code == 200:
                                            estado_final = "ASIGNADO"
                                            detalle_final = f"Ruta asignada al ID {id_lec}."
                                        else:
                                            estado_final = "ERROR DE GUARDADO"
                                            detalle_final = f"SIGOF rechazo la peticion. Codigo: {res_guardar.status_code}."
                                else:
                                    estado_final = "NO ASIGNADO"
                                    detalle_final = "La ruta esta VACIA (Sin recibos)."
                            except ValueError:
                                estado_final = "ERROR INTERNO"
                                detalle_final = "Fallo en la lectura JSON desde SIGOF."
                        except Exception:
                            estado_final = "ERROR DE RED"
                            detalle_final = "Se perdio la conexion."
                            
                    reporte_asignacion.append({
                        "Ciclo": ciclo, 
                        "SECTOR": sector, 
                        "RUTA REP.": ruta, 
                        "ID": id_lec, 
                        "Estado Final": estado_final, 
                        "Detalle": detalle_final
                    })
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.divider()
                st.subheader("Resumen Final de Asignacion")
                
                df_resultado = pd.DataFrame(reporte_asignacion)
                st.dataframe(df_resultado, use_container_width=True)
                
                buffer_resultado = io.BytesIO()
                with pd.ExcelWriter(buffer_resultado, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name='Resultados')
                
                st.download_button(
                    label="Descargar Reporte Final en Excel",
                    data=buffer_resultado.getvalue(),
                    file_name="Reporte_Asignacion_SIGOF.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        else:
            st.error("Tu Excel debe tener las columnas 'Ciclo', 'SECTOR', 'RUTA REP.' e 'ID'.")

def modulo_plantilla_field_service():
    st.header("Validacion de Inspecciones (Field Service)")
    
    # --- BLOQUE DE AUTENTICACION DEL MODULO ---
    if not st.session_state.autenticado_field:
        st.subheader("Acceso a Field Service requerido")
        st.info("Para utilizar este modulo, por favor inicia sesion con tus credenciales de Field Service.")
        usu_field = st.text_input("Usuario Field Service", key="login_usu_field")
        cla_field = st.text_input("Contrasena Field Service", type="password", key="login_cla_field")
        
        if st.button("Ingresar a Field Service"):
            if usu_field and cla_field:
                if iniciar_sesion_field(usu_field, cla_field):
                    st.session_state.autenticado_field = True
                    st.rerun()
                else:
                    st.error("Error al conectar con Field Service.")
            else:
                st.warning("Por favor ingresa usuario y contrasena.")
        return # Detiene la ejecucion del resto del modulo si no esta logueado
        
    # --- INTERFAZ DEL MODULO (Si esta logueado) ---
    col_estado, col_btn = st.columns([3, 1])
    with col_estado:
        st.success("[ OK ] Sesion activa en Field Service.")
    with col_btn:
        if st.button("Cerrar Sesion Field", use_container_width=True):
            st.session_state.autenticado_field = False
            st.session_state.sesion_field.cookies.clear()
            st.rerun()
            
    st.divider()
    st.info("Aqui se programara la logica del modulo asociado al sistema Field Service una vez extraigamos sus datos de red.")

# ==========================================
# RUTADOR PRINCIPAL (SIDEBAR)
# ==========================================
st.sidebar.title("Navegacion de Modulos")

# Indicadores de estado en el panel lateral
estado_sigof = "Conectado" if st.session_state.autenticado_sigof else "Desconectado"
estado_field = "Conectado" if st.session_state.autenticado_field else "Desconectado"
st.sidebar.markdown(f"**SIGOF:** {estado_sigof}")
st.sidebar.markdown(f"**Field Service:** {estado_field}")
st.sidebar.divider()

modulo_seleccionado = st.sidebar.radio(
    "Selecciona una herramienta:",
    (
        "Asignacion Masiva de Rutas (SIGOF)", 
        "Validacion de Inspecciones (Field Service)"
    )
)

# Renderizado dinamico segun la seleccion
if modulo_seleccionado == "Asignacion Masiva de Rutas (SIGOF)":
    modulo_asignacion_masiva_sigof()
elif modulo_seleccionado == "Validacion de Inspecciones (Field Service)":
    modulo_plantilla_field_service()
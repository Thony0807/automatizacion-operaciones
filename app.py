import streamlit as st
import requests
import pandas as pd
import io

# Configuracion basica
st.set_page_config(page_title="Herramientas SIGOF", layout="centered")

if 'sesion_sigof' not in st.session_state:
    st.session_state.sesion_sigof = requests.Session()

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

# Funcion para leer el Excel sin alterar los numeros
def limpiar_dato(valor):
    if pd.isna(valor): 
        return ""
    v = str(valor).strip()
    # Si Pandas le pone un ".0" al final de un numero entero, se lo quitamos
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
    st.write("Recuerda: En la columna 'Ciclo' debes poner el codigo interno de SIGOF (ej. 6157), no la fecha.")
    
    columnas_requeridas = ['Ciclo', 'SECTOR', 'RUTA REP.', 'ID']
    df_plantilla = pd.DataFrame(columns=columnas_requeridas)
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_plantilla.to_excel(writer, index=False, sheet_name='Cronograma')
    
    st.download_button(
        label="Descargar Plantilla Excel",
        data=buffer.getvalue(),
        file_name="Plantilla_Asignacion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    archivo_excel = st.file_uploader("Sube el archivo Excel", type=["xlsx"])

    if archivo_excel is not None:
        df = pd.read_excel(archivo_excel)
        columnas_excel = [str(col).strip().upper() for col in df.columns]
        columnas_req_upper = [col.upper() for col in columnas_requeridas]
        
        if all(col in columnas_excel for col in columnas_req_upper):
            col1, col2 = st.columns(2)
            btn_analizar = col1.button("Analizar Estado de Rutas")
            btn_asignar = col2.button("Iniciar Asignacion Masiva")
            
            url_leer = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_listacreatelibro2"
            url_guardar = "http://sigof.distriluz.com.pe/plus/ComrepOrdenrepartos/ajax_guardarlecturistalibro"
            total_filas = len(df)
            
            # --- FASE DE ANALISIS ---
            if btn_analizar:
                st.info("Revisando como estan las rutas en el sistema...")
                barra_progreso = st.progress(0)
                reporte_analisis = []
                
                for index, fila in df.iterrows():
                    ciclo = limpiar_dato(fila['Ciclo'])
                    sector = limpiar_dato(fila['SECTOR'])
                    ruta = limpiar_dato(fila['RUTA REP.'])
                    id_lec = limpiar_dato(fila['ID'])
                    
                    if not ciclo or not sector or not ruta:
                        estado = "[ FALTAN DATOS ] Revisa que tu Excel tenga el Ciclo, Sector y Ruta escritos."
                    elif not id_lec:
                        estado = "[ SIN LECTURISTA ] Falta colocar el ID del trabajador en el Excel."
                    else:
                        payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                        try:
                            res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                            datos = res_leer.json()
                            
                            if "aaData" in datos and len(datos["aaData"]) > 0:
                                lecturista_actual = str(datos["aaData"][0][5]).strip()
                                if lecturista_actual and lecturista_actual != "None" and lecturista_actual != "":
                                    estado = f"[ OCUPADA ] Ya esta asignada a: {lecturista_actual}."
                                else:
                                    estado = "[ LISTA ] Ruta libre y lista para asignar."
                            else:
                                estado = "[ NO EXISTE ] El sistema no encuentra esta ruta. Revisa el Ciclo, Sector y Ruta."
                        except Exception:
                            estado = "[ ERROR ] El servidor no respondio al consultar esta ruta."
                    
                    reporte_analisis.append({"Ruta": ruta, "Sector": sector, "ID": id_lec, "Diagnostico": estado})
                    barra_progreso.progress((index + 1) / total_filas)
                
                st.subheader("Resultado del Analisis")
                st.dataframe(pd.DataFrame(reporte_analisis), use_container_width=True)

            # --- FASE DE ASIGNACION ---
            if btn_asignar:
                st.info("Iniciando la asignacion real en el sistema...")
                barra_progreso = st.progress(0)
                log_resultados = []
                
                for index, fila in df.iterrows():
                    ciclo = limpiar_dato(fila['Ciclo'])
                    sector = limpiar_dato(fila['SECTOR'])
                    ruta = limpiar_dato(fila['RUTA REP.'])
                    id_lec = limpiar_dato(fila['ID'])
                    
                    if not ciclo or not sector or not ruta:
                        log_resultados.append(f"Fila {index+1}: SALTADA. Le falta Ciclo, Sector o Ruta.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                        
                    if not id_lec:
                        log_resultados.append(f"Ruta {ruta}: SALTADA. No pusiste el ID del lecturista en el Excel.")
                        barra_progreso.progress((index + 1) / total_filas)
                        continue
                    
                    payload_leer = {"ciclo": ciclo, "sector": sector, "ruta": ruta, "id_tipo_reparto": "N"}
                    
                    try:
                        res_leer = st.session_state.sesion_sigof.post(url_leer, data=payload_leer)
                        datos = res_leer.json()
                        
                        if "aaData" in datos and len(datos["aaData"]) > 0:
                            lecturista_actual = str(datos["aaData"][0][5]).strip()
                            
                            if lecturista_actual and lecturista_actual != "None" and lecturista_actual != "":
                                log_resultados.append(f"Ruta {ruta}: IGNORADA. Ya la tiene {lecturista_actual}.")
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
                                    log_resultados.append(f"Ruta {ruta}: ERROR. El sistema rechazo los datos al guardar.")
                        else:
                            log_resultados.append(f"Ruta {ruta}: IGNORADA. No existe en el sistema.")
                            
                    except Exception:
                        log_resultados.append(f"Ruta {ruta}: ERROR DE RED. Fallo la conexion con el servidor.")
                    
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
            st.error(f"Tu Excel no tiene las columnas correctas. Descarga la plantilla y usala.")
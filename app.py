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
        st.session_state.sesion_sigof.get("http://sigof.distriluz.com.pe/plus/dashboard/init")
        respuesta = st.session_state.sesion_sigof.post(url_login, data=payload)
        
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
    st.success("[ OK ] Sesion iniciada. Conexion fijada en la sede Huanuco.")
    
    if st.button("Cerrar Sesion"):
        st.session_state.autenticado = False
        st.session_state.sesion_sigof.cookies.clear()
        st.rerun()
        
    st.divider()
    
    st.subheader("Modulo: Asignacion Masiva de Rutas")
    
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
                                estado = "[ SIN RESPUESTA ] Servidor vacio."
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
                st.info("Iniciando asignacion masiva...")
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
                    if not ciclo: faltantes.append("Ciclo")
                    if not sector: faltantes.append("Sector")
                    if not ruta: faltantes.append("Ruta")
                    if not id_lec: faltantes.append("ID Lecturista")
                    
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
                                break
                            
                            try:
                                datos = res_leer.json()
                                if "aaData" in datos and len(datos["aaData"]) > 0:
                                    lecturista_actual = limpiar_html(datos["aaData"][0][5])
                                    
                                    if lecturista_actual and lecturista_actual.lower() != "none" and lecturista_actual != "":
                                        estado_final = "IGNORADO"
                                        detalle_final = f"Ruta ya se encontraba asignada a: {lecturista_actual}."
                                    else:
                                        item_inicio = limpiar_html(datos["aaData"][0][0])
                                        item_fin = limpiar_html(datos["aaData"][-1][0])
                                        suministro_inicio = limpiar_html(datos["aaData"][0][2])
                                        suministro_fin = limpiar_html(datos["aaData"][-1][2])
                                        
                                        # Nombres exactos de las variables del formulario obtenidos del HAR
                                        payload_guardar = {
                                            "negocio": "82",
                                            "ciclo": str(ciclo),
                                            "sector": str(sector),
                                            "rutas": str(ruta),
                                            "desde": str(item_inicio),
                                            "hasta": str(item_fin),
                                            "lecturista": str(id_lec),
                                            "txt_suministro_inicio": str(suministro_inicio),
                                            "txt_suministro_fin": str(suministro_fin)
                                        }
                                        
                                        res_guardar = st.session_state.sesion_sigof.post(url_guardar, data=payload_guardar)
                                        
                                        if res_guardar.status_code == 200:
                                            estado_final = "ASIGNADO"
                                            detalle_final = f"Ruta guardada exitosamente para el ID {id_lec}."
                                        else:
                                            estado_final = "ERROR DE GUARDADO"
                                            detalle_final = f"SIGOF rechazo la peticion. Codigo: {res_guardar.status_code}."
                                else:
                                    estado_final = "NO ASIGNADO"
                                    detalle_final = "El servidor indica que la ruta esta VACIA (Sin recibos)."
                            except ValueError:
                                estado_final = "ERROR INTERNO"
                                detalle_final = "Fallo en la lectura de datos JSON desde SIGOF."
                        except Exception as e:
                            estado_final = "ERROR DE RED"
                            detalle_final = "Se perdio la conexion con el servidor."
                            
                    reporte_asignacion.append({
                        "Ciclo": ciclo, 
                        "Sector": sector, 
                        "Ruta": ruta, 
                        "ID Excel": id_lec, 
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
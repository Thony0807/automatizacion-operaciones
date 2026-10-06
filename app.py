import pandas as pd
import requests

# --- PEGAR ESTO DEBAJO DEL CODIGO DE LOGIN ---

st.subheader("Modulo: Asignacion Masiva de Rutas")

archivo_excel = st.file_uploader("Sube el cronograma en formato Excel (.xlsx)", type=["xlsx"])

if archivo_excel is not None:
    df = pd.read_excel(archivo_excel)
    
    # Actualizamos las columnas requeridas para incluir el 'ID'
    columnas_requeridas = ['Ciclo', 'SECTOR', 'RUTA REP.', 'NOMBRE COMPLETO', 'ID']
    
    # Validar que existan todas las columnas (ignorando mayusculas/minusculas y espacios extras)
    columnas_excel = [str(col).strip().upper() for col in df.columns]
    columnas_req_upper = [col.upper() for col in columnas_requeridas]
    
    if all(col in columnas_excel for col in columnas_req_upper):
        st.success("Archivo leido correctamente. Columnas validadas.")
        
        if st.button("Iniciar Asignacion Masiva"):
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
                    
                    # Extraer el ID asegurando que sea un numero entero limpio (sin decimales .0)
                    if pd.notna(fila['ID']):
                        id_lecturista = str(int(float(fila['ID'])))
                    else:
                        id_lecturista = ""
                except Exception as e:
                    log_resultados.append(f"[ERROR] Fila {index+1}: Datos invalidos o vacios. Detalle: {e}")
                    barra_progreso.progress((index + 1) / total_filas)
                    continue
                
                if not id_lecturista or id_lecturista == "":
                    log_resultados.append(f"[ERROR] Fila {index+1}: ID vacio para {nombre}")
                    barra_progreso.progress((index + 1) / total_filas)
                    continue
                
                # --- PASO 1: LEER LIMITES DE LA RUTA ---
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
                        
                        # Validar si ya tiene un lecturista asignado (indice 5 del JSON)
                        lecturista_actual = str(filas_ruta[0][5]).strip()
                        if lecturista_actual and lecturista_actual != "None" and lecturista_actual != "":
                            log_resultados.append(f"[OMITIDO] Ruta {ruta} ya esta asignada a: {lecturista_actual}")
                        else:
                            # Extraer suministro inicial y final
                            suministro_inicio = filas_ruta[0][2]
                            suministro_fin = filas_ruta[-1][2]
                            
                            # --- PASO 2: GUARDAR ASIGNACION ---
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
                        log_resultados.append(f"[ALERTA] Ruta {ruta} en Sector {sector} no fue encontrada en SIGOF o no tiene items.")
                        
                except Exception as e:
                    log_resultados.append(f"[ERROR] Fallo de conexion en Ruta {ruta}: {e}")
                
                # Actualizar barra de progreso
                barra_progreso.progress((index + 1) / total_filas)
            
            # Mostrar resumen al finalizar
            st.divider()
            st.subheader("Reporte de Asignacion")
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
        st.error(f"El Excel debe contener exactamente estas columnas: {columnas_requeridas}. Revisa que los nombres esten bien escritos.")
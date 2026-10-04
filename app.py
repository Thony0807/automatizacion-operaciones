import streamlit as st
import requests

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
    
    # Estructura exacta capturada desde la pestaña Network
    payload = {
        "data[Usuario][usuario]": usuario,
        "data[Usuario][pass]": clave
    }
    
    try:
        # Enviamos la petición POST para iniciar sesión
        respuesta = st.session_state.sesion_sigof.post(url_login, data=payload)
        
        # Validamos si el login fue exitoso comprobando si nos redirigió al dashboard
        # o si la respuesta contiene texto del panel principal
        if "dashboard/modulos" in respuesta.url or respuesta.status_code == 200:
            # Una validación extra por si el servidor devuelve 200 pero recarga el login por error
            if "Módulo de Lectura" in respuesta.text or "Módulo de Entrega" in respuesta.text:
                 return True
            # Si la URL cambió al dashboard, también es éxito seguro
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
    
    # Cajas de texto para ingresar credenciales
    usuario_input = st.text_input("Usuario")
    clave_input = st.text_input("Contraseña", type="password")
    
    # Botón de acción
    if st.button("Ingresar"):
        if usuario_input and clave_input:
            if iniciar_sesion(usuario_input, clave_input):
                st.session_state.autenticado = True
                st.rerun() # Recarga la interfaz para mostrar el panel de control
            else:
                st.error("Credenciales incorrectas o el sistema no respondió como se esperaba.")
        else:
            st.warning("Por favor ingresa tu usuario y contraseña.")

else:
    # Esta sección solo se muestra si el login fue exitoso
    st.success("Sesión iniciada correctamente en SIGOF. Conexión establecida.")
    st.write("El sistema está listo para ejecutar operaciones.")
    
    st.divider()
    
    st.subheader("Módulo: Asignación de Reparto (Pendiente)")
    st.info("Aquí integraremos el código para procesar el Excel una vez capturemos los datos de asignación.")
    
    st.divider()
    
    # Botón para salir y limpiar la sesión
    if st.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.sesion_sigof.cookies.clear()
        st.rerun()
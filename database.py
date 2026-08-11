import streamlit as st

# Usamos la conexión gestionada de Streamlit para PostgreSQL (Neon)
def get_connection():
    return st.connection("postgresql", type="sql")

# 1. Caché de la configuración inicial del negocio (Expira en 1 hora o al reiniciar)
@st.cache_data(ttl=3600)
def cargar_configuracion_db():
    conn = get_connection()
    df = conn.query("SELECT nombre_negocio, tema_color, logo_bytes, fondo_bytes FROM configuracion WHERE id = 1;", ttl=0)
    if not df.empty:
        row = df.iloc[0]
        return {
            "nombre_negocio": row["nombre_negocio"],
            "tema_color": row["tema_color"],
            "logo_bytes": row["logo_bytes"],
            "fondo_bytes": row["fondo_bytes"]
        }
    return None

# 2. Función genérica para ejecutar escrituras (INSERT, UPDATE, DELETE)
def ejecutar_escritura(query, params=None):
    conn = get_connection()
    with conn.session as session:
        session.execute(query, params or {})
        session.commit()

# 3. Función genérica para consultas rápidas sin caché
def ejecutar_consulta(query, params=None):
    conn = get_connection()
    return conn.query(query, params=params)

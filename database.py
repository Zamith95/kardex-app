import streamlit as st
from sqlalchemy import text

# Conexión gestionada de Streamlit para PostgreSQL (Neon)
def get_connection():
    return st.connection("postgresql", type="sql")

# 1. Caché de la configuración inicial del negocio (Expira en 1 hora)
@st.cache_data(ttl=3600)
def cargar_configuracion_db():
    conn = get_connection()
    try:
        df = conn.query("SELECT nombre_negocio, tema_color, logo_bytes, fondo_bytes FROM configuracion WHERE id = 1;", ttl=0)
        if not df.empty:
            row = df.iloc[0]
            return {
                "nombre_negocio": row["nombre_negocio"],
                "tema_color": row["tema_color"],
                "logo_bytes": row["logo_bytes"],
                "fondo_bytes": row["fondo_bytes"]
            }
    except Exception as e:
        st.error(f"Error al cargar configuración: {e}")
    return None

# 2. Cargar productos (Retorna un DataFrame válido o None si está vacío/falla)
@st.cache_data(ttl=60)
def cargar_productos_db():
    conn = get_connection()
    try:
        df = conn.query("SELECT * FROM productos;", ttl=0)
        if df is None or df.empty:
            return None
        return df
    except Exception as e:
        st.error(f"Error al cargar productos: {e}")
        return None

# 3. Función genérica para escrituras (INSERT, UPDATE, DELETE)
def ejecutar_escritura(query, params=None):
    conn = get_connection()
    with conn.session as session:
        session.execute(text(query) if isinstance(query, str) else query, params or {})
        session.commit()

# 4. Función genérica para consultas rápidas sin caché
def ejecutar_consulta(query, params=None):
    conn = get_connection()
    return conn.query(query, params=params, ttl=0)

# 5. Función compatible para consultas genéricas (Maneja %s, tuplas y diccionarios)
def ejecutar_query(query, params=None, fetch=False):
    conn = get_connection()
    
    # Reemplaza los placeholders estilo %s por parámetros con nombre
    if isinstance(params, (tuple, list)):
        formatted_params = {}
        for idx, val in enumerate(params):
            param_key = f"param_{idx}"
            query = query.replace("%s", f":{param_key}", 1)
            formatted_params[param_key] = val
        params = formatted_params

    if fetch:
        try:
            df = conn.query(query, params=params, ttl=0)
            if df is None or df.empty:
                return []
            return [tuple(x) for x in df.to_numpy()]
        except Exception:
            # En caso de desconexión SSL/Neon, intenta resetear el caché y reconectar
            st.cache_data.clear()
            df = conn.query(query, params=params, ttl=0)
            if df is None or df.empty:
                return []
            return [tuple(x) for x in df.to_numpy()]
    else:
        with conn.session as session:
            session.execute(text(query), params or {})
            session.commit()

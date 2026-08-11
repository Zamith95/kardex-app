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
        if df is not None and not df.empty:
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

# 2. Cargar productos (Retorna una lista de diccionarios que app.py puede iterar por clave)
@st.cache_data(ttl=60)
def cargar_productos_db():
    conn = get_connection()
    try:
        df = conn.query("SELECT * FROM productos ORDER BY nombre ASC;", ttl=0)
        if df is None or df.empty:
            return None
        # Convertir el DataFrame a lista de diccionarios para compatibilidad total con app.py
        return df.to_dict(orient="records")
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

# 5. Función compatible para consultas genéricas (Maneja %s, tuplas, diccionarios y commit opcional)
def ejecutar_query(query, params=None, fetch=False, commit=False):
    conn = get_connection()
    
    # Reemplaza los placeholders estilo %s por parámetros con nombre para SQLAlchemy
    if isinstance(params, (tuple, list)):
        formatted_params = {}
        query_mod = query
        for idx, val in enumerate(params):
            param_key = f"param_{idx}"
            query_mod = query_mod.replace("%s", f":{param_key}", 1)
            formatted_params[param_key] = val
        params = formatted_params
        query = query_mod

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

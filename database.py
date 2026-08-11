import streamlit as st
import re
from sqlalchemy import text

# Conexión gestionada de Streamlit para PostgreSQL (Neon)
def get_connection():
    return st.connection("postgresql", type="sql")

# 1. Configuración inicial
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

# 2. Cargar productos (Devuelve lista de diccionarios)
@st.cache_data(ttl=60)
def cargar_productos_db():
    conn = get_connection()
    try:
        df = conn.query("SELECT * FROM productos ORDER BY nombre ASC;", ttl=0)
        if df is None or df.empty:
            return None
        return df.to_dict(orient="records")
    except Exception as e:
        st.error(f"Error al cargar productos: {e}")
        return None

# 3. Escrituras genéricas
def ejecutar_escritura(query, params=None):
    conn = get_connection()
    with conn.session as session:
        session.execute(text(query) if isinstance(query, str) else query, params or {})
        session.commit()

# 4. Consultas sin caché
def ejecutar_consulta(query, params=None):
    conn = get_connection()
    return conn.query(query, params=params, ttl=0)

# 5. Función de compatibilidad total con app.py (Soporta %s, tuples, fetch y commit)
def ejecutar_query(query, params=None, fetch=False, commit=False):
    conn = get_connection()
    
    # Formateo seguro de parámetros (%s -> :param_X) para SQLAlchemy
    formatted_params = {}
    query_mod = query

    if isinstance(params, (tuple, list)):
        for idx, val in enumerate(params):
            param_key = f"param_{idx}"
            # Reemplazar exactamente la primera ocurrencia de %s que no sea escapada
            query_mod = re.sub(r'%s', f":{param_key}", query_mod, count=1)
            formatted_params[param_key] = val
    elif isinstance(params, dict):
        formatted_params = params

    if fetch:
        try:
            df = conn.query(query_mod, params=formatted_params if formatted_params else None, ttl=0)
            if df is None or df.empty:
                return []
            return [tuple(x) for x in df.to_numpy()]
        except Exception as e:
            st.error(f"Error en consulta SQL: {e}")
            return []
    else:
        try:
            with conn.session as session:
                session.execute(text(query_mod), formatted_params if formatted_params else {})
                session.commit()
        except Exception as e:
            st.error(f"Error al ejecutar escritura SQL: {e}")

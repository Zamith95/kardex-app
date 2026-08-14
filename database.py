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
    ejecutar_query(query, params, commit=True)

# 4. Consultas sin caché
def ejecutar_consulta(query, params=None):
    df = ejecutar_query(query, params, fetch=True, format_as_df=True)
    return df

# 5. Función de compatibilidad total con app.py (Soporta %s, tuples, fetch y commit)
def ejecutar_query(query, params=None, fetch=False, commit=False, format_as_df=False):
    conn = get_connection()
    
    formatted_params = {}
    query_mod = query

    if isinstance(params, (tuple, list)):
        for idx, val in enumerate(params):
            param_key = f"param_{idx}"
            query_mod = query_mod.replace("%s", f":{param_key}", 1)
            formatted_params[param_key] = val
    elif isinstance(params, dict):
        formatted_params = params

    try:
        if commit or not fetch:
            with conn.session as session:
                session.execute(text(query_mod), formatted_params)
                session.commit()
            return True
        else:
            if format_as_df:
                return conn.query(query_mod, params=formatted_params, ttl=0)
            else:
                with conn.session as session:
                    res = session.execute(text(query_mod), formatted_params)
                    return res.fetchall()
    except Exception as e:
        st.error(f"Error en consulta BD: {e}")
        return None

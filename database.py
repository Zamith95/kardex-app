import streamlit as st
from sqlalchemy import create_engine, text
from sqlalchemy.exc import OperationalError, DBAPIError

# Obtener URL desde secrets de Streamlit
def get_db_url():
    if "connections" in st.secrets and "postgresql" in st.secrets["connections"]:
        return st.secrets["connections"]["postgresql"]["url"]
    elif "DATABASE_URL" in st.secrets:
        return st.secrets["DATABASE_URL"]
    else:
        raise ValueError("No se encontró la URL de la base de datos en .streamlit/secrets.toml")

# Crear el Engine con protección anti-desconexión
@st.cache_resource(ttl=300)
def get_engine():
    db_url = get_db_url()
    return create_engine(
        db_url,
        pool_pre_ping=True,      # Valida conexión antes de consultar
        pool_recycle=300,        # Recicla conexiones cada 5 min
        pool_size=10,            # Tamaño del pool de conexiones
        max_overflow=20,         # Capacidad adicional para picos de uso
        connect_args={"connect_timeout": 10}
    )

# 1. Configuración inicial
@st.cache_data(ttl=3600)
def cargar_configuracion_db():
    try:
        df = ejecutar_consulta("SELECT nombre_negocio, tema_color, logo_bytes, fondo_bytes FROM configuracion WHERE id = 1;")
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
    try:
        df = ejecutar_consulta("SELECT * FROM productos ORDER BY nombre ASC;")
        if df is None or df.empty:
            return None
        return df.to_dict(orient="records")
    except Exception as e:
        st.error(f"Error al cargar productos: {e}")
        return None

# 3. Escrituras genéricas
def ejecutar_escritura(query, params=None):
    return ejecutar_query(query, params, commit=True)

# 4. Consultas sin caché
def ejecutar_consulta(query, params=None):
    import pandas as pd
    return ejecutar_query(query, params, fetch=True, format_as_df=True)

# 5. Función principal con reintento automático anti-desconexión
def ejecutar_query(query, params=None, fetch=False, commit=False, format_as_df=False):
    import pandas as pd
    
    formatted_params = {}
    query_mod = query

    if isinstance(params, (tuple, list)):
        for idx, val in enumerate(params):
            param_key = f"param_{idx}"
            query_mod = query_mod.replace("%s", f":{param_key}", 1)
            formatted_params[param_key] = val
    elif isinstance(params, dict):
        formatted_params = params

    def _ejecutar():
        engine = get_engine()
        with engine.connect() as conn:
            if commit or not fetch:
                with conn.begin():
                    conn.execute(text(query_mod), formatted_params)
                return True
            else:
                if format_as_df:
                    return pd.read_sql_query(text(query_mod), conn, params=formatted_params)
                else:
                    res = conn.execute(text(query_mod), formatted_params)
                    return res.fetchall()

    try:
        return _ejecutar()
    except (OperationalError, DBAPIError):
        # Si falla por caída de conexión, recarga el motor e intenta una 2da vez automáticamente
        st.cache_resource.clear()
        try:
            return _ejecutar()
        except Exception as e:
            st.error(f"Error en consulta BD: {e}")
            return None
    except Exception as e:
        st.error(f"Error en consulta BD: {e}")
        return None

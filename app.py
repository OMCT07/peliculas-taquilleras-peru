from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Taquilla cinematográfica en Perú",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

ANIOS = list(range(2015, 2026))

COLUMNAS_REQUERIDAS = [
    "Título",
    "Recaudación",
    "Dirección",
    "Productora",
    "País",
]


# ============================================================
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>
        .block-container {
            max-width: 1450px;
            padding-top: 1.5rem;
            padding-bottom: 3rem;
        }

        [data-testid="stSidebar"] {
            background-color: #f5f6f8;
        }

        .titulo-principal {
            font-size: 2.5rem;
            font-weight: 800;
            margin-bottom: 0.15rem;
        }

        .subtitulo-principal {
            color: #5e6472;
            font-size: 1.05rem;
            margin-bottom: 1.3rem;
        }

        .tarjeta-informativa {
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            padding: 1rem 1.2rem;
            background-color: #ffffff;
            margin-bottom: 0.8rem;
        }

        .nota-datos {
            border-left: 5px solid #f0a500;
            border-radius: 5px;
            padding: 0.9rem 1rem;
            background-color: #fff8e6;
            color: #574300;
            margin-top: 1rem;
            margin-bottom: 1rem;
        }

        .pie-pagina {
            color: #6b7280;
            font-size: 0.85rem;
            text-align: center;
            margin-top: 3rem;
        }

        div[data-testid="stMetric"] {
            border: 1px solid #e5e7eb;
            border-radius: 12px;
            padding: 1rem;
            background-color: #ffffff;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid #e5e7eb;
            border-radius: 10px;
            overflow: hidden;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def quitar_tildes(texto: str) -> str:
    """Devuelve el texto sin tildes para facilitar comparaciones."""
    texto_normalizado = unicodedata.normalize("NFKD", str(texto))
    return "".join(
        caracter
        for caracter in texto_normalizado
        if not unicodedata.combining(caracter)
    )


def normalizar_nombre_columna(nombre: str) -> str:
    """Normaliza nombres de columnas procedentes de Excel."""
    nombre = str(nombre).strip()
    nombre = re.sub(r"\s+", " ", nombre)

    mapa_columnas = {
        "titulo": "Título",
        "recaudacion": "Recaudación",
        "direccion": "Dirección",
        "productora": "Productora",
        "pais": "País",
    }

    nombre_clave = quitar_tildes(nombre).lower()

    return mapa_columnas.get(nombre_clave, nombre)


def limpiar_recaudacion(valor) -> float | None:
    """
    Convierte valores como '$11,300,000', 'US$ 2.500.000'
    o números de Excel en valores numéricos.
    """
    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if not texto:
        return None

    texto = texto.replace("US$", "")
    texto = texto.replace("USD", "")
    texto = texto.replace("$", "")
    texto = texto.replace("S/", "")
    texto = texto.replace(" ", "")

    # Elimina cualquier carácter que no sea número, coma, punto o signo.
    texto = re.sub(r"[^0-9,.\-]", "", texto)

    if not texto:
        return None

    # Casos con coma y punto:
    # 1,500,000.50 -> elimina comas
    # 1.500.000,50 -> elimina puntos y cambia coma decimal por punto
    if "," in texto and "." in texto:
        ultima_coma = texto.rfind(",")
        ultimo_punto = texto.rfind(".")

        if ultima_coma > ultimo_punto:
            texto = texto.replace(".", "")
            texto = texto.replace(",", ".")
        else:
            texto = texto.replace(",", "")

    elif texto.count(",") > 1:
        texto = texto.replace(",", "")

    elif texto.count(".") > 1:
        texto = texto.replace(".", "")

    elif "," in texto:
        parte_decimal = texto.split(",")[-1]

        if len(parte_decimal) == 3:
            texto = texto.replace(",", "")
        else:
            texto = texto.replace(",", ".")

    elif "." in texto:
        parte_decimal = texto.split(".")[-1]

        if len(parte_decimal) == 3:
            texto = texto.replace(".", "")

    try:
        return float(texto)
    except ValueError:
        return None


def limpiar_texto(valor) -> str:
    """Limpia espacios innecesarios en valores textuales."""
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()
    texto = re.sub(r"\s+", " ", texto)

    return texto


def separar_valores_categoricos(serie: pd.Series) -> list[str]:
    """
    Obtiene países o productoras individuales aunque estén escritos
    como 'Estados Unidos / Reino Unido'.
    """
    resultados: set[str] = set()

    for valor in serie.dropna():
        texto = limpiar_texto(valor)

        if not texto:
            continue

        partes = re.split(r"\s*/\s*|\s*;\s*|\s*\|\s*", texto)

        for parte in partes:
            parte = parte.strip()

            if parte:
                resultados.add(parte)

    return sorted(resultados)


def contiene_categoria(valor: str, categorias: list[str]) -> bool:
    """Comprueba si un valor compuesto contiene alguna categoría."""
    if not categorias:
        return True

    texto = limpiar_texto(valor).lower()

    return any(categoria.lower() in texto for categoria in categorias)


def formato_moneda(valor) -> str:
    """Formatea un valor numérico como moneda estadounidense."""
    if pd.isna(valor):
        return "Sin dato"

    return f"US$ {valor:,.0f}"


def nombre_archivo_excel(anio: int) -> str:
    return f"Peliculas_mas_taquilleras_Peru_{anio}.xlsx"


def ruta_archivo_excel(anio: int) -> Path:
    return DATA_DIR / nombre_archivo_excel(anio)


# ============================================================
# CARGA DE DATOS
# ============================================================

@st.cache_data(show_spinner=False)
def cargar_archivo(anio: int) -> pd.DataFrame:
    """Carga y limpia el Excel de un año determinado."""
    ruta = ruta_archivo_excel(anio)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo: {nombre_archivo_excel(anio)}"
        )

    df = pd.read_excel(ruta, engine="openpyxl")

    # Elimina columnas completamente vacías.
    df = df.dropna(axis=1, how="all")

    # Normaliza los encabezados.
    df.columns = [
        normalizar_nombre_columna(columna)
        for columna in df.columns
    ]

    columnas_faltantes = [
        columna
        for columna in COLUMNAS_REQUERIDAS
        if columna not in df.columns
    ]

    if columnas_faltantes:
        raise ValueError(
            f"El archivo de {anio} no contiene estas columnas: "
            + ", ".join(columnas_faltantes)
        )

    df = df[COLUMNAS_REQUERIDAS].copy()

    for columna in ["Título", "Dirección", "Productora", "País"]:
        df[columna] = df[columna].apply(limpiar_texto)

    df["Recaudación"] = df["Recaudación"].apply(limpiar_recaudacion)

    # Elimina filas sin título.
    df = df[df["Título"] != ""].copy()

    # Elimina filas totalmente duplicadas.
    df = df.drop_duplicates().reset_index(drop=True)

    df.insert(0, "Puesto", range(1, len(df) + 1))
    df.insert(1, "Año", anio)

    return df


@st.cache_data(show_spinner=False)
def cargar_todos_los_datos() -> tuple[pd.DataFrame, list[str]]:
    """Carga todos los Excel disponibles entre 2015 y 2025."""
    dataframes: list[pd.DataFrame] = []
    errores: list[str] = []

    for anio in ANIOS:
        try:
            dataframes.append(cargar_archivo(anio))
        except Exception as error:
            errores.append(f"{anio}: {error}")

    if not dataframes:
        return pd.DataFrame(), errores

    consolidado = pd.concat(dataframes, ignore_index=True)

    return consolidado, errores


# ============================================================
# EXPORTACIÓN
# ============================================================

def dataframe_a_excel(df: pd.DataFrame, nombre_hoja: str = "Películas") -> bytes:
    """Convierte un DataFrame en un archivo Excel descargable."""
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name=nombre_hoja[:31],
        )

        hoja = writer.sheets[nombre_hoja[:31]]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions

        anchos = {
            "A": 10,
            "B": 10,
            "C": 42,
            "D": 20,
            "E": 34,
            "F": 38,
            "G": 28,
        }

        for columna, ancho in anchos.items():
            hoja.column_dimensions[columna].width = ancho

    salida.seek(0)

    return salida.getvalue()


def dataframe_a_csv(df: pd.DataFrame) -> bytes:
    """Convierte un DataFrame a CSV compatible con Excel."""
    return df.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# FILTROS
# ============================================================

def aplicar_filtros(
    df: pd.DataFrame,
    busqueda: str,
    paises: list[str],
    productoras: list[str],
    directores: list[str],
    solo_con_recaudacion: bool,
) -> pd.DataFrame:
    resultado = df.copy()

    if busqueda:
        patron = re.escape(busqueda.strip())

        mascara = (
            resultado["Título"].str.contains(
                patron,
                case=False,
                na=False,
                regex=True,
            )
            | resultado["Dirección"].str.contains(
                patron,
                case=False,
                na=False,
                regex=True,
            )
            | resultado["Productora"].str.contains(
                patron,
                case=False,
                na=False,
                regex=True,
            )
            | resultado["País"].str.contains(
                patron,
                case=False,
                na=False,
                regex=True,
            )
        )

        resultado = resultado[mascara]

    if paises:
        resultado = resultado[
            resultado["País"].apply(
                lambda valor: contiene_categoria(valor, paises)
            )
        ]

    if productoras:
        resultado = resultado[
            resultado["Productora"].apply(
                lambda valor: contiene_categoria(valor, productoras)
            )
        ]

    if directores:
        resultado = resultado[
            resultado["Dirección"].isin(directores)
        ]

    if solo_con_recaudacion:
        resultado = resultado[
            resultado["Recaudación"].notna()
            & (resultado["Recaudación"] > 0)
        ]

    return resultado.reset_index(drop=True)


# ============================================================
# GRÁFICOS
# ============================================================

def grafico_ranking(df: pd.DataFrame, anio: int):
    datos = (
        df.dropna(subset=["Recaudación"])
        .query("Recaudación > 0")
        .sort_values("Recaudación", ascending=True)
    )

    if datos.empty:
        return None

    figura = px.bar(
        datos,
        x="Recaudación",
        y="Título",
        orientation="h",
        title=f"Recaudación por película en {anio}",
        labels={
            "Recaudación": "Recaudación en USD",
            "Título": "Película",
        },
        hover_data={
            "Dirección": True,
            "Productora": True,
            "País": True,
            "Recaudación": ":,.0f",
        },
        text_auto=".3s",
    )

    figura.update_layout(
        height=max(500, len(datos) * 34),
        margin=dict(l=20, r=20, t=70, b=20),
        yaxis_title=None,
        xaxis_tickprefix="US$ ",
        xaxis_tickformat=",",
    )

    figura.update_traces(
        textposition="outside",
        cliponaxis=False,
    )

    return figura


def grafico_por_pais(df: pd.DataFrame):
    registros: list[dict] = []

    for _, fila in df.iterrows():
        paises = re.split(
            r"\s*/\s*|\s*;\s*|\s*\|\s*",
            limpiar_texto(fila["País"]),
        )

        paises = [pais.strip() for pais in paises if pais.strip()]

        for pais in paises:
            registros.append(
                {
                    "País": pais,
                    "Películas": 1,
                    "Recaudación": fila["Recaudación"],
                }
            )

    if not registros:
        return None

    datos = pd.DataFrame(registros)

    resumen = (
        datos.groupby("País", as_index=False)
        .agg(
            Películas=("Películas", "sum"),
            Recaudación=("Recaudación", "sum"),
        )
        .sort_values("Películas", ascending=True)
    )

    figura = px.bar(
        resumen,
        x="Películas",
        y="País",
        orientation="h",
        title="Cantidad de películas según país de producción",
        hover_data={
            "Recaudación": ":,.0f",
        },
        text_auto=True,
    )

    figura.update_layout(
        height=max(450, len(resumen) * 45),
        margin=dict(l=20, r=20, t=70, b=20),
        yaxis_title=None,
    )

    return figura


def grafico_evolucion_anual(df: pd.DataFrame):
    datos_validos = df[
        df["Recaudación"].notna()
        & (df["Recaudación"] > 0)
    ].copy()

    if datos_validos.empty:
        return None

    resumen = (
        datos_validos.groupby("Año", as_index=False)
        .agg(
            Recaudación_total=("Recaudación", "sum"),
            Recaudación_promedio=("Recaudación", "mean"),
            Películas_con_dato=("Recaudación", "count"),
        )
    )

    figura = px.line(
        resumen,
        x="Año",
        y="Recaudación_total",
        markers=True,
        title="Evolución de la recaudación registrada por año",
        labels={
            "Recaudación_total": "Recaudación total en USD",
            "Año": "Año",
        },
        hover_data={
            "Recaudación_promedio": ":,.0f",
            "Películas_con_dato": True,
            "Recaudación_total": ":,.0f",
        },
    )

    figura.update_layout(
        height=500,
        xaxis=dict(dtick=1),
        yaxis_tickprefix="US$ ",
        yaxis_tickformat=",",
        margin=dict(l=20, r=20, t=70, b=20),
    )

    return figura


def grafico_pelicula_lider(df: pd.DataFrame):
    datos = df[
        df["Recaudación"].notna()
        & (df["Recaudación"] > 0)
    ].copy()

    if datos.empty:
        return None

    indices = datos.groupby("Año")["Recaudación"].idxmax()
    lideres = datos.loc[indices].sort_values("Año")

    figura = px.bar(
        lideres,
        x="Año",
        y="Recaudación",
        text="Título",
        title="Película con mayor recaudación registrada por año",
        labels={
            "Recaudación": "Recaudación en USD",
            "Año": "Año",
        },
        hover_data={
            "Título": True,
            "Dirección": True,
            "Productora": True,
            "País": True,
            "Recaudación": ":,.0f",
        },
    )

    figura.update_traces(
        textposition="outside",
        textangle=-35,
        cliponaxis=False,
    )

    figura.update_layout(
        height=620,
        xaxis=dict(dtick=1),
        yaxis_tickprefix="US$ ",
        yaxis_tickformat=",",
        margin=dict(l=20, r=20, t=70, b=150),
    )

    return figura


# ============================================================
# ENCABEZADO
# ============================================================

st.markdown(
    '<div class="titulo-principal">🎬 Taquilla cinematográfica en Perú</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="subtitulo-principal">
        Las 20 películas más taquilleras registradas en Perú
        entre 2015 y 2025.
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CARGA PRINCIPAL
# ============================================================

with st.spinner("Cargando los archivos Excel..."):
    df_total, errores_carga = cargar_todos_los_datos()

if df_total.empty:
    st.error(
        "No se pudo cargar ningún archivo. Comprueba que la carpeta "
        "`data` exista y contenga los 11 archivos Excel."
    )

    if errores_carga:
        with st.expander("Ver detalles de los errores"):
            for error in errores_carga:
                st.write(f"- {error}")

    st.stop()


if errores_carga:
    st.warning(
        "Algunos archivos no pudieron cargarse. La aplicación mostrará "
        "únicamente los años disponibles."
    )

    with st.expander("Ver archivos con problemas"):
        for error in errores_carga:
            st.write(f"- {error}")


anios_disponibles = sorted(df_total["Año"].unique().tolist())


# ============================================================
# BARRA LATERAL
# ============================================================

st.sidebar.title("🎛️ Panel de control")

st.sidebar.caption(
    "Utiliza los filtros para explorar las películas y sus datos."
)

pagina = st.sidebar.radio(
    "Sección",
    options=[
        "Ranking anual",
        "Análisis histórico",
        "Base completa",
        "Metodología",
    ],
)

st.sidebar.divider()


# ============================================================
# PÁGINA: RANKING ANUAL
# ============================================================

if pagina == "Ranking anual":
    anio_seleccionado = st.sidebar.selectbox(
        "Selecciona un año",
        options=anios_disponibles,
        index=len(anios_disponibles) - 1,
    )

    df_anio = df_total[
        df_total["Año"] == anio_seleccionado
    ].copy()

    busqueda = st.sidebar.text_input(
        "Buscar",
        placeholder="Título, dirección, país...",
    )

    opciones_pais = separar_valores_categoricos(df_anio["País"])

    paises_seleccionados = st.sidebar.multiselect(
        "País",
        options=opciones_pais,
    )

    opciones_productora = separar_valores_categoricos(
        df_anio["Productora"]
    )

    productoras_seleccionadas = st.sidebar.multiselect(
        "Productora",
        options=opciones_productora,
    )

    opciones_director = sorted(
        [
            valor
            for valor in df_anio["Dirección"].dropna().unique()
            if limpiar_texto(valor)
        ]
    )

    directores_seleccionados = st.sidebar.multiselect(
        "Dirección",
        options=opciones_director,
    )

    solo_con_recaudacion = st.sidebar.checkbox(
        "Mostrar solo películas con recaudación disponible",
        value=False,
    )

    orden = st.sidebar.selectbox(
        "Ordenar resultados",
        options=[
            "Puesto original",
            "Mayor recaudación",
            "Menor recaudación",
            "Título A-Z",
            "Título Z-A",
        ],
    )

    df_filtrado = aplicar_filtros(
        df=df_anio,
        busqueda=busqueda,
        paises=paises_seleccionados,
        productoras=productoras_seleccionadas,
        directores=directores_seleccionados,
        solo_con_recaudacion=solo_con_recaudacion,
    )

    if orden == "Mayor recaudación":
        df_filtrado = df_filtrado.sort_values(
            "Recaudación",
            ascending=False,
            na_position="last",
        )

    elif orden == "Menor recaudación":
        df_filtrado = df_filtrado.sort_values(
            "Recaudación",
            ascending=True,
            na_position="last",
        )

    elif orden == "Título A-Z":
        df_filtrado = df_filtrado.sort_values("Título")

    elif orden == "Título Z-A":
        df_filtrado = df_filtrado.sort_values(
            "Título",
            ascending=False,
        )

    else:
        df_filtrado = df_filtrado.sort_values("Puesto")

    df_filtrado = df_filtrado.reset_index(drop=True)

    st.subheader(f"Ranking de películas — {anio_seleccionado}")

    recaudacion_valida = df_filtrado[
        df_filtrado["Recaudación"].notna()
        & (df_filtrado["Recaudación"] > 0)
    ]

    total_recaudacion = recaudacion_valida["Recaudación"].sum()

    promedio_recaudacion = (
        recaudacion_valida["Recaudación"].mean()
        if not recaudacion_valida.empty
        else 0
    )

    if not recaudacion_valida.empty:
        fila_lider = recaudacion_valida.loc[
            recaudacion_valida["Recaudación"].idxmax()
        ]

        pelicula_lider = fila_lider["Título"]
        recaudacion_lider = fila_lider["Recaudación"]
    else:
        pelicula_lider = "Sin información"
        recaudacion_lider = None

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Películas mostradas",
        f"{len(df_filtrado)}",
    )

    col2.metric(
        "Recaudación total",
        formato_moneda(total_recaudacion),
    )

    col3.metric(
        "Recaudación promedio",
        formato_moneda(promedio_recaudacion),
    )

    col4.metric(
        "Película líder",
        pelicula_lider,
        delta=(
            formato_moneda(recaudacion_lider)
            if recaudacion_lider is not None
            else None
        ),
        delta_color="off",
    )

    if df_filtrado["Recaudación"].isna().any():
        st.markdown(
            """
            <div class="nota-datos">
                Algunas películas no tienen una cifra de recaudación
                registrada. Esos valores se muestran como “Sin dato” y no
                se incluyen en los cálculos monetarios.
            </div>
            """,
            unsafe_allow_html=True,
        )

    pestana_tabla, pestana_graficos, pestana_ficha = st.tabs(
        [
            "📋 Tabla",
            "📊 Gráficos",
            "🎞️ Ficha de película",
        ]
    )

    with pestana_tabla:
        columnas_mostradas = [
            "Puesto",
            "Título",
            "Recaudación",
            "Dirección",
            "Productora",
            "País",
        ]

        st.dataframe(
            df_filtrado[columnas_mostradas],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Puesto": st.column_config.NumberColumn(
                    "Puesto",
                    format="%d",
                    width="small",
                ),
                "Título": st.column_config.TextColumn(
                    "Título",
                    width="large",
                ),
                "Recaudación": st.column_config.NumberColumn(
                    "Recaudación",
                    format="US$ %.0f",
                    width="medium",
                ),
                "Dirección": st.column_config.TextColumn(
                    "Dirección",
                    width="medium",
                ),
                "Productora": st.column_config.TextColumn(
                    "Productora",
                    width="large",
                ),
                "País": st.column_config.TextColumn(
                    "País",
                    width="medium",
                ),
            },
        )

        col_descarga1, col_descarga2, col_descarga3 = st.columns(3)

        columnas_exportacion = [
            "Puesto",
            "Año",
            "Título",
            "Recaudación",
            "Dirección",
            "Productora",
            "País",
        ]

        datos_exportacion = df_filtrado[columnas_exportacion]

        col_descarga1.download_button(
            label="⬇️ Descargar resultados en Excel",
            data=dataframe_a_excel(
                datos_exportacion,
                nombre_hoja=str(anio_seleccionado),
            ),
            file_name=(
                f"ranking_peliculas_peru_{anio_seleccionado}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

        col_descarga2.download_button(
            label="⬇️ Descargar resultados en CSV",
            data=dataframe_a_csv(datos_exportacion),
            file_name=(
                f"ranking_peliculas_peru_{anio_seleccionado}.csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )

        archivo_original = ruta_archivo_excel(anio_seleccionado)

        if archivo_original.exists():
            with archivo_original.open("rb") as archivo:
                contenido_original = archivo.read()

            col_descarga3.download_button(
                label="📁 Descargar Excel original",
                data=contenido_original,
                file_name=archivo_original.name,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    with pestana_graficos:
        figura_ranking = grafico_ranking(
            df_filtrado,
            anio_seleccionado,
        )

        if figura_ranking is not None:
            st.plotly_chart(
                figura_ranking,
                use_container_width=True,
                config={
                    "displaylogo": False,
                    "responsive": True,
                },
            )
        else:
            st.info(
                "No hay datos de recaudación disponibles para crear "
                "el gráfico."
            )

        figura_pais = grafico_por_pais(df_filtrado)

        if figura_pais is not None:
            st.plotly_chart(
                figura_pais,
                use_container_width=True,
                config={
                    "displaylogo": False,
                    "responsive": True,
                },
            )

    with pestana_ficha:
        if df_filtrado.empty:
            st.info(
                "No hay películas disponibles con los filtros actuales."
            )
        else:
            pelicula_seleccionada = st.selectbox(
                "Selecciona una película",
                options=df_filtrado["Título"].tolist(),
            )

            ficha = df_filtrado[
                df_filtrado["Título"] == pelicula_seleccionada
            ].iloc[0]

            st.markdown(
                f"""
                <div class="tarjeta-informativa">
                    <h2>{ficha["Título"]}</h2>
                    <p><strong>Puesto:</strong> {ficha["Puesto"]}</p>
                    <p><strong>Año:</strong> {ficha["Año"]}</p>
                    <p><strong>Recaudación:</strong>
                    {formato_moneda(ficha["Recaudación"])}</p>
                    <p><strong>Dirección:</strong>
                    {ficha["Dirección"] or "Sin información"}</p>
                    <p><strong>Productora:</strong>
                    {ficha["Productora"] or "Sin información"}</p>
                    <p><strong>País:</strong>
                    {ficha["País"] or "Sin información"}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )


# ============================================================
# PÁGINA: ANÁLISIS HISTÓRICO
# ============================================================

elif pagina == "Análisis histórico":
    st.subheader("Análisis histórico 2015–2025")

    anios_analisis = st.sidebar.multiselect(
        "Años incluidos",
        options=anios_disponibles,
        default=anios_disponibles,
    )

    solo_datos_validos = st.sidebar.checkbox(
        "Excluir recaudaciones vacías o iguales a cero",
        value=True,
    )

    if not anios_analisis:
        st.warning("Selecciona por lo menos un año.")
        st.stop()

    df_historico = df_total[
        df_total["Año"].isin(anios_analisis)
    ].copy()

    if solo_datos_validos:
        df_historico = df_historico[
            df_historico["Recaudación"].notna()
            & (df_historico["Recaudación"] > 0)
        ]

    total_peliculas = len(df_historico)
    total_recaudado = df_historico["Recaudación"].sum()
    promedio = df_historico["Recaudación"].mean()

    peliculas_unicas = df_historico["Título"].nunique()

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Registros analizados",
        f"{total_peliculas}",
    )

    col2.metric(
        "Películas diferentes",
        f"{peliculas_unicas}",
    )

    col3.metric(
        "Recaudación acumulada",
        formato_moneda(total_recaudado),
    )

    col4.metric(
        "Promedio por película",
        formato_moneda(promedio),
    )

    figura_evolucion = grafico_evolucion_anual(df_historico)

    if figura_evolucion is not None:
        st.plotly_chart(
            figura_evolucion,
            use_container_width=True,
            config={
                "displaylogo": False,
                "responsive": True,
            },
        )

    figura_lideres = grafico_pelicula_lider(df_historico)

    if figura_lideres is not None:
        st.plotly_chart(
            figura_lideres,
            use_container_width=True,
            config={
                "displaylogo": False,
                "responsive": True,
            },
        )

    st.subheader("Películas con mayor recaudación registrada")

    top_n = st.slider(
        "Cantidad de películas que deseas comparar",
        min_value=5,
        max_value=min(50, max(5, len(df_historico))),
        value=min(20, max(5, len(df_historico))),
    )

    top_historico = (
        df_historico.dropna(subset=["Recaudación"])
        .sort_values("Recaudación", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )

    if not top_historico.empty:
        top_historico.insert(
            0,
            "Posición histórica",
            range(1, len(top_historico) + 1),
        )

        st.dataframe(
            top_historico[
                [
                    "Posición histórica",
                    "Año",
                    "Título",
                    "Recaudación",
                    "Dirección",
                    "Productora",
                    "País",
                ]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Recaudación": st.column_config.NumberColumn(
                    "Recaudación",
                    format="US$ %.0f",
                )
            },
        )

        st.download_button(
            label="⬇️ Descargar análisis histórico en Excel",
            data=dataframe_a_excel(
                top_historico,
                nombre_hoja="Análisis histórico",
            ),
            file_name="analisis_historico_taquilla_peru.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )


# ============================================================
# PÁGINA: BASE COMPLETA
# ============================================================

elif pagina == "Base completa":
    st.subheader("Base consolidada de películas")

    busqueda_general = st.sidebar.text_input(
        "Buscar en toda la base",
        placeholder="Título, dirección, productora...",
    )

    anios_base = st.sidebar.multiselect(
        "Años",
        options=anios_disponibles,
        default=anios_disponibles,
    )

    df_base = df_total[
        df_total["Año"].isin(anios_base)
    ].copy()

    if busqueda_general:
        patron = re.escape(busqueda_general.strip())

        mascara = (
            df_base["Título"].str.contains(
                patron,
                case=False,
                na=False,
            )
            | df_base["Dirección"].str.contains(
                patron,
                case=False,
                na=False,
            )
            | df_base["Productora"].str.contains(
                patron,
                case=False,
                na=False,
            )
            | df_base["País"].str.contains(
                patron,
                case=False,
                na=False,
            )
        )

        df_base = df_base[mascara]

    df_base = df_base.sort_values(
        ["Año", "Puesto"],
        ascending=[False, True],
    )

    st.write(
        f"Se encontraron **{len(df_base)} registros**."
    )

    st.dataframe(
        df_base[
            [
                "Año",
                "Puesto",
                "Título",
                "Recaudación",
                "Dirección",
                "Productora",
                "País",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Recaudación": st.column_config.NumberColumn(
                "Recaudación",
                format="US$ %.0f",
            )
        },
    )

    col1, col2 = st.columns(2)

    col1.download_button(
        label="⬇️ Descargar base consolidada en Excel",
        data=dataframe_a_excel(
            df_base,
            nombre_hoja="Base consolidada",
        ),
        file_name="peliculas_taquilleras_peru_2015_2025.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    col2.download_button(
        label="⬇️ Descargar base consolidada en CSV",
        data=dataframe_a_csv(df_base),
        file_name="peliculas_taquilleras_peru_2015_2025.csv",
        mime="text/csv",
        use_container_width=True,
    )


# ============================================================
# PÁGINA: METODOLOGÍA
# ============================================================

elif pagina == "Metodología":
    st.subheader("Metodología y consideraciones sobre los datos")

    st.markdown(
        """
        ### Alcance

        Esta aplicación presenta una base organizada de películas
        correspondientes al periodo 2015–2025.

        Cada archivo anual contiene las siguientes variables:

        - **Título**
        - **Recaudación**
        - **Dirección**
        - **Productora**
        - **País**

        ### Procesamiento realizado

        La aplicación:

        1. Lee automáticamente los archivos Excel ubicados en la carpeta
           `data`.
        2. Comprueba que cada archivo tenga las columnas requeridas.
        3. Limpia los valores monetarios.
        4. Normaliza los encabezados.
        5. Añade las variables `Año` y `Puesto`.
        6. Consolida todos los años en una sola base.
        7. Permite filtrar, ordenar, comparar y descargar los resultados.

        ### Interpretación de la recaudación

        Los valores se presentan en dólares estadounidenses cuando están
        disponibles.

        Las películas que no tienen una cifra registrada se muestran como
        **“Sin dato”** y no se incluyen en los cálculos de suma o promedio.

        ### Advertencia

        Antes de utilizar la aplicación como fuente académica, periodística
        o institucional, deben verificarse las cifras y rankings con las
        fuentes originales de taquilla.
        """
    )

    st.info(
        "Los archivos Excel constituyen la fuente de datos de la aplicación. "
        "Las modificaciones realizadas dentro de la web no alteran los "
        "archivos originales."
    )


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.markdown(
    """
    <div class="pie-pagina">
        Proyecto de visualización de la taquilla cinematográfica en Perú,
        2015–2025 · Desarrollado con Python y Streamlit.
    </div>
    """,
    unsafe_allow_html=True,
)

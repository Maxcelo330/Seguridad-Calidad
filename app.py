# -*- coding: utf-8 -*-
"""
DASHBOARD DE SEGURIDAD - ASEGURAMIENTO DE CALIDAD (EMBOL S.A.)
==============================================================
Lee la base generada por el CONSOLIDADOR (carpeta BD_CONSOLIDADA) y muestra
los desvios SSMA con foco en:
    * Reportes EMITIDOS por Calidad  (quien reporta pertenece a LPZ-Calidad)
    * Reportes RECIBIDOS por Calidad (el desvio ocurrio en la seccion LPZ-Calidad)

Interaccion tipo Power BI: al hacer clic en una barra (o celda del mapa de calor)
se filtra todo el dashboard por ese valor. Doble clic en el grafico, la X de
cada filtro o el boton "Limpiar" quitan la seleccion. El boton "Ver detalle"
abre el resumen y los registros de la seleccion.

Logos: coloque los archivos oficiales en la carpeta LOGOS con estos nombres
    logo_embol.png      y      logo_cocacola.png     (tambien .jpg / .jpeg / .svg)

Ejecutar con INICIAR_DASHBOARD.bat  (o: streamlit run app.py)
"""

import io
import json
import os
import sqlite3
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# CONFIGURACION
# ---------------------------------------------------------------------------
DIR_APP = os.path.dirname(os.path.abspath(__file__))
DIR_LOGOS = os.path.join(DIR_APP, "LOGOS")


def _buscar_base():
    """Busca la base en: ./data (version GitHub), ./ o ../BD_CONSOLIDADA (version local)."""
    for carpeta in (os.path.join(DIR_APP, "data"), DIR_APP,
                    os.path.join(os.path.dirname(DIR_APP), "BD_CONSOLIDADA")):
        db = os.path.join(carpeta, "BD_SEGURIDAD_CALIDAD.db")
        xlsx = os.path.join(carpeta, "BD_SEGURIDAD_CALIDAD.xlsx")
        if os.path.exists(db) or os.path.exists(xlsx):
            return db, xlsx
    carpeta = os.path.join(os.path.dirname(DIR_APP), "BD_CONSOLIDADA")
    return os.path.join(carpeta, "BD_SEGURIDAD_CALIDAD.db"), os.path.join(carpeta, "BD_SEGURIDAD_CALIDAD.xlsx")


RUTA_DB, RUTA_XLSX = _buscar_base()

CATEGORIAS = ["Acto inseguro", "Condición insegura", "Cuasi accidente", "Medio ambiente"]
MESES_CORTOS = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
                7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
RANGOS_ANTIG = ["0–30 días", "31–60 días", "61–90 días", "91–180 días", "> 180 días"]

# Nombre legible de cada dimension que se puede filtrar con clic
DIMENSIONES = {
    "ANIO_MES": "Mes", "CATEGORIA": "Categoría", "SECCION_DESVIO": "Sección del desvío",
    "SECCION_REPORTA": "Sección que reporta", "DESVIO_SECTOR": "Sector", "REPORTA_NOMBRE": "Reportante",
    "COMPORTAMIENTO_INSEGURO": "Comportamiento", "ORIGEN": "Origen", "DIA_SEMANA": "Día",
    "HORA_TXT": "Hora", "RANGO_ANTIG": "Antigüedad",
}


def buscar_logo(nombre):
    for ext in (".png", ".jpg", ".jpeg", ".svg", ".webp"):
        ruta = os.path.join(DIR_LOGOS, nombre + ext)
        if os.path.exists(ruta):
            return ruta
    return None


LOGO_EMBOL = buscar_logo("logo_embol")
LOGO_COCACOLA = buscar_logo("logo_cocacola")

st.set_page_config(page_title="Dashboard Seguridad · Calidad · Embol S.A.",
                   page_icon=LOGO_EMBOL or "🛡️", layout="wide")

# ---------------------------------------------------------------------------
# COLORES (dependen del tema claro / oscuro elegido en el menu ⋮)
# ---------------------------------------------------------------------------
try:
    MODO_OSCURO = st.context.theme.type == "dark"
except Exception:
    MODO_OSCURO = False

if MODO_OSCURO:
    ROJO, GRAFITO, GRIS = "#FF4B55", "#C9CCD1", "#6E737B"
    COLOR_CAT = {"Acto inseguro": "#FF4B55", "Condición insegura": "#4F8FD9",
                 "Cuasi accidente": "#E0A21A", "Medio ambiente": "#3FAE72"}
    ESCALA_ROJA = [[0, "#262730"], [0.25, "#5A2A30"], [0.5, "#9C2A34"], [0.75, "#E41E2B"], [1, "#FF7A82"]]
    COLORES_ANTIG = ["#6B3A40", "#9C2A34", "#E41E2B", "#FF4B55", "#FF8A91"]
    COLOR_EMITIDO, COLOR_EMITIDO_CLARO = "#FF4B55", "#FFA3A8"
    COLOR_RECIBIDO, COLOR_RECIBIDO_CLARO = "#5B9BD5", "#A9CBEF"
else:
    ROJO, GRAFITO, GRIS = "#E41E2B", "#3A3A3A", "#A6A6A6"
    COLOR_CAT = {"Acto inseguro": "#E41E2B", "Condición insegura": "#2A6BAF",
                 "Cuasi accidente": "#D18A00", "Medio ambiente": "#2E8B57"}
    ESCALA_ROJA = [[0, "#FFF5F5"], [0.25, "#FBC9CC"], [0.5, "#F2767F"], [0.75, "#E41E2B"], [1, "#9E0E18"]]
    COLORES_ANTIG = ["#F2A0A6", "#EC6B74", "#E41E2B", "#B5131E", "#7A0A12"]
    COLOR_EMITIDO, COLOR_EMITIDO_CLARO = "#E41E2B", "#F4A3A9"
    COLOR_RECIBIDO, COLOR_RECIBIDO_CLARO = "#1F5A96", "#9DBBDB"

OPACIDAD_ATENUADA = 0.28   # barras no seleccionadas (como Power BI)

st.markdown("""
<style>
  .block-container { padding-top: 2.2rem; }
  [data-testid="stMetricValue"] { font-weight: 600; }
  [class*="st-key-dbl_"] { display: none !important; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# ACCESO CON CONTRASEÑA
# ---------------------------------------------------------------------------
# La contraseña NO se guarda en el codigo: solo su huella (PBKDF2-SHA256 con sal) en
# .streamlit/secrets.toml (local) o en "Secrets" de Streamlit Community Cloud:
#   [auth]
#   password_hash = "..."     salt = "..."     cookie_key = "..."
#   iteraciones = 200000      horas_sesion = 8
# Se genera con CONFIGURAR_CONTRASENA.bat. Al ingresar, el navegador guarda una
# cookie firmada (HMAC) valida por "horas_sesion": al refrescar no pide la clave otra vez.
import hashlib
import hmac
import time

COOKIE_SESION = "dsc_sesion"
MAX_INTENTOS, BLOQUEO_SEG = 5, 60


def _conf_auth():
    try:
        conf = dict(st.secrets["auth"])
        return conf if all(k in conf for k in ("password_hash", "salt", "cookie_key")) else None
    except Exception:
        return None


AUTH = _conf_auth()


def _firma(exp):
    return hmac.new(AUTH["cookie_key"].encode(), f"v1|{exp}|{AUTH['password_hash']}".encode(),
                    hashlib.sha256).hexdigest()


def _token_nuevo():
    exp = int(time.time() + float(AUTH.get("horas_sesion", 8)) * 3600)
    return f"{exp}.{_firma(exp)}", exp - int(time.time())


def _token_valido(token):
    try:
        exp_txt, firma = str(token).split(".", 1)
        exp = int(exp_txt)
    except (ValueError, AttributeError):
        return False
    return exp > time.time() and hmac.compare_digest(firma, _firma(exp))


def _clave_correcta(clave):
    huella = hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(AUTH["salt"]),
                                 int(AUTH.get("iteraciones", 200_000))).hex()
    return hmac.compare_digest(huella, AUTH["password_hash"])


def _script_cookie(valor, max_age):
    js = (f"document.cookie='{COOKIE_SESION}={valor}; Max-Age={max_age}; Path=/; SameSite=Strict'"
          "+(location.protocol==='https:'?'; Secure':'');")
    try:
        st.html(f"<script>{js}</script>", unsafe_allow_javascript=True)
    except TypeError:   # versiones antiguas de Streamlit
        import streamlit.components.v1 as components
        components.html(f"<script>{js.replace('document.', 'window.parent.document.').replace('location.', 'window.parent.location.')}</script>",
                        height=0)


def cerrar_sesion():
    st.session_state.auth_ok = False
    st.session_state.borrar_cookie = True
    st.session_state.sesion_cerrada = True   # no reusar la cookie leida al abrir la pagina


def control_de_acceso():
    if AUTH is None:
        st.error("El acceso con contraseña no está configurado. Ejecute **CONFIGURAR_CONTRASENA.bat** "
                 "(local) o cargue el bloque **[auth]** en *Settings → Secrets* de Streamlit Cloud.")
        st.stop()
    if st.session_state.get("auth_ok"):
        if st.session_state.pop("guardar_cookie", False):
            _script_cookie(*_token_nuevo())
        return
    if not st.session_state.get("sesion_cerrada"):
        try:
            token = st.context.cookies.get(COOKIE_SESION)
        except Exception:
            token = None
        if token and _token_valido(token):
            st.session_state.auth_ok = True
            return
    if st.session_state.pop("borrar_cookie", False):
        _script_cookie("", 0)

    # Pantalla de ingreso
    _, centro, _ = st.columns([1, 1.2, 1])
    with centro:
        st.write("")
        if LOGO_EMBOL:
            _, l, _ = st.columns([1, 1, 1])
            l.image(LOGO_EMBOL, width="stretch")
        st.markdown("<h3 style='text-align:center;margin-bottom:0'>Seguridad · Aseguramiento de Calidad</h3>",
                    unsafe_allow_html=True)
        st.markdown("<p style='text-align:center;opacity:.7'>Embol S.A. · Acceso restringido</p>",
                    unsafe_allow_html=True)
        bloqueado_hasta = st.session_state.get("bloqueo_hasta", 0)
        with st.form("login", border=True):
            clave = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Ingresar", type="primary", width="stretch",
                                           icon=":material/lock_open:")
        if entrar:
            if time.time() < bloqueado_hasta:
                st.error(f"Demasiados intentos. Espere {int(bloqueado_hasta - time.time())} s.")
            elif clave and _clave_correcta(clave):
                st.session_state.auth_ok = True
                st.session_state.guardar_cookie = True
                st.session_state.sesion_cerrada = False
                st.session_state.intentos = 0
                st.rerun()
            else:
                st.session_state.intentos = st.session_state.get("intentos", 0) + 1
                if st.session_state.intentos >= MAX_INTENTOS:
                    st.session_state.bloqueo_hasta = time.time() + BLOQUEO_SEG
                    st.session_state.intentos = 0
                st.error("Contraseña incorrecta.")
        st.caption(f"La sesión queda guardada en este navegador por {AUTH.get('horas_sesion', 8)} horas.")
    st.stop()


control_de_acceso()


# ---------------------------------------------------------------------------
# DATOS
# ---------------------------------------------------------------------------
def _firma_archivos():
    return tuple(os.path.getmtime(p) if os.path.exists(p) else 0 for p in (RUTA_DB, RUTA_XLSX))


@st.cache_data(show_spinner="Cargando base consolidada ...")
def cargar_datos(_firma):
    if os.path.exists(RUTA_DB):
        con = sqlite3.connect(f"file:{RUTA_DB}?mode=ro", uri=True)
        try:
            df = pd.read_sql("SELECT * FROM BD_DESVIOS", con)
            log = pd.read_sql("SELECT * FROM LOG_CARGA", con)
        finally:
            con.close()
    elif os.path.exists(RUTA_XLSX):
        df = pd.read_excel(RUTA_XLSX, sheet_name="BD_DESVIOS")
        log = pd.read_excel(RUTA_XLSX, sheet_name="LOG_CARGA")
    else:
        return None, None

    for c in ["FECHA_REPORTE", "FECHA", "FECHA_CIERRE", "FECHA_PLANIFICADA", "FECHA_EXTRACCION"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    texto = df.select_dtypes(exclude=["number", "datetime"]).columns
    df[texto] = df[texto].fillna("").astype(str)
    df["SECCION_DESVIO"] = df["DESVIO_SECCION"].str.replace("LPZ-", "", regex=False)
    df["SECCION_REPORTA"] = df["REPORTA_SECCION"].str.replace("LPZ-", "", regex=False)
    df["MES"] = df["FECHA"].dt.to_period("M").dt.to_timestamp()
    df["ANIO_MES"] = df["FECHA"].dt.strftime("%Y-%m").fillna("")
    df["HORA_NUM"] = df["FECHA_REPORTE"].dt.hour
    df["HORA_TXT"] = df["HORA_NUM"].map(lambda h: f"{int(h):02d} h" if pd.notna(h) else "")
    dias = pd.to_numeric(df["DIAS_ABIERTO"], errors="coerce")
    df["RANGO_ANTIG"] = pd.cut(dias, bins=[-1, 30, 60, 90, 180, 100_000], labels=RANGOS_ANTIG) \
        .astype(str).replace({"nan": "", "<NA>": "", "None": ""})
    return df, log


df_total, log = cargar_datos(_firma_archivos())

if df_total is None:
    st.error("No se encontró la base consolidada en la carpeta **BD_CONSOLIDADA**. "
             "Ejecute primero **CONSOLIDADOR/EJECUTAR_CONSOLIDADOR.bat**.")
    st.stop()


# ---------------------------------------------------------------------------
# UTILIDADES
# ---------------------------------------------------------------------------
def fmt(n):
    return f"{int(round(n)):,}".replace(",", ".")


def pct(a, b):
    return f"{(a / b * 100):.1f}%".replace(".", ",") if b else "0%"


def etiqueta_mes(valor):
    """'2026-03' o Timestamp -> 'Mar 2026'."""
    ts = pd.Timestamp(valor + "-01") if isinstance(valor, str) else pd.Timestamp(valor)
    return f"{MESES_CORTOS[ts.month]} {ts.year}"


def texto_valor(dim, valor):
    return etiqueta_mes(valor) if dim == "ANIO_MES" else str(valor)


def recortar(t, n=50):
    t = str(t)
    return t if len(t) <= n else t[: n - 1] + "…"


def estilo(fig, alto=340, leyenda=True):
    fig.update_layout(height=alto, margin=dict(l=0, r=10, t=30 if leyenda else 10, b=0),
                      showlegend=leyenda, clickmode="event+select", dragmode=False,
                      legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title_text=""))
    return fig


# ---------------------------------------------------------------------------
# FILTRO CRUZADO (clic en graficos, estilo Power BI)
# ---------------------------------------------------------------------------
if "xf" not in st.session_state:
    st.session_state.xf = {}          # {dimension: [valores]}
if "xf_graficos" not in st.session_state:
    st.session_state.xf_graficos = {}  # {key_grafico: [dimensiones que fijo]}


def aplicar_xf(d, excluir=()):
    for dim, valores in st.session_state.xf.items():
        if dim in excluir or dim not in d.columns:
            continue
        d = d[d[dim].astype(str).isin([str(v) for v in valores])]
    return d


def seleccion(dim):
    return [str(v) for v in st.session_state.xf.get(dim, [])]


def _rgba(color, alfa):
    c = color.lstrip("#")
    r, g, b_ = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    return f"rgba({r},{g},{b_},{alfa})"


def mk(color, ops=None):
    """Marcador con las barras no seleccionadas atenuadas (la transparencia va en el color;
    usar marker.opacity con etiquetas de texto rompe los clics repetidos en Plotly)."""
    if not ops:
        return dict(color=color)
    colores = color if isinstance(color, (list, tuple)) else [color] * len(ops)
    return dict(color=[_rgba(c, o) if o < 1 else c for c, o in zip(colores, ops)])


def opacidades(valores, dim):
    sel = seleccion(dim)
    if not sel:
        return None
    return [1.0 if str(v) in sel else OPACIDAD_ATENUADA for v in valores]


def cd(**kw):
    """customdata de un punto: dict {dimension: valor} serializado."""
    return json.dumps({k: str(v) for k, v in kw.items()}, ensure_ascii=False)


def _nonce(key):
    return st.session_state.setdefault("xf_nonce", {}).get(key, 0)


def _reiniciar_grafico(key):
    """Crea el grafico con una clave nueva: asi un nuevo clic en la misma barra vuelve a
    enviar el evento (permite quitar el filtro con clic / doble clic sobre la misma barra)."""
    nonces = st.session_state.setdefault("xf_nonce", {})
    viejo = f"{key}__{nonces.get(key, 0)}"
    nonces[key] = nonces.get(key, 0) + 1
    st.session_state.pop(viejo, None)


def _al_seleccionar(key, widget_key):
    estado = st.session_state.get(widget_key)
    puntos = []
    if estado is not None:
        try:
            puntos = estado["selection"]["points"]
        except (KeyError, TypeError):
            puntos = getattr(getattr(estado, "selection", None), "points", []) or []
    previas = st.session_state.xf_graficos.get(key, [])
    if not puntos:                    # doble clic en la barra seleccionada: quitar su filtro
        for dim in previas:
            st.session_state.xf.pop(dim, None)
        st.session_state.xf_graficos.pop(key, None)
        return
    nuevos = {}
    for p in puntos:
        dato = p.get("customdata")
        if isinstance(dato, (list, tuple)):
            dato = dato[0] if dato else None
        if not dato:
            continue
        try:
            par = json.loads(dato)
        except (TypeError, ValueError):
            continue
        for dim, val in par.items():
            nuevos.setdefault(dim, [])
            if val not in nuevos[dim]:
                nuevos[dim].append(val)
    if not nuevos:
        return
    # Clic en lo mismo que ya estaba filtrado -> quitar (toggle)
    if all(sorted(st.session_state.xf.get(d, [])) == sorted(v) for d, v in nuevos.items()):
        for dim in nuevos:
            st.session_state.xf.pop(dim, None)
        st.session_state.xf_graficos.pop(key, None)
        return
    for dim in previas:
        if dim not in nuevos:
            st.session_state.xf.pop(dim, None)
    st.session_state.xf.update(nuevos)
    st.session_state.xf_graficos[key] = list(nuevos)


def limpiar_xf():
    st.session_state.xf = {}
    for key in list(st.session_state.xf_graficos):
        if key != "pills_tipo":
            _reiniciar_grafico(key)
    st.session_state.xf_graficos = {}


def quitar_dim(dim):
    st.session_state.xf.pop(dim, None)
    for key, dims in list(st.session_state.xf_graficos.items()):
        if dim in dims:
            st.session_state.xf_graficos.pop(key)
            if key != "pills_tipo":
                _reiniciar_grafico(key)


def _al_cambiar_tipo():
    valores = st.session_state.get("pills_tipo") or []
    if valores:
        st.session_state.xf["CATEGORIA"] = list(valores)
        st.session_state.xf_graficos["pills_tipo"] = ["CATEGORIA"]
    else:
        quitar_dim("CATEGORIA")


CONFIG_PLOTLY = {"displaylogo": False, "doubleClick": False,
                 "modeBarButtonsToRemove": ["zoom2d", "pan2d", "select2d", "lasso2d", "zoomIn2d", "zoomOut2d",
                                            "autoScale2d", "resetScale2d"]}


def quitar_de_grafico(key):
    """Doble clic en un grafico: quita el filtro que ese grafico puso."""
    for dim in list(st.session_state.xf_graficos.get(key, [])):
        quitar_dim(dim)


def mostrar(fig, key=None):
    if key is None:
        st.plotly_chart(fig, width="stretch", config=CONFIG_PLOTLY)
    else:
        widget_key = f"{key}__{_nonce(key)}"
        st.plotly_chart(fig, width="stretch", key=widget_key, on_select=lambda k=key, w=widget_key: _al_seleccionar(k, w),
                        selection_mode="points", config=CONFIG_PLOTLY)
        # Boton oculto que el doble clic del grafico "presiona" (ver JS_DOBLE_CLIC)
        st.button("quitar", key=f"dbl_{key}", on_click=quitar_de_grafico, args=(key,))


# Doble clic en un grafico -> presiona su boton oculto "dbl_<grafico>" (quita el filtro de ese grafico)
JS_DOBLE_CLIC = """<script>
(function(){
  if (window.__dscDoble) return; window.__dscDoble = true;
  function enlazar(){
    document.querySelectorAll('[class*="st-key-"]').forEach(function(c){
      var m = (c.className.match(/st-key-([a-z0-9_]+?)__\\d+/) || [])[1];
      if (!m) return;
      var gd = c.querySelector('.js-plotly-plot');
      if (!gd || gd.__dscDoble || !gd.on) return;
      gd.__dscDoble = true;
      gd.on('plotly_doubleclick', function(){
        var b = document.querySelector('.st-key-dbl_' + m + ' button');
        if (b) b.click();
        return false;
      });
    });
  }
  new MutationObserver(enlazar).observe(document.body, {childList: true, subtree: true});
  setInterval(enlazar, 1500);
  enlazar();
})();
</script>"""


def activar_doble_clic():
    try:
        st.html(JS_DOBLE_CLIC, unsafe_allow_javascript=True)
    except TypeError:
        import streamlit.components.v1 as components
        components.html(JS_DOBLE_CLIC.replace("document.", "window.parent.document.")
                        .replace("new MutationObserver", "new window.parent.MutationObserver"), height=0)


def barras_h(serie, color, dim, key, top=10, total=None, resaltar=None, color_resaltado=None,
             hover="Reportes", etiquetas_cortas=False):
    s = serie.sort_values(ascending=False).head(top)[::-1]
    if s.empty:
        st.caption("Sin datos para la selección actual.")
        return
    if resaltar is not None and not isinstance(resaltar, (set, list, tuple)):
        resaltar = {resaltar}
    colores = [(color_resaltado or ROJO) if k in resaltar else color for k in s.index] if resaltar else color
    texto = [fmt(v) + (f" ({pct(v, total)})" if total else "") for v in s.values]
    y = [recortar(i) if etiquetas_cortas else str(i) for i in s.index]
    fig = go.Figure(go.Bar(
        x=s.values, y=y, orientation="h", text=texto, textposition="outside", cliponaxis=False,
        marker=mk(colores, opacidades(s.index, dim)),
        customdata=[cd(**{dim: k}) for k in s.index],
        hovertemplate="%{y}<br>" + hover + ": %{x}<extra></extra>"))
    estilo(fig, max(240, 30 * len(s) + 40), leyenda=False)
    if len(s) < 5:
        fig.update_layout(bargap=0.75 - 0.1 * len(s))
    fig.update_xaxes(range=[0, s.max() * 1.22], showticklabels=False)
    mostrar(fig, key)


def evolucion_categoria(d, key, alto=380, titulo_y="N° de desvíos"):
    """Barras apiladas por mes y categoria (clic = filtra mes + categoria)."""
    piv = d.pivot_table(index="ANIO_MES", columns="CATEGORIA", values="ID_DESVIO", aggfunc="count", fill_value=0)
    if piv.empty:
        st.caption("Sin datos para la selección actual.")
        return
    etiquetas = [etiqueta_mes(m) for m in piv.index]
    sel_m, sel_c = seleccion("ANIO_MES"), seleccion("CATEGORIA")
    fig = go.Figure()
    for c in CATEGORIAS:
        if c in piv.columns:
            op = None
            if sel_m or sel_c:
                op = [1.0 if ((not sel_m or m in sel_m) and (not sel_c or c in sel_c)) else OPACIDAD_ATENUADA
                      for m in piv.index]
            fig.add_bar(x=etiquetas, y=piv[c], name=c, marker=mk(COLOR_CAT[c], op),
                        customdata=[cd(ANIO_MES=m, CATEGORIA=c) for m in piv.index])
    tot = piv.sum(axis=1)
    fig.add_scatter(x=etiquetas, y=tot.values, mode="text", text=[fmt(v) for v in tot.values],
                    textposition="top center", showlegend=False, hoverinfo="skip")
    estilo(fig, alto)
    fig.update_layout(barmode="stack", hovermode="x unified", legend=dict(traceorder="normal"))
    fig.update_yaxes(title_text=titulo_y, range=[0, tot.max() * 1.15])
    mostrar(fig, key)


def heatmap(d, key, etiqueta_hover="Desvíos"):
    if d.empty:
        st.caption("Sin datos para la selección actual.")
        return
    hm = d.pivot_table(index="SECCION_DESVIO", columns="ANIO_MES", values="ID_DESVIO", aggfunc="count", fill_value=0)
    hm = hm.loc[hm.sum(axis=1).sort_values(ascending=True).index]
    meses = list(hm.columns)
    custom = [[cd(SECCION_DESVIO=sec, ANIO_MES=m) for m in meses] for sec in hm.index]
    xs = [etiqueta_mes(m) for m in meses]
    fig = go.Figure(go.Heatmap(z=hm.values, x=xs, y=list(hm.index), colorscale=ESCALA_ROJA, text=hm.values,
                               texttemplate="%{text}", xgap=2, ygap=2, colorbar=dict(thickness=10),
                               hoverinfo="skip"))
    # Capa invisible de puntos (una por celda) para poder hacer clic en el mapa de calor
    sel_s, sel_m = seleccion("SECCION_DESVIO"), seleccion("ANIO_MES")
    px, py, pc, pt, lineas = [], [], [], [], []
    for i, sec in enumerate(hm.index):
        for j, m in enumerate(meses):
            px.append(xs[j]); py.append(sec); pc.append(custom[i][j]); pt.append(int(hm.values[i][j]))
            marcado = (sel_s or sel_m) and (not sel_s or sec in sel_s) and (not sel_m or m in sel_m)
            lineas.append(2.5 if marcado else 0)
    fig.add_scatter(x=px, y=py, mode="markers", customdata=pc, text=pt, showlegend=False,
                    marker=dict(symbol="square", size=30, color="rgba(0,0,0,0)",
                                line=dict(color=GRAFITO, width=lineas)),
                    selected=dict(marker=dict(opacity=1)), unselected=dict(marker=dict(opacity=1)),
                    hovertemplate="%{y} · %{x}<br>" + etiqueta_hover + ": %{text}<extra></extra>")
    estilo(fig, max(280, 34 * len(hm) + 60), leyenda=False)
    mostrar(fig, key)


# ---------------------------------------------------------------------------
# BARRA LATERAL
# ---------------------------------------------------------------------------
with st.sidebar:
    # Logo Coca-Cola centrado en la barra lateral (PNG con fondo transparente)
    if LOGO_COCACOLA:
        st.markdown('<div style="height:0.8rem"></div>', unsafe_allow_html=True)
        _, centro, _ = st.columns([1, 5, 1])
        centro.image(LOGO_COCACOLA, width="stretch")
        st.divider()
    st.header("Filtros")
    fmin, fmax = df_total["FECHA"].min().date(), df_total["FECHA"].max().date()
    rango = st.date_input("Periodo", value=(fmin, fmax), min_value=fmin, max_value=fmax, format="DD/MM/YYYY")
    cats = st.multiselect("Categoría de desvío", CATEGORIAS, default=CATEGORIAS)
    origenes = sorted(o for o in df_total["ORIGEN"].unique() if o)
    orig_sel = st.multiselect("Origen del reporte", origenes, placeholder="Todos")
    tipos = sorted(t for t in df_total["REPORTA_TIPO_PERSONAL"].unique() if t)
    tipo_sel = st.multiselect("Tipo de personal (quien reporta)", tipos, placeholder="Todos")

    st.divider()
    if st.button("Recargar datos", icon=":material/refresh:", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    st.button("Cerrar sesión", icon=":material/logout:", width="stretch", on_click=cerrar_sesion)
    extr = pd.to_datetime(log["FECHA_EXTRACCION"], errors="coerce").max() if log is not None else None
    st.caption(f"**Fuente:** Plataforma Web de Desvíos SSMA y Calidad  \n"
               f"**Última extracción:** {extr:%d/%m/%Y %H:%M}  \n"
               f"**Registros en base:** {fmt(len(df_total))}")
    with st.expander("Definiciones"):
        st.markdown("- **Emitido por Calidad:** quien reporta pertenece a la sección LPZ-Calidad.\n"
                    "- **Recibido por Calidad:** el desvío ocurrió en la sección LPZ-Calidad, "
                    "sin importar quién lo reportó.\n"
                    "- **Autorreporte:** emitido y recibido a la vez (Calidad reporta en su propia área).")
    with st.expander("¿Cómo filtrar con clic?"):
        st.markdown("- **Clic** en una barra o celda: filtra todo el dashboard.\n"
                    "- **Doble clic** en el gráfico (o clic otra vez en la misma barra): quita el filtro de ese gráfico.\n"
                    "- **Clic en otra barra** del mismo gráfico: cambia el filtro.\n"
                    "- **✕** en el filtro (arriba) o **Limpiar**: quita la selección.\n"
                    "- **Shift + clic**: selecciona varias barras.\n"
                    "- **Ver detalle**: abre el resumen y los registros de la selección.")

# Filtros de la barra lateral
if isinstance(rango, (list, tuple)) and len(rango) == 2:
    f_ini, f_fin = pd.Timestamp(rango[0]), pd.Timestamp(rango[1])
else:
    f_ini, f_fin = pd.Timestamp(fmin), pd.Timestamp(fmax)

df_base = df_total[(df_total["FECHA"] >= f_ini) & (df_total["FECHA"] <= f_fin)
                   & (df_total["CATEGORIA"].isin(cats))]
if orig_sel:
    df_base = df_base[df_base["ORIGEN"].isin(orig_sel)]
if tipo_sel:
    df_base = df_base[df_base["REPORTA_TIPO_PERSONAL"].isin(tipo_sel)]

# Filtro cruzado de los graficos
df = aplicar_xf(df_base)


def partes(d):
    """(emitidos, recibidos) por Calidad de un subconjunto."""
    return d[d["REPORTA_ES_CALIDAD"] == "Sí"], d[d["DESVIO_EN_CALIDAD"] == "Sí"]


def datos(*excluir):
    """Datos para un grafico: aplica todos los filtros de clic excepto los de su propia dimension."""
    return aplicar_xf(df_base, excluir=excluir)


# ---------------------------------------------------------------------------
# DETALLE DE LA SELECCION (drill-through)
# ---------------------------------------------------------------------------
COLS_DETALLE = ["ID_DESVIO", "CATEGORIA", "FECHA_REPORTE", "REPORTA_NOMBRE", "SECCION_REPORTA",
                "SECCION_DESVIO", "DESVIO_SECTOR", "DESCRIPCION", "COMPORTAMIENTO_INSEGURO", "ESTADO"]
CONFIG_DETALLE = {"ID_DESVIO": "N° desvío", "CATEGORIA": "Categoría",
                  "FECHA_REPORTE": st.column_config.DatetimeColumn("Fecha", format="DD/MM/YYYY HH:mm"),
                  "REPORTA_NOMBRE": "Reporta", "SECCION_REPORTA": "Sección reporta",
                  "SECCION_DESVIO": "Sección desvío", "DESVIO_SECTOR": "Sector",
                  "DESCRIPCION": st.column_config.TextColumn("Descripción", width="large"),
                  "COMPORTAMIENTO_INSEGURO": "Comportamiento", "ESTADO": "Estado"}


def a_excel(d, cols):
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as xw:
        d[cols].to_excel(xw, index=False, sheet_name="DATOS")
    return salida.getvalue()


@st.dialog("Detalle de la selección", width="large")
def ver_detalle():
    filtros = " · ".join(f"**{DIMENSIONES.get(k, k)}:** {', '.join(texto_valor(k, v) for v in vals)}"
                         for k, vals in st.session_state.xf.items())
    st.markdown(filtros or "Sin selección: se muestran todos los registros filtrados.")
    e, r = partes(df)
    alcance = st.segmented_control("Registros", ["Todos", "Emitidos por Calidad", "Recibidos por Calidad"],
                                   default="Todos", label_visibility="collapsed", key="alcance_detalle")
    d = {"Emitidos por Calidad": e, "Recibidos por Calidad": r}.get(alcance, df)
    c = st.columns(4)
    c[0].metric("Desvíos", fmt(len(df)), border=True)
    c[1].metric(":red[■] Emitidos por Calidad", fmt(len(e)), border=True)
    c[2].metric(":blue[■] Recibidos por Calidad", fmt(len(r)), border=True)
    c[3].metric("Abiertos", fmt((df["ESTADO_CIERRE"] == "Abierto").sum()), border=True)

    a, b = st.columns(2)
    res_cat = d.groupby("CATEGORIA").size().reindex(CATEGORIAS, fill_value=0)
    a.markdown(f"**Por tipo de desvío** · {alcance.lower()}")
    a.dataframe(pd.DataFrame({"Tipo de desvío": CATEGORIAS, "Desvíos": res_cat.values,
                              "%": (res_cat.values / max(len(d), 1) * 100).round(1)}),
                width="stretch", hide_index=True)
    res_sec = d.groupby("SECCION_DESVIO").size().sort_values(ascending=False).head(8)
    b.markdown(f"**Por sección del desvío** · {alcance.lower()}")
    b.dataframe(pd.DataFrame({"Sección": res_sec.index, "Desvíos": res_sec.values}), width="stretch", hide_index=True)

    st.markdown(f"**Registros ({fmt(len(d))})**")
    st.dataframe(d[COLS_DETALLE].sort_values("FECHA_REPORTE", ascending=False), width="stretch",
                 hide_index=True, height=320, column_config=CONFIG_DETALLE)
    st.download_button("Descargar Excel de la selección", a_excel(d, COLS_DETALLE), icon=":material/download:",
                       file_name=f"seleccion_desvios_{datetime.now():%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# ENCABEZADO
# ---------------------------------------------------------------------------
if LOGO_EMBOL:
    h1, h2 = st.columns([0.7, 9], vertical_alignment="center", gap="small")
    h1.image(LOGO_EMBOL, width=72)
    contenedor_titulo = h2
else:
    contenedor_titulo = st.container()

with contenedor_titulo:
    st.title("Seguridad · Aseguramiento de Calidad", anchor=False)
    st.caption(f"Embol S.A. · Planta La Paz · Desvíos SSMA  |  Periodo {f_ini:%d/%m/%Y} – {f_fin:%d/%m/%Y}"
               f"  |  {fmt(len(df))} desvíos")

# Barra de seleccion activa
with st.container(border=bool(st.session_state.xf)):
    if st.session_state.xf:
        c1, c2, c3 = st.columns([7, 1.3, 1.1], vertical_alignment="center")
        with c1:
            st.markdown("**Selección activa:**")
            chips = st.columns(max(len(st.session_state.xf), 1))
            for i, (dim, vals) in enumerate(list(st.session_state.xf.items())):
                etiqueta = f"{DIMENSIONES.get(dim, dim)}: {', '.join(recortar(texto_valor(dim, v), 28) for v in vals)}"
                chips[i].button(etiqueta, icon=":material/close:", key=f"chip_{dim}", on_click=quitar_dim,
                                args=(dim,), help="Quitar este filtro", width="stretch")
        if c2.button("Ver detalle", icon=":material/table_view:", type="primary", width="stretch"):
            ver_detalle()
        c3.button("Limpiar", icon=":material/filter_alt_off:", on_click=limpiar_xf, width="stretch")
    else:
        st.caption(":material/touch_app: Haz clic en cualquier barra o celda para filtrar todo el dashboard, "
                   "como en Power BI. Doble clic en el gráfico para quitar su filtro.")

if df_base.empty:
    st.warning("No hay registros con los filtros seleccionados.")
    st.stop()

emit, recib = partes(df)
recib_ext = recib[recib["REPORTA_ES_CALIDAD"] != "Sí"]
auto = recib[recib["REPORTA_ES_CALIDAD"] == "Sí"]
emit_ext = emit[emit["DESVIO_EN_CALIDAD"] != "Sí"]
meses_periodo = max(1, df["ANIO_MES"].nunique())

activar_doble_clic()

tab1, tab2, tab3, tab4 = st.tabs(["Calidad: emitidos vs recibidos", "Visión general planta",
                                  "Seguimiento y cierre", "Datos"])

# ===========================================================================
# TAB 1 - CALIDAD
# ===========================================================================
with tab1:
    # Clasificacion por los 4 tipos de desvio de seguridad (sincronizado con el filtro de clic)
    st.session_state["pills_tipo"] = [c for c in CATEGORIAS if c in seleccion("CATEGORIA")]
    st.pills("Tipo de desvío de seguridad", CATEGORIAS, selection_mode="multi", key="pills_tipo",
             on_change=_al_cambiar_tipo, help="Sin selección = los 4 tipos. Puede elegir uno o varios.")

    ratio = (len(emit) / len(recib)) if len(recib) else 0
    rank_sec = df.groupby("SECCION_REPORTA").size().sort_values(ascending=False)
    pos = list(rank_sec.index).index("Calidad") + 1 if "Calidad" in rank_sec.index else "-"

    cols = st.columns(6)
    cols[0].metric(":red[■] Emitidos", fmt(len(emit)), border=True,
                   help=f"Desvíos reportados por personal de LPZ-Calidad · {fmt(len(emit) / meses_periodo)} por mes · "
                        f"{pct(len(emit), len(df))} de la planta · Calidad es la sección N° {pos} de {len(rank_sec)} que más reporta.")
    cols[1].metric(":blue[■] Recibidos", fmt(len(recib)), border=True,
                   help=f"Desvíos ocurridos en la sección LPZ-Calidad · {fmt(len(recib) / meses_periodo)} por mes · "
                        f"{pct(len(recib), len(df))} de la planta.")
    cols[2].metric(":red[■] A otras áreas", fmt(len(emit_ext)), border=True,
                   help=f"{pct(len(emit_ext), len(emit))} de lo emitido por Calidad.")
    cols[3].metric(":blue[■] De otras áreas", fmt(len(recib_ext)), border=True,
                   help=f"{pct(len(recib_ext), len(recib))} de lo recibido por Calidad.")
    cols[4].metric("Autorreportes", fmt(len(auto)), border=True,
                   help="Desvíos que Calidad reportó en su propia área (emitidos y recibidos a la vez).")
    cols[5].metric("Reportantes", fmt(emit["REPORTA_ID"].nunique()), border=True,
                   help="Personas de Calidad que emitieron al menos un reporte · índice emitidos/recibidos: "
                        + f"{ratio:.1f}".replace(".", ","))

    # Resumen por tipo de desvio: emitidos y recibidos
    e_t, r_t = partes(datos("CATEGORIA"))
    cols = st.columns(4)
    for i, c in enumerate(CATEGORIAS):
        with cols[i].container(border=True):
            atenuado = bool(seleccion("CATEGORIA")) and c not in seleccion("CATEGORIA")
            st.markdown(f'<span style="display:inline-block;width:10px;height:10px;border-radius:2px;'
                        f'background:{COLOR_CAT[c]};opacity:{0.35 if atenuado else 1};margin-right:6px"></span>'
                        f'<b>{c}</b>', unsafe_allow_html=True)
            x1, x2 = st.columns(2)
            x1.metric(":red[■] Emitidos", fmt((e_t["CATEGORIA"] == c).sum()))
            x2.metric(":blue[■] Recibidos", fmt((r_t["CATEGORIA"] == c).sum()))

    st.subheader("Evolución mensual: emitidos vs recibidos", anchor=False)
    with st.container(border=True):
        e, r = partes(datos("ANIO_MES"))
        m_e, m_r = e.groupby("ANIO_MES").size(), r.groupby("ANIO_MES").size()
        idx = sorted(set(m_e.index) | set(m_r.index))
        if not idx:
            st.caption("Sin datos para la selección actual.")
        else:
            m_e, m_r = m_e.reindex(idx, fill_value=0), m_r.reindex(idx, fill_value=0)
            etiquetas = [etiqueta_mes(m) for m in idx]
            op = opacidades(idx, "ANIO_MES")
            fig = go.Figure()
            fig.add_bar(x=etiquetas, y=m_e.values, name="Emitidos por Calidad", text=m_e.values,
                        marker=mk(COLOR_EMITIDO, op), textposition="outside", cliponaxis=False,
                        customdata=[cd(ANIO_MES=m) for m in idx])
            fig.add_bar(x=etiquetas, y=m_r.values, name="Recibidos por Calidad", text=m_r.values,
                        marker=mk(COLOR_RECIBIDO, op), textposition="outside", cliponaxis=False,
                        customdata=[cd(ANIO_MES=m) for m in idx])
            fig.add_hline(y=m_e.mean(), line_dash="dot", line_color=ROJO, line_width=1,
                          annotation_text=f"Promedio emitidos: {m_e.mean():.0f}",
                          annotation_position="top right", annotation_font_color=ROJO)
            estilo(fig, 380)
            fig.update_layout(barmode="group", hovermode="x unified")
            fig.update_yaxes(title_text="N° de reportes", range=[0, max(m_e.max(), m_r.max(), 1) * 1.2])
            mostrar(fig, "c_mes")

    st.subheader("Evolución mensual por categoría", anchor=False)
    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Emitidos por Calidad, por tipo de desvío** &nbsp; :red-badge[Emitidos]")
        e, _ = partes(datos("ANIO_MES", "CATEGORIA"))
        evolucion_categoria(e, "c_mes_cat_e", alto=340, titulo_y="N° de reportes")
    with b.container(border=True):
        st.markdown("**Recibidos por Calidad, por tipo de desvío** &nbsp; :blue-badge[Recibidos]")
        _, r = partes(datos("ANIO_MES", "CATEGORIA"))
        evolucion_categoria(r, "c_mes_cat_r", alto=340, titulo_y="N° de reportes")

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**¿A qué áreas reportamos?** &nbsp; :red-badge[Emitidos]")
        st.caption("Sección donde ocurrió el desvío emitido por Calidad · tono claro: nuestra propia área")
        e, _ = partes(datos("SECCION_DESVIO"))
        barras_h(e.groupby("SECCION_DESVIO").size(), COLOR_EMITIDO, "SECCION_DESVIO", "c_areas", top=12,
                 total=len(e), resaltar="Calidad", color_resaltado=COLOR_EMITIDO_CLARO, hover="Emitidos")
    with b.container(border=True):
        st.markdown("**¿Quién nos reporta?** &nbsp; :blue-badge[Recibidos]")
        st.caption("Sección de quien reporta desvíos en el área de Calidad · tono claro: autorreportes")
        _, r = partes(datos("SECCION_REPORTA"))
        barras_h(r.groupby("SECCION_REPORTA").size(), COLOR_RECIBIDO, "SECCION_REPORTA", "c_quien", top=12,
                 total=len(r), resaltar="Calidad", color_resaltado=COLOR_RECIBIDO_CLARO, hover="Recibidos")

    st.subheader("Reportes emitidos por Calidad, por sección y mes &nbsp; :red-badge[Emitidos]", anchor=False)
    with st.container(border=True):
        e, _ = partes(datos("SECCION_DESVIO", "ANIO_MES"))
        heatmap(e, "c_heat", "Emitidos")

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Emitidos y recibidos por categoría**")
        st.caption("Tipo de desvío")
        e, r = partes(datos("CATEGORIA"))
        ce = e["CATEGORIA"].value_counts().reindex(CATEGORIAS, fill_value=0)
        cr = r["CATEGORIA"].value_counts().reindex(CATEGORIAS, fill_value=0)
        op = opacidades(CATEGORIAS, "CATEGORIA")
        fig = go.Figure()
        fig.add_bar(x=CATEGORIAS, y=ce.values, name="Emitidos", marker=mk(COLOR_EMITIDO, op),
                    text=ce.values, textposition="outside", cliponaxis=False,
                    customdata=[cd(CATEGORIA=c) for c in CATEGORIAS])
        fig.add_bar(x=CATEGORIAS, y=cr.values, name="Recibidos", marker=mk(COLOR_RECIBIDO, op),
                    text=cr.values, textposition="outside", cliponaxis=False,
                    customdata=[cd(CATEGORIA=c) for c in CATEGORIAS])
        estilo(fig, 340)
        fig.update_layout(barmode="group")
        fig.update_yaxes(range=[0, max(ce.max(), cr.max(), 1) * 1.2])
        mostrar(fig, "c_cat")
    with b.container(border=True):
        st.markdown("**¿Dónde ocurren los desvíos recibidos?** &nbsp; :blue-badge[Recibidos]")
        st.caption("Sector dentro del área de Calidad")
        _, r = partes(datos("DESVIO_SECTOR"))
        barras_h(r.groupby("DESVIO_SECTOR").size(), COLOR_RECIBIDO, "DESVIO_SECTOR", "c_sector",
                 total=len(r), hover="Recibidos")

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Top 10 reportantes de Calidad** &nbsp; :red-badge[Emitidos]")
        st.caption("Personas de Calidad con más reportes emitidos")
        e, _ = partes(datos("REPORTA_NOMBRE"))
        barras_h(e.groupby("REPORTA_NOMBRE").size(), COLOR_EMITIDO, "REPORTA_NOMBRE", "c_personas", hover="Emitidos")
    with b.container(border=True):
        st.markdown("**Top 10 comportamientos inseguros detectados por Calidad** &nbsp; :red-badge[Emitidos]")
        e, _ = partes(datos("COMPORTAMIENTO_INSEGURO"))
        e = e[(e["CATEGORIA"] == "Acto inseguro") & (e["COMPORTAMIENTO_INSEGURO"] != "")]
        st.caption(f"Actos inseguros emitidos por Calidad ({fmt(len(e))} en total)")
        barras_h(e.groupby("COMPORTAMIENTO_INSEGURO").size(), COLOR_EMITIDO, "COMPORTAMIENTO_INSEGURO",
                 "c_comport", hover="Emitidos", etiquetas_cortas=True)

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Top 10 personas que reportan hacia Calidad** &nbsp; :blue-badge[Recibidos]")
        _, r = partes(datos("REPORTA_NOMBRE"))
        personal_calidad = set(r.loc[r["REPORTA_ES_CALIDAD"] == "Sí", "REPORTA_NOMBRE"])
        st.caption("Quién reporta desvíos ocurridos en el área de Calidad · tono claro: personal de Calidad (autorreporte)")
        barras_h(r.groupby("REPORTA_NOMBRE").size(), COLOR_RECIBIDO, "REPORTA_NOMBRE", "c_personas_r",
                 resaltar=personal_calidad, color_resaltado=COLOR_RECIBIDO_CLARO, hover="Recibidos")
    with b.container(border=True):
        st.markdown("**Top 10 comportamientos inseguros recibidos por Calidad** &nbsp; :blue-badge[Recibidos]")
        _, r = partes(datos("COMPORTAMIENTO_INSEGURO"))
        r = r[(r["CATEGORIA"] == "Acto inseguro") & (r["COMPORTAMIENTO_INSEGURO"] != "")]
        st.caption(f"Actos inseguros ocurridos en el área de Calidad ({fmt(len(r))} en total)")
        barras_h(r.groupby("COMPORTAMIENTO_INSEGURO").size(), COLOR_RECIBIDO, "COMPORTAMIENTO_INSEGURO",
                 "c_comport_r", hover="Recibidos", etiquetas_cortas=True)

    st.subheader("Detalle de reportes de Calidad", anchor=False)
    vista = st.segmented_control("Mostrar", ["Emitidos por Calidad", "Recibidos por Calidad"],
                                 default="Emitidos por Calidad", label_visibility="collapsed")
    det = recib if vista == "Recibidos por Calidad" else emit
    st.caption(f"{fmt(len(det))} registros" + (" · filtrados por la selección" if st.session_state.xf else ""))
    st.dataframe(det[COLS_DETALLE].sort_values("FECHA_REPORTE", ascending=False),
                 width="stretch", hide_index=True, height=360, column_config=CONFIG_DETALLE)

# ===========================================================================
# TAB 2 - PLANTA
# ===========================================================================
with tab2:
    cols = st.columns(6)
    cols[0].metric("Total desvíos", fmt(len(df)), border=True, help=f"{fmt(len(df) / meses_periodo)} por mes")
    for i, c in enumerate(CATEGORIAS):
        n = (df["CATEGORIA"] == c).sum()
        cols[i + 1].metric(c, fmt(n), border=True, help=f"{pct(n, len(df))} del total")
    cols[5].metric("Personas que reportan", fmt(df["REPORTA_ID"].nunique()), border=True)

    st.subheader("Evolución mensual por categoría", anchor=False)
    with st.container(border=True):
        evolucion_categoria(datos("ANIO_MES", "CATEGORIA"), "p_mes")

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Secciones que más reportan**")
        st.caption("Calidad resaltado en rojo")
        d = datos("SECCION_REPORTA")
        barras_h(d.groupby("SECCION_REPORTA").size(), GRIS, "SECCION_REPORTA", "p_reportan", top=12,
                 total=len(d), resaltar="Calidad")
    with b.container(border=True):
        st.markdown("**Secciones con más desvíos**")
        d = datos("SECCION_DESVIO")
        s = d.groupby("SECCION_DESVIO").size().sort_values(ascending=False)
        st.caption(f"Dónde ocurren los desvíos · las 3 primeras concentran el {pct(s.head(3).sum(), s.sum())}")
        barras_h(s, GRAFITO, "SECCION_DESVIO", "p_secciones", top=12, total=len(d), resaltar="Calidad")

    st.subheader("Desvíos por sección y mes", anchor=False)
    with st.container(border=True):
        heatmap(datos("SECCION_DESVIO", "ANIO_MES"), "p_heat")

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Top 10 comportamientos inseguros**")
        st.caption("Actos inseguros de toda la planta")
        d = datos("COMPORTAMIENTO_INSEGURO")
        d = d[(d["CATEGORIA"] == "Acto inseguro") & (d["COMPORTAMIENTO_INSEGURO"] != "")]
        barras_h(d.groupby("COMPORTAMIENTO_INSEGURO").size(), COLOR_CAT["Acto inseguro"], "COMPORTAMIENTO_INSEGURO",
                 "p_comport", etiquetas_cortas=True)
    with b.container(border=True):
        st.markdown("**Origen de los reportes**")
        st.caption("Cómo se detectan los desvíos")
        d = datos("ORIGEN")
        barras_h(d.groupby("ORIGEN").size(), GRAFITO, "ORIGEN", "p_origen", total=len(d))

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Reportes por día de la semana**")
        s = datos("DIA_SEMANA")["DIA_SEMANA"].value_counts().reindex(DIAS, fill_value=0)
        fig = go.Figure(go.Bar(x=s.index, y=s.values, text=s.values, textposition="outside", cliponaxis=False,
                               marker=mk(GRAFITO, opacidades(s.index, "DIA_SEMANA")),
                               customdata=[cd(DIA_SEMANA=x) for x in s.index]))
        estilo(fig, 300, leyenda=False)
        fig.update_yaxes(range=[0, max(s.max(), 1) * 1.2])
        mostrar(fig, "p_dia")
    with b.container(border=True):
        st.markdown("**Reportes por hora del día**")
        horas = [f"{h:02d} h" for h in range(24)]
        s = datos("HORA_TXT")["HORA_TXT"].value_counts().reindex(horas, fill_value=0)
        fig = go.Figure(go.Bar(x=horas, y=s.values, marker=mk(ROJO, opacidades(horas, "HORA_TXT")),
                               customdata=[cd(HORA_TXT=h) for h in horas],
                               hovertemplate="%{x}: %{y}<extra></extra>"))
        estilo(fig, 300, leyenda=False)
        fig.update_xaxes(tickmode="array", tickvals=horas[::2])
        mostrar(fig, "p_hora")

# ===========================================================================
# TAB 3 - SEGUIMIENTO
# ===========================================================================
with tab3:
    abiertos = df[df["ESTADO_CIERRE"] == "Abierto"]
    cerrados = df[df["ESTADO_CIERRE"] == "Cerrado"]
    con_plan = df[df["TIENE_PLAN_ACCION"] == "Sí"]
    ai = df[df["CATEGORIA"] == "Acto inseguro"]
    acepta = (ai["ACEPTA_INFRACCION"] == "Si").sum()
    dias_ab = pd.to_numeric(abiertos["DIAS_ABIERTO"], errors="coerce")
    k = st.columns(5)
    k[0].metric("Desvíos abiertos", fmt(len(abiertos)), border=True, help=f"{pct(len(abiertos), len(df))} del total")
    k[1].metric("Desvíos cerrados", fmt(len(cerrados)), border=True, help=f"{pct(len(cerrados), len(df))} del total")
    k[2].metric("Con plan de acción", fmt(len(con_plan)), border=True, help=f"{pct(len(con_plan), len(df))} del total")
    k[3].metric("Antigüedad media (abiertos)", f"{dias_ab.mean():.0f} días" if dias_ab.notna().any() else "-",
                border=True, help="Días desde la fecha de reporte hasta la extracción del informe.")
    k[4].metric("Aceptación de infracciones", pct(acepta, len(ai)), border=True,
                help=f"{fmt(len(ai) - acepta)} infractores no aceptaron la infracción.")

    a, b = st.columns(2)
    with a.container(border=True):
        st.markdown("**Antigüedad de desvíos abiertos**")
        st.caption("Días desde el reporte")
        d = datos("RANGO_ANTIG")
        s = d[d["ESTADO_CIERRE"] == "Abierto"]["RANGO_ANTIG"].value_counts().reindex(RANGOS_ANTIG, fill_value=0)
        fig = go.Figure(go.Bar(x=RANGOS_ANTIG, y=s.values, text=[fmt(v) for v in s.values], textposition="outside",
                               cliponaxis=False,
                               marker=mk(COLORES_ANTIG, opacidades(RANGOS_ANTIG, "RANGO_ANTIG")),
                               customdata=[cd(RANGO_ANTIG=x) for x in RANGOS_ANTIG]))
        estilo(fig, 320, leyenda=False)
        fig.update_yaxes(range=[0, max(s.max(), 1) * 1.2])
        mostrar(fig, "s_antig")
    with b.container(border=True):
        st.markdown("**Desvíos abiertos con más de 90 días**")
        st.caption("Pendientes por sección, según categoría")
        d = datos("SECCION_DESVIO", "CATEGORIA")
        viejos = d[(d["ESTADO_CIERRE"] == "Abierto") & (pd.to_numeric(d["DIAS_ABIERTO"], errors="coerce") > 90)]
        piv = viejos.pivot_table(index="SECCION_DESVIO", columns="CATEGORIA", values="ID_DESVIO",
                                 aggfunc="count", fill_value=0)
        if piv.empty:
            st.caption("Sin desvíos abiertos de más de 90 días.")
        else:
            piv = piv.loc[piv.sum(axis=1).sort_values(ascending=True).index]
            sel_s, sel_c = seleccion("SECCION_DESVIO"), seleccion("CATEGORIA")
            fig = go.Figure()
            for c in CATEGORIAS:
                if c in piv.columns:
                    op = None
                    if sel_s or sel_c:
                        op = [1.0 if ((not sel_s or sec in sel_s) and (not sel_c or c in sel_c)) else OPACIDAD_ATENUADA
                              for sec in piv.index]
                    fig.add_bar(y=piv.index, x=piv[c], name=c, orientation="h",
                                marker=mk(COLOR_CAT[c], op),
                                customdata=[cd(SECCION_DESVIO=sec, CATEGORIA=c) for sec in piv.index])
            estilo(fig, 320)
            fig.update_layout(barmode="stack", legend=dict(traceorder="normal"))
            mostrar(fig, "s_viejos")

    st.subheader("Planes de acción registrados", anchor=False)
    pa = df[(df["TIENE_PLAN_ACCION"] == "Sí") | (df["ESTADO_CIERRE"] == "Cerrado") | (df["ESTADO"] != "Notificado")]
    if pa.empty:
        st.info("No hay planes de acción en la selección actual.")
    else:
        st.dataframe(pa[["ID_DESVIO", "CATEGORIA", "FECHA_REPORTE", "SECCION_DESVIO", "DESCRIPCION",
                         "ACCION_A_EJECUTAR", "RESPONSABLE_ACCION", "FECHA_PLANIFICADA", "FECHA_CIERRE",
                         "ESTADO", "DIAS_RETRASO"]],
                     width="stretch", hide_index=True,
                     column_config={"ID_DESVIO": "N° desvío", "CATEGORIA": "Categoría", "SECCION_DESVIO": "Sección",
                                    "DESCRIPCION": "Descripción", "ACCION_A_EJECUTAR": "Acción a ejecutar",
                                    "RESPONSABLE_ACCION": "Responsable", "ESTADO": "Estado",
                                    "DIAS_RETRASO": "Días retraso",
                                    "FECHA_REPORTE": st.column_config.DatetimeColumn("Reporte", format="DD/MM/YYYY"),
                                    "FECHA_PLANIFICADA": st.column_config.DatetimeColumn("Planificada", format="DD/MM/YYYY"),
                                    "FECHA_CIERRE": st.column_config.DatetimeColumn("Cierre", format="DD/MM/YYYY")})
    st.caption("La plataforma registra muy pocos planes de acción y cierres: la mayoría de los desvíos "
               "permanece en estado «Notificado».")

# ===========================================================================
# TAB 4 - DATOS
# ===========================================================================
with tab4:
    a, b = st.columns([2, 3])
    alcance = a.selectbox("Registros", ["Todos (filtrados)", "Emitidos por Calidad", "Recibidos por Calidad",
                                        "Relacionados a Calidad"])
    buscar = b.text_input("Buscar texto (descripción, nombre, sector ...)", "")
    base = {"Todos (filtrados)": df, "Emitidos por Calidad": emit, "Recibidos por Calidad": recib,
            "Relacionados a Calidad": df[df["RELACIONADO_CALIDAD"] == "Sí"]}[alcance]
    if buscar:
        mask = pd.Series(False, index=base.index)
        for c in ["DESCRIPCION", "REPORTA_NOMBRE", "DESVIO_SECTOR", "COMPORTAMIENTO_INSEGURO",
                  "INFRACTOR_NOMBRE", "ID_DESVIO"]:
            mask |= base[c].str.contains(buscar, case=False, na=False, regex=False)
        base = base[mask]
    cols = ["ID_DESVIO", "CATEGORIA", "FECHA_REPORTE", "REPORTA_NOMBRE", "REPORTA_SECCION", "ORIGEN",
            "DESVIO_AREA", "DESVIO_SECCION", "DESVIO_SECTOR", "DESCRIPCION", "COMPORTAMIENTO_INSEGURO",
            "INFRACTOR_NOMBRE", "INFRACTOR_TIPO_PERSONAL", "ESTADO", "REPORTA_ES_CALIDAD", "DESVIO_EN_CALIDAD"]
    st.caption(f"{fmt(len(base))} registros" + (" · filtrados por la selección de gráficos" if st.session_state.xf else ""))
    st.dataframe(base[cols].sort_values("FECHA_REPORTE", ascending=False), width="stretch",
                 hide_index=True, height=520,
                 column_config={"FECHA_REPORTE": st.column_config.DatetimeColumn("Fecha", format="DD/MM/YYYY HH:mm"),
                                "ID_DESVIO": "N° desvío", "CATEGORIA": "Categoría", "REPORTA_NOMBRE": "Reporta",
                                "REPORTA_SECCION": "Sección reporta", "ORIGEN": "Origen", "DESVIO_AREA": "Área desvío",
                                "DESVIO_SECCION": "Sección desvío", "DESVIO_SECTOR": "Sector",
                                "DESCRIPCION": st.column_config.TextColumn("Descripción", width="large"),
                                "COMPORTAMIENTO_INSEGURO": "Comportamiento inseguro", "INFRACTOR_NOMBRE": "Infractor",
                                "INFRACTOR_TIPO_PERSONAL": "Tipo infractor", "ESTADO": "Estado",
                                "REPORTA_ES_CALIDAD": "¿Emitido por Calidad?",
                                "DESVIO_EN_CALIDAD": "¿Recibido por Calidad?"})
    st.download_button("Descargar Excel", a_excel(base, cols), icon=":material/download:",
                       file_name=f"desvios_{alcance.split()[0].lower()}_{datetime.now():%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.caption("Embol S.A. · Aseguramiento de Calidad · Información generada a partir de la Plataforma Web "
           "de Desvíos SSMA y Calidad.")

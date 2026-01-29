import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
import matplotlib.patheffects as pe 
import textwrap
import io
import sys
import requests
from datetime import datetime, timedelta
import numpy as np

# ==========================================
# 0. CONFIGURACIÓN Y ENLACE NUBE
# ==========================================
st.set_page_config(layout="wide", page_title="Línea de Tiempo Automática")

# 👇👇👇 PEGA AQUÍ TU ENLACE DE GOOGLE SHEETS O ONEDRIVE 👇👇👇
# Enlace proporcionado por ti:
URL_ORIGINAL = "https://colbun-my.sharepoint.com/personal/ep_tvaldes_colbun_cl/_layouts/15/guestaccess.aspx?share=IQD3gVYvlakxQJSzuVvTQAR4AcK2dfpMmRikeD4OSW0kSEE&e=muZP0V"

# Función para intentar convertir el link de vista a link de descarga directa
def transformar_url_onedrive(url):
    if "sharepoint.com" in url or "onedrive.live.com" in url:
        # Intento 1: Reemplazar guestaccess.aspx por download.aspx
        if "guestaccess.aspx" in url:
            return url.replace("guestaccess.aspx", "download.aspx")
        # Intento 2: Añadir &download=1 si no lo tiene (común en enlaces nuevos)
        if "download=1" not in url:
            if "?" in url:
                return url + "&download=1"
            else:
                return url + "?download=1"
    return url

URL_ARCHIVO_NUBE = transformar_url_onedrive(URL_ORIGINAL)

# ==========================================
# 1. FUNCIONES DE CARGA Y CACHÉ
# ==========================================

# CORRECCIÓN APLICADA: Usamos cache_resource en lugar de cache_data
# cache_resource es para objetos como conexiones o archivos abiertos (ExcelFile)
@st.cache_resource(ttl=60)
def cargar_datos_desde_nube(url):
    try:
        # Usamos requests para bajar el contenido binario primero
        response = requests.get(url, timeout=10) # Timeout de seguridad
        response.raise_for_status() # Lanza error si el link falla (404, 403)
        
        # Leemos el contenido como archivo Excel
        file_content = io.BytesIO(response.content)
        xl_file = pd.ExcelFile(file_content)
        return xl_file
    except Exception as e:
        return None

def normalizar_columnas(df):
    df.columns = df.columns.str.strip()
    mapa_cols = {
        'Fecha_Vigente': ['Fecha_Vigente', 'Fecha Vigente', 'Fecha Real', 'Fecha_Real_Manual', 'Fecha Actual', 'Fecha_Real'],
        'Fecha_teorica': ['Fecha_teórica', 'Fecha_teorica', 'Fecha Teórica', 'Fecha Teorica', 'Fecha_Proyectada', 'Fecha Proyectada', 'Fecha Planificada'],
        'Hito / Etapa': ['Hito / Etapa', 'Hito', 'Etapa', 'Nombre Hito', 'Actividad'],
        'Agente': ['Agente', 'Responsable', 'Actor', 'Encargado']
    }
    renombres = {}
    for estandar, variantes in mapa_cols.items():
        for variante in variantes:
            if variante in df.columns:
                renombres[variante] = estandar
                break 
    df = df.rename(columns=renombres)
    return df

def fecha_es(fecha, formato="corto"):
    if pd.isnull(fecha): return ""
    meses = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}
    meses_full = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
    if formato == "corto": return f"{fecha.day}-{meses[fecha.month]}"
    elif formato == "eje": return f"{meses[fecha.month]}-{str(fecha.year)[2:]}"
    elif formato == "hoy_full": return f"{fecha.day}/{meses_full[fecha.month]}/{fecha.year}"
    return f"{fecha.day}/{fecha.month}/{fecha.year}"

def requiere_formato_arbol(df, col_fecha='Fecha_Vigente'):
    if df.empty: return False
    conteo = df[col_fecha].value_counts()
    return (conteo > 1).any()

# ==========================================
# 2. MOTORES GRÁFICOS
# ==========================================

def graficar_modo_arbol(df_plot, titulo, f_inicio, f_fin, mapa_colores, mostrar_hoy, tipo_rango):
    fig, ax = plt.subplots(figsize=(16, 9), constrained_layout=True)
    ax.axhline(0, color="#34495e", linewidth=2, zorder=1)
    plt.figtext(0.015, 0.98, f"Generado: {datetime.now().strftime('%d/%m/%Y')}", fontsize=10, color='#555555')

    ANCHO_CAJA_DIAS = max(25, (f_fin - f_inicio).days * 0.08)
    OFFSET_CONEXION_DIAS = ANCHO_CAJA_DIAS * 0.05
    NIVEL_MIN_SINGLE = 3.5; STEP_SINGLE = 2.5; NIVEL_MIN_ARBOL = 5.0; STEP_ARBOL = 3.0

    cajas_ocupadas = []; elementos_finales = []
    grupos = df_plot.groupby('Fecha_Vigente')
    lista_singles = []; lista_arboles = []
    for fecha, grupo in grupos:
        if len(grupo) == 1: lista_singles.append((fecha, grupo.iloc[0]))
        else: lista_arboles.append((fecha, grupo))

    for i, (fecha, row) in enumerate(lista_singles):
        es_positivo = (i % 2 == 0)
        nivel_base = NIVEL_MIN_SINGLE if es_positivo else -NIVEL_MIN_SINGLE
        encontrado = False; intentos = 0; nivel_actual = nivel_base
        while intentos < 20:
            y_min, y_max = nivel_actual - 1.0, nivel_actual + 1.0
            x_min, x_max = fecha - timedelta(days=ANCHO_CAJA_DIAS/2), fecha + timedelta(days=ANCHO_CAJA_DIAS/2)
            colision = False
            for (ox1, ox2, oy1, oy2) in cajas_ocupadas:
                if (x_min < ox2 and x_max > ox1) and (y_min < oy2 and y_max > oy1): colision = True; break
            if not colision:
                cajas_ocupadas.append((x_min, x_max, y_min, y_max))
                elementos_finales.append({'tipo': 'single', 'row': row, 'x': fecha, 'y': nivel_actual})
                encontrado = True; break
            if nivel_actual > 0: nivel_actual += STEP_SINGLE
            else: nivel_actual -= STEP_SINGLE
            intentos += 1
        if not encontrado: elementos_finales.append({'tipo': 'single', 'row': row, 'x': fecha, 'y': nivel_actual})

    for i_arbol, (fecha, grupo) in enumerate(lista_arboles):
        cantidad = len(grupo); trunk_dir = 1 if i_arbol % 2 == 0 else -1
        altura_base_tronco = NIVEL_MIN_ARBOL; encontrado_tronco = False; intentos_tronco = 0
        while intentos_tronco < 15:
            colision_arbol_entero = False; temp_cajas_ramas = []; temp_posiciones = []
            alto_total = altura_base_tronco + (cantidad - 1) * STEP_ARBOL
            y_fin_tronco = alto_total * trunk_dir
            grupo = grupo.reset_index(drop=True)
            for i_hito, row in grupo.iterrows():
                y_nivel = (altura_base_tronco + (i_hito * STEP_ARBOL)) * trunk_dir
                es_derecha = (i_hito % 2 == 0); dir_h = 1 if es_derecha else -1
                x_caja = fecha + timedelta(days=ANCHO_CAJA_DIAS * dir_h)
                y_min_box, y_max_box = y_nivel - 1.2, y_nivel + 1.2
                x_min_box, x_max_box = min(fecha, x_caja) - timedelta(days=5), max(fecha, x_caja) + timedelta(days=5)
                for (ox1, ox2, oy1, oy2) in cajas_ocupadas:
                    if (x_min_box < ox2 and x_max_box > ox1) and (y_min_box < oy2 and y_max_box > oy1):
                        colision_arbol_entero = True; break
                if colision_arbol_entero: break
                temp_cajas_ramas.append((x_min_box, x_max_box, y_min_box, y_max_box))
                temp_posiciones.append({'row': row, 'y_nivel': y_nivel, 'x_caja': x_caja, 'es_derecha': es_derecha})
            if not colision_arbol_entero:
                cajas_ocupadas.extend(temp_cajas_ramas)
                elementos_finales.append({'tipo': 'arbol', 'fecha': fecha, 'y_fin_tronco': y_fin_tronco, 'ramas': temp_posiciones, 'agente_raiz': grupo.iloc[0]['Agente']})
                encontrado_tronco = True; break
            altura_base_tronco += STEP_ARBOL; intentos_tronco += 1
        if not encontrado_tronco:
             elementos_finales.append({'tipo': 'arbol', 'fecha': fecha, 'y_fin_tronco': y_fin_tronco, 'ramas': temp_posiciones, 'agente_raiz': grupo.iloc[0]['Agente']})

    max_abs_y = 4.0
    for item in elementos_finales:
        if item['tipo'] == 'single':
            row = item['row']; x = item['x']; y = item['y']
            if abs(y) > max_abs_y: max_abs_y = abs(y)
            agente = str(row.get('Agente', 'N/A')); color = mapa_colores.get(agente, '#7f8c8d')
            ax.vlines(x, 0, y, color=color, alpha=0.5, linewidth=1, linestyle='-', zorder=1)
            ax.scatter(x, 0, s=60, color=color, marker='o', zorder=3)
            f_teorica = row.get('Fecha_teorica', pd.NaT)
            if pd.notnull(f_teorica) and abs((x - f_teorica).days) > 5:
                carril = 1.0 if y > 0 else -1.0
                ax.annotate("", xy=(x, carril), xytext=(max(f_teorica, f_inicio), carril), arrowprops=dict(arrowstyle="->", color='#555555', lw=0.9), zorder=5)
                dias = (x - f_teorica).days
                pos_txt = max(max(f_teorica, f_inicio), x - timedelta(days=6))
                ax.text(pos_txt, carril-0.25, f"{'+' if dias>0 else ''}{dias}d", ha='center', va='top', fontsize=7, color='#555555', fontweight='bold', zorder=30).set_path_effects([pe.withStroke(linewidth=2.0, foreground='white')])
            texto_lbl = f"{textwrap.fill(agente.upper(), 20)}\n{textwrap.fill(str(row.get('Hito / Etapa','')), 25)}\n{fecha_es(x)}"
            ax.annotate(texto_lbl, xy=(x, y), xytext=(x, y), bbox=dict(boxstyle="round,pad=0.4", fc="white", ec=color, lw=1.5, alpha=0.95), ha='center', va='center', fontsize=8, color='#2c3e50', zorder=10)
        elif item['tipo'] == 'arbol':
            fecha = item['fecha']; y_fin = item['y_fin_tronco']; ramas = item['ramas']; agente_raiz = item['agente_raiz']
            if abs(y_fin) > max_abs_y: max_abs_y = abs(y_fin)
            color_raiz = mapa_colores.get(agente_raiz, '#34495e')
            ax.scatter(fecha, 0, s=80, color=color_raiz, marker='o', zorder=4, edgecolor='white', linewidth=1.5)
            ax.vlines(fecha, 0, y_fin, color='#7f8c8d', alpha=0.5, linewidth=2, zorder=1, linestyle='--')
            for rama in ramas:
                row = rama['row']; y_nivel = rama['y_nivel']; x_caja = rama['x_caja']; es_derecha = rama['es_derecha']
                if abs(y_nivel) > max_abs_y: max_abs_y = abs(y_nivel)
                agente = str(row.get('Agente', 'N/A')); color = mapa_colores.get(agente, '#7f8c8d')
                offset_gap = timedelta(days=OFFSET_CONEXION_DIAS)
                x_linea_fin = x_caja - offset_gap if es_derecha else x_caja + offset_gap
                ax.plot([fecha, x_linea_fin], [y_nivel, y_nivel], color=color, linewidth=1.5, zorder=2)
                ax.scatter(fecha, y_nivel, s=30, color=color, zorder=3)
                texto_lbl = f"{textwrap.fill(agente.upper(), 20)}\n{textwrap.fill(str(row.get('Hito / Etapa','')), 25)}\n{fecha_es(fecha)}"
                ax.annotate(texto_lbl, xy=(x_caja, y_nivel), xytext=(x_caja, y_nivel), bbox=dict(boxstyle="round,pad=0.4", fc="white", ec=color, lw=1.5, alpha=1.0), ha='center', va='center', fontsize=7.5, color='#2c3e50', zorder=10)
                f_teorica = row.get('Fecha_teorica', pd.NaT)
                if pd.notnull(f_teorica) and abs((fecha - f_teorica).days) > 3:
                    offset_flecha = 1.2 if y_nivel > 0 else -1.2
                    y_flecha = y_nivel + offset_flecha
                    if abs(y_flecha) > max_abs_y: max_abs_y = abs(y_flecha)
                    f_ini_vis = max(f_teorica, f_inicio)
                    if f_teorica >= f_inicio: ax.vlines(f_teorica, 0, y_flecha, color='#bdc3c7', alpha=0.6, linestyles=':', zorder=1); ax.scatter(f_teorica, 0, s=20, color='#bdc3c7', marker='|', zorder=1)
                    ax.annotate("", xy=(x_caja, y_flecha), xytext=(f_ini_vis, y_flecha), arrowprops=dict(arrowstyle="->", color='#555555', lw=0.9), zorder=5)
                    dias = (fecha - f_teorica).days
                    pos_txt = max(f_ini_vis, x_caja - timedelta(days=6)) if x_caja > f_ini_vis else min(f_ini_vis, x_caja + timedelta(days=6))
                    va_txt = 'top' if y_nivel > 0 else 'bottom'; offset_txt_y = -0.25 if y_nivel > 0 else 0.25
                    ax.text(pos_txt, y_flecha + offset_txt_y, f"{'+' if dias>0 else ''}{dias}d", ha='center', va=va_txt, fontsize=7, color='#555555', fontweight='bold', zorder=30).set_path_effects([pe.withStroke(linewidth=2.0, foreground='white')])

    ax.spines['left'].set_visible(False); ax.spines['right'].set_visible(False); ax.spines['top'].set_visible(False); ax.yaxis.set_visible(False)
    ax.set_xlim(f_inicio, f_fin)
    margen_y_final = max_abs_y + 3.0
    ax.set_ylim(-margen_y_final, margen_y_final)

    if mostrar_hoy and (f_inicio <= datetime.now() <= f_fin):
        hoy = datetime.now()
        ax.axvline(hoy, color='#e74c3c', ls='--', alpha=0.8, linewidth=1.5, zorder=0)
        offset_dias_hoy = (f_fin - f_inicio).days * 0.008
        ax.text(hoy - timedelta(days=offset_dias_hoy), margen_y_final - 0.5, f"HOY\n{fecha_es(hoy, 'hoy_full')}", color='#e74c3c', ha='right', va='top', fontweight='bold', fontsize=9)

    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x,p: fecha_es(mdates.num2date(x), "eje")))
    plt.title(f"Línea de Tiempo: {titulo}", fontsize=18, fontweight='bold', color='#2c3e50', pad=20)
    leyenda = [Patch(facecolor=mapa_colores.get(a, '#7f8c8d'), label=a) for a in df_plot['Agente'].unique()]
    leyenda.append(Line2D([0],[0], color='#555555', lw=1, marker='>', label='Días Retraso'))
    ax.legend(handles=leyenda, title="Agentes Responsables", loc='upper center', bbox_to_anchor=(0.5, -0.14), ncol=4, fancybox=True, shadow=True)
    if tipo_rango == 3:
        txt_rango = f"Periodo Personalizado:\n{f_inicio.strftime('%d/%m/%Y')} - {f_fin.strftime('%d/%m/%Y')}"
        ax.text(0.0, -0.14, txt_rango, transform=ax.transAxes, fontsize=9, color='#555555', va='top', ha='left', bbox=dict(boxstyle="round,pad=0.4", fc="#ecf0f1", ec="#bdc3c7", lw=1))
    return fig

def graficar_modo_estandar(df_plot, titulo, f_inicio, f_fin, mapa_colores, mostrar_hoy, tipo_rango):
    col_vigente, col_teorica = 'Fecha_Vigente', 'Fecha_teorica'
    ocupacion_niveles, niveles_asignados = {}, []
    BASE_NIVEL_POS = 5.0; BASE_NIVEL_NEG = -5.0; STEP_NIVEL = 2.5; MARGEN_DIAS = 75
    niveles_flechas_pos, niveles_flechas_neg = [0.8, 1.6, 2.4], [-0.8, -1.6, -2.4]
    ocupacion_flechas = {lvl: [] for lvl in niveles_flechas_pos + niveles_flechas_neg}

    for index, row in df_plot.iterrows():
        fecha = row[col_vigente]; es_positivo = (index % 2 == 0)
        nivel_actual = BASE_NIVEL_POS if es_positivo else BASE_NIVEL_NEG
        encontrado, intentos = False, 0
        while intentos < 20:
            colision = False
            if nivel_actual in ocupacion_niveles:
                for f_ocupada in ocupacion_niveles[nivel_actual]:
                    if abs((fecha - f_ocupada).days) < MARGEN_DIAS: colision = True; break
            if not colision:
                if nivel_actual not in ocupacion_niveles: ocupacion_niveles[nivel_actual] = []
                ocupacion_niveles[nivel_actual].append(fecha)
                niveles_asignados.append(nivel_actual); encontrado = True; break
            if es_positivo: nivel_actual += STEP_NIVEL
            else: nivel_actual -= STEP_NIVEL
            intentos += 1
        if not encontrado:
            niveles_asignados.append(nivel_actual)
            if nivel_actual not in ocupacion_niveles: ocupacion_niveles[nivel_actual] = []
            ocupacion_niveles[nivel_actual].append(fecha)
    
    df_plot['nivel'] = niveles_asignados

    def obtener_carril_flecha(f_inicio, f_fin, es_arriba):
        carriles = niveles_flechas_pos if es_arriba else niveles_flechas_neg
        start, end = min(f_inicio, f_fin), max(f_inicio, f_fin)
        for carril in carriles:
            libre = True
            for (o_start, o_end) in ocupacion_flechas[carril]:
                if (start <= o_end + timedelta(days=5)) and (end >= o_start - timedelta(days=5)): libre = False; break
            if libre: ocupacion_flechas[carril].append((start, end)); return carril
        return carriles[len(carriles) // 2]

    fig, ax = plt.subplots(figsize=(16, 9), constrained_layout=True)
    ax.axhline(0, color="#34495e", linewidth=2, zorder=1)
    plt.figtext(0.015, 0.98, f"Generado: {datetime.now().strftime('%d/%m/%Y')}", fontsize=10, color='#555555')

    for i, row in df_plot.iterrows():
        f_vigente = row[col_vigente]; f_teorica = row[col_teorica]; nivel = row['nivel']
        agente = str(row['Agente']); color = mapa_colores.get(agente, '#7f8c8d')
        
        ax.vlines(f_vigente, 0, nivel, color=color, alpha=0.5, linewidth=1, linestyle='-', zorder=1)
        ax.scatter(f_vigente, 0, s=60, color=color, marker='o', zorder=3)

        if pd.notnull(f_teorica):
            dias = (f_vigente - f_teorica).days
            if abs(dias) > 5:
                es_arriba = (nivel > 0)
                altura_cota = obtener_carril_flecha(f_teorica, f_vigente, es_arriba)
                f_ini_vis = max(f_teorica, f_inicio)
                if f_teorica >= f_inicio:
                    ax.vlines(f_teorica, 0, altura_cota, color='#bdc3c7', alpha=0.8, linestyles=':', zorder=1)
                    ax.scatter(f_teorica, 0, s=20, color='#bdc3c7', marker='|', zorder=2)
                ax.annotate("", xy=(f_vigente, altura_cota), xytext=(f_ini_vis, altura_cota), arrowprops=dict(arrowstyle="->", color='#555555', lw=0.9), zorder=20)
                pos_txt = max(f_ini_vis, f_vigente - timedelta(days=6))
                signo = "+" if dias > 0 else ""
                ax.text(pos_txt, altura_cota - 0.25, f"{signo}{dias}d", ha='center', va='top', fontsize=7, color='#555555', fontweight='bold', zorder=30).set_path_effects([pe.withStroke(linewidth=2.0, foreground='white')])

        texto_lbl = f"{textwrap.fill(agente.upper(), 20)}\n{textwrap.fill(str(row['Hito / Etapa']), 25)}\n{fecha_es(f_vigente)}"
        ax.annotate(texto_lbl, xy=(f_vigente, nivel), xytext=(f_vigente, nivel), bbox=dict(boxstyle="round,pad=0.4", fc="white", ec=color, lw=1.5, alpha=0.95), ha='center', va='center', fontsize=8, color='#2c3e50', zorder=10)

    ax.spines['left'].set_visible(False); ax.spines['right'].set_visible(False); ax.spines['top'].set_visible(False); ax.yaxis.set_visible(False)
    ax.set_xlim(f_inicio, f_fin)
    
    max_y = df_plot['nivel'].max() if not df_plot['nivel'].empty else 4
    min_y = df_plot['nivel'].min() if not df_plot['nivel'].empty else -4
    limite_superior = max(8, max_y + 3.0); limite_inferior = min(-8, min_y - 3.0)
    ax.set_ylim(limite_inferior, limite_superior)

    if mostrar_hoy and (f_inicio <= datetime.now() <= f_fin):
        hoy = datetime.now()
        ax.axvline(hoy, color='#e74c3c', linestyle='--', alpha=0.8, linewidth=1.5, zorder=0)
        offset_dias_hoy = (f_fin - f_inicio).days * 0.008
        ax.text(hoy - timedelta(days=offset_dias_hoy), limite_superior * 0.95, f"HOY\n{fecha_es(hoy, 'hoy_full')}", color='#e74c3c', fontsize=9, fontweight='bold', ha='right', va='top')

    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: fecha_es(mdates.num2date(x), "eje")))
    plt.title(f"Línea de Tiempo: {titulo}", fontsize=18, fontweight='bold', color='#2c3e50', pad=25)
    
    leyenda = [Patch(facecolor=mapa_colores.get(a, '#7f8c8d'), label=a) for a in df_plot['Agente'].unique()]
    leyenda.append(Line2D([0],[0], color='#555555', lw=1, marker='>', label='Días Retraso'))
    ax.legend(handles=leyenda, title="Agentes Responsables", loc='upper center', bbox_to_anchor=(0.5, -0.14), ncol=4, fancybox=True, shadow=True)

    if tipo_rango == 3:
        txt_rango = f"Periodo Personalizado:\n{f_inicio.strftime('%d/%m/%Y')} - {f_fin.strftime('%d/%m/%Y')}"
        ax.text(0.0, -0.14, txt_rango, transform=ax.transAxes, fontsize=9, color='#555555', va='top', ha='left', bbox=dict(boxstyle="round,pad=0.4", fc="#ecf0f1", ec="#bdc3c7", lw=1))

    return fig

# ==========================================
# 3. INTERFAZ STREAMLIT
# ==========================================

st.title("📊 Generador de Líneas de Tiempo - Normativas")

# CARGA DE ARCHIVO: Aquí Streamlit leerá la URL o mostrará error si no es válido
try:
    if not URL_ARCHIVO_NUBE or "PON_AQUI" in URL_ARCHIVO_NUBE:
        st.warning("⚠️ No se ha configurado la URL del archivo de Excel. Por favor, edite la variable 'URL_ARCHIVO_NUBE' en el código.")
    else:
        # Intentamos cargar el archivo desde la nube
        xl_file = cargar_datos_desde_nube(URL_ARCHIVO_NUBE)
        
        if xl_file is None:
            st.error("❌ Error al descargar el archivo. Verifique que el enlace sea público y de descarga directa.")
        else:
            hojas = [h for h in xl_file.sheet_names if not h.startswith('_')]
            
            st.sidebar.header("⚙️ Configuración")
            hoja_seleccionada = st.sidebar.selectbox("Seleccione Normativa:", hojas)
            
            opcion_fecha = st.sidebar.radio("Rango de Fechas:", ("Año Calendario Actual", "Ventana Móvil (-12/+12 meses)", "Personalizado"))
            
            hoy = datetime.now()
            tipo_rango = 0
            
            if opcion_fecha == "Año Calendario Actual":
                f_inicio = datetime(hoy.year, 1, 1)
                f_fin = datetime(hoy.year, 12, 31)
                tipo_rango = 1
            elif opcion_fecha == "Ventana Móvil (-12/+12 meses)":
                f_inicio = hoy - timedelta(days=365)
                f_fin = hoy + timedelta(days=365)
                tipo_rango = 2
            else: 
                col1, col2 = st.sidebar.columns(2)
                d_inicio = col1.date_input("Inicio", hoy - timedelta(days=30))
                d_fin = col2.date_input("Fin", hoy + timedelta(days=30))
                f_inicio = pd.to_datetime(d_inicio)
                f_fin = pd.to_datetime(d_fin)
                tipo_rango = 3
                
            mostrar_hoy = st.sidebar.checkbox("Mostrar línea de HOY", value=True)
            
            df = pd.read_excel(xl_file, sheet_name=hoja_seleccionada)
            df = normalizar_columnas(df)
            
            filtro_proceso = "Todo"
            if 'Hito / Etapa' in df.columns:
                df['Hito_Upper'] = df['Hito / Etapa'].astype(str).str.upper()
                tiene_zonal = df['Hito_Upper'].str.contains('ZONAL').any()
                tiene_nacional = df['Hito_Upper'].str.contains('NACIONAL').any()
                if tiene_zonal and tiene_nacional:
                    filtro_proceso = st.sidebar.radio("Filtro Proceso:", ("Todo", "Zonal", "Nacional"))
            
            if st.sidebar.button("Generar Gráfico"):
                with st.spinner('Generando visualización...'):
                    if filtro_proceso == "Zonal":
                        df = df[df['Hito_Upper'].str.contains('ZONAL') | df['Hito_Upper'].str.contains('COMÚN')].copy()
                    elif filtro_proceso == "Nacional":
                        df = df[df['Hito_Upper'].str.contains('NACIONAL') | df['Hito_Upper'].str.contains('COMÚN')].copy()
                    
                    if 'Fecha_Vigente' not in df.columns:
                        st.error("❌ El archivo no tiene la columna 'Fecha Vigente'.")
                    else:
                        df['Fecha_Vigente'] = pd.to_datetime(df['Fecha_Vigente'], errors='coerce')
                        df['Fecha_teorica'] = pd.to_datetime(df.get('Fecha_teorica', pd.NaT), errors='coerce')
                        df = df.dropna(subset=['Fecha_Vigente'])
                        
                        df_plot = df[(df['Fecha_Vigente'] >= f_inicio) & (df['Fecha_Vigente'] <= f_fin)].copy()
                        
                        if df_plot.empty:
                            st.warning("⚠️ No hay datos en el rango de fechas seleccionado.")
                        else:
                            mapa_colores = {
                                'Comisión': "#0400ff", 'CNE': "#0400ff", 'Participantes': '#27ae60', 
                                'Participantes e interesados': '#27ae60', 'Interesados': '#27ae60', 
                                'Empresas': '#27ae60', 'Coordinador': '#e67e22', 'CEN': '#e67e22'
                            }
                            colores_extra = ["#ff0000", "#8e44ad", "#16a085", "#34495e", "#d35400"]
                            idx_extra = 0
                            for ag in df_plot['Agente'].unique():
                                if ag and ag not in mapa_colores:
                                    mapa_colores[ag] = colores_extra[idx_extra % len(colores_extra)]
                                    idx_extra += 1
                            
                            titulo_limpio = hoja_seleccionada.replace('_', ' ')
                            
                            if requiere_formato_arbol(df_plot):
                                fig = graficar_modo_arbol(df_plot, titulo_limpio, f_inicio, f_fin, mapa_colores, mostrar_hoy, tipo_rango)
                            else:
                                fig = graficar_modo_estandar(df_plot, titulo_limpio, f_inicio, f_fin, mapa_colores, mostrar_hoy, tipo_rango)
                            
                            st.pyplot(fig)
                            
                            fn = f"timeline_{titulo_limpio}.png"
                            img = io.BytesIO()
                            plt.savefig(img, format='png', dpi=400, bbox_inches='tight', pad_inches=0.2)
                            img.seek(0)
                            st.download_button(label="💾 Descargar Imagen HD", data=img, file_name=fn, mime="image/png")

except Exception as e:
    st.error(f"Error al procesar el archivo: {e}")

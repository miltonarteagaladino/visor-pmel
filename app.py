import os
import re
import textwrap
import itertools
from collections import Counter, defaultdict
import pandas as pd
import streamlit as st
import openpyxl
import plotly.graph_objects as go
import plotly.express as px
from pyvis.network import Network
import streamlit.components.v1 as components

# BLOQUEO ABSOLUTO DE PYARROW (ANTI-CRASH PARA MAC)
os.environ["ARROW_USER_SIMD_LEVEL"] = "NONE"
os.environ["STREAMLIT_SERVER_MAX_MESSAGE_SIZE"] = "200"

st.set_page_config(layout="wide", page_title="PMEL - Fundación Corona", initial_sidebar_state="expanded")

# --- 1. IDENTIDAD CORPORATIVA Y NAVEGACIÓN ---
st.markdown("""
    <style>
    .stApp { background-color: #F8F9FA; }
    .corona-header { background-color: #003366 !important; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .corona-title { color: #FFFFFF !important; margin: 0; font-family: 'Helvetica Neue', sans-serif; font-size: 32px; font-weight: 700; }
    .corona-subtitle { color: #FFB300 !important; margin: 5px 0 0 0; font-family: 'Helvetica Neue', sans-serif; font-size: 16px; }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        pass
        
    st.markdown("### Menú de Navegación")
    pagina_actual = st.radio("Ir a:", [
        "🕸️ Mapa Sistémico (Redes)", 
        "📊 Analítica de Portafolio", 
        "🧬 Patrones de Co-Ocurrencia"
    ])
    st.markdown("---")

titulo_limpio = pagina_actual.split(" ", 1)[1] if " " in pagina_actual else pagina_actual
st.markdown(f"""
    <div class="corona-header">
        <h1 class="corona-title">{titulo_limpio}</h1>
        <p class="corona-subtitle">ESTRATEGIA 2030 | Fundación Corona</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. DECODIFICADORES ---
MAPA_AREAS = {
    'ED': 'Educación', 'EM': 'Empleo', 'LE': 'Libre Elección', 'IN': 'Incidencia',
    'AA': 'Aprendizaje/Adaptación', 'TC': 'Transformación Cultural', 'CC': 'Comunicaciones',
    'DE': 'Dirección Ejecutiva', 'SA': 'Servicios Administrativos'
}

def extraer_area_codigo(codigo):
    partes = codigo.split('-')
    for p in partes:
        if p in MAPA_AREAS: return f"{p} - {MAPA_AREAS[p]}"
    for clave, valor in MAPA_AREAS.items():
        if f"-{clave}-" in codigo or f"-{clave}" in codigo: return f"{clave} - {valor}"
    return "Múltiples / Otra Área"

def detectar_accion_oficial(texto):
    t = str(texto).lower()
    if "analiza" in t or "comprend" in t: return "1. Analizamos y comprendemos el contexto"
    if "conocimiento" in t or "evidencia" in t: return "2. Generamos conocimiento y evidencia"
    if "agenda" in t: return "3. Posicionamos agendas"
    if "colectiv" in t: return "4. Promovemos acciones colectivas"
    if "diseño" in t or "solucion" in t: return "5. Promovemos el diseño y desarrollo de soluciones"
    if "capacidad" in t or "lider" in t: return "6. Fortalecemos capacidades (liderazgo)"
    return None

def obtener_color_borde_categoria(col_index):
    if col_index == 4: return "#9C27B0" 
    elif 5 <= col_index <= 10: return "#C0CA33" 
    elif 11 <= col_index <= 18: return "#D81B60" 
    elif 19 <= col_index <= 22: return "#FFB300" 
    elif 23 <= col_index <= 26: return "#4CAF50" 
    return "#BDBDBD"

def limpiar_codigo_historia(codigo):
    c_limpio = str(codigo).strip().upper().replace(" ", "")
    c_limpio = c_limpio.replace(".", "-").replace("--", "-")
    if "02EM-FC" in c_limpio: c_limpio = c_limpio.replace("02EM-FC", "02-EM-FC")
    return c_limpio

def obtener_estado_por_sufijo(codigo):
    cod_upper = str(codigo).upper().strip()
    if cod_upper.endswith("-EC"): return 'Negro (Ejemplo de Cambio)'
    if cod_upper.endswith("-BC"): return 'Rojo (Señal de Buen Camino)'
    if cod_upper.endswith("-IC"): return 'Naranja (Intención de Cambio)'
    if cod_upper.endswith("-EE"): return 'Morado (Efecto Estancado)'
    return 'Negro (Ejemplo de Cambio)' 

def extraer_corazon_codigo(codigo):
    c_limpio = limpiar_codigo_historia(codigo)
    if c_limpio.endswith(("-EC", "-IC", "-EE", "-BC")):
        c_limpio = c_limpio[:-3]
        
    match = re.search(r'([A-Z]+-\w+-\d+-[A-Z]+-[A-Z]+)', c_limpio)
    if match: return match.group(1)
    
    match2 = re.search(r'([A-Z]+-\d+)', c_limpio)
    if match2: return match2.group(1)
    
    return c_limpio

def acortar_codigo(codigo):
    corazon = extraer_corazon_codigo(codigo)
    partes = corazon.split('-')
    if len(partes) >= 3: return f"{partes[1]}{partes[2]}"
    return corazon

def formatear_caja(texto, ancho=35):
    lineas = textwrap.wrap(texto, width=ancho)
    if len(lineas) > 3: return "\n".join(lineas[:3]) + "..."
    return "\n".join(lineas)

# EXTRACCIÓN ESTRUCTURAL REVISADA
@st.cache_data
def extraer_datos_puros(ruta_archivo):
    wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
    datos = []
    pestañas = [sht for sht in wb.sheetnames if sht.startswith('I-') or sht.startswith('M-') or sht.startswith('B-') or sht.startswith('H-')]
    
    for sheet_name in pestañas:
        ws = wb[sheet_name]
        iniciativa = str(ws.cell(row=1, column=2).value or sheet_name).strip()
        
        fila_cabeceras = -1
        for r in range(1, min(30, ws.max_row + 1)):
            val = ws.cell(row=r, column=1).value
            if val and str(val).strip().lower().startswith("acciones est"):
                fila_cabeceras = r
                break
                
        if fila_cabeceras == -1:
            continue
            
        colores_cambios_font = {}
        for col in range(2, ws.max_column + 1):
            val_cambio = ws.cell(row=fila_cabeceras, column=col).value
            if val_cambio and str(val_cambio).strip() != 'None':
                c_texto = str(val_cambio).replace('\n', ' ').strip()
                colores_cambios_font[c_texto] = obtener_color_borde_categoria(col)

        textos_hist = {}
        for row in range(1, fila_cabeceras):
            cod = ws.cell(row=row, column=1).value
            txt = ws.cell(row=row, column=2).value
            if cod and isinstance(cod, str) and '-' in cod:
                c_limpio = extraer_corazon_codigo(cod)
                textos_hist[c_limpio] = str(txt).strip() if txt else "Sin narrativa documentada."

        accion_actual = None
        for row in range(fila_cabeceras + 1, ws.max_row + 1):
            val_accion = ws.cell(row=row, column=1).value
            val_str = str(val_accion).strip() if val_accion else ""
            
            if val_str.lower().startswith('antecedentes') or val_str.lower().startswith('contexto'):
                break 
                
            nueva_accion = detectar_accion_oficial(val_str)
            if nueva_accion:
                accion_actual = nueva_accion
                
            if not accion_actual:
                continue 
            
            for col in range(2, ws.max_column + 1):
                val_cambio = ws.cell(row=fila_cabeceras, column=col).value
                if not val_cambio or str(val_cambio).strip() == 'None': continue
                
                cambio_texto = str(val_cambio).replace('\n', ' ').strip()
                if re.match(r'^[A-Z]-[A-Z]+-\d+', cambio_texto): continue
                
                cell_conexion = ws.cell(row=row, column=col)
                val_conexion = cell_conexion.value
                
                if val_conexion and isinstance(val_conexion, str) and '-' in val_conexion:
                    codigos = [cd.strip() for cd in re.split(r'[\n,;\s]+', val_conexion) if cd.strip() and '-' in cd]
                    for codigo in codigos:
                        cod_limpio = limpiar_codigo_historia(codigo)
                        c_corto = extraer_corazon_codigo(cod_limpio)
                        estado_nom = obtener_estado_por_sufijo(cod_limpio)
                        color_fuente = colores_cambios_font.get(cambio_texto, "#BDBDBD")
                        
                        datos.append({
                            'Área': extraer_area_codigo(c_corto), 
                            'Iniciativa': iniciativa, 
                            'Acción Estratégica': accion_actual,
                            'Cambio Esperado': cambio_texto, 
                            'Historia_Cod': c_corto, 
                            'Historia_Corta': acortar_codigo(c_corto),
                            'Texto': textos_hist.get(c_corto, "Narrativa no encontrada."), 
                            'Estado': estado_nom,
                            'Color_Borde': color_fuente
                        })
    return datos

# LECTOR CENTRAL
datos_list = extraer_datos_puros("Matriz.xlsx")

# --- PANEL DE DIAGNÓSTICO ESTRUCTURAL ---
st.sidebar.markdown("### Diagnóstico")
st.sidebar.write("Registros:", len(datos_list))
st.sidebar.write("Historias:", len(set(d["Historia_Cod"] for d in datos_list)))
st.sidebar.write("Acciones:", len(set(d["Acción Estratégica"] for d in datos_list)))
st.sidebar.write("Cambios:", len(set(d["Cambio Esperado"] for d in datos_list)))
st.sidebar.write("Iniciativas:", len(set(d["Iniciativa"] for d in datos_list)))

acciones_unicas = sorted(list(set(d['Acción Estratégica'] for d in datos_list)))
cambios_unicos = sorted(list(set(d['Cambio Esperado'] for d in datos_list)))
dict_acciones = {acc: f"A{i+1}" for i, acc in enumerate(acciones_unicas)}
dict_cambios = {cam: f"C{i+1}" for i, cam in enumerate(cambios_unicos)}

COLORES_ESTADO = {'Negro (Ejemplo de Cambio)': '#212121', 'Naranja (Intención de Cambio)': '#FF9800', 'Rojo (Señal de Buen Camino)': '#D32F2F', 'Morado (Efecto Estancado)': '#351C75'}
COLORES_HEX_PUROS = {'Negro (Ejemplo de Cambio)': '#212121', 'Naranja (Intención de Cambio)': '#FF9800', 'Rojo (Señal de Buen Camino)': '#D32F2F', 'Morado (Efecto Estancado)': '#7B1FA2'}

# ==========================================
# PÁGINA 1: MAPA SISTÉMICO 
# ==========================================
if pagina_actual == "🕸️ Mapa Sistémico (Redes)":
    st.markdown("### 🎛️ Filtros del Mapa")
    col_f0, col_f1, col_f2, col_f3 = st.columns(4)
    with col_f0:
        area_sel = st.selectbox("🎯 1. Área Oportunidad:", ["Todas las Áreas"] + sorted(list(set(d['Área'] for d in datos_list))))
        datos_f0 = datos_list if area_sel == "Todas las Áreas" else [d for d in datos_list if d['Área'] == area_sel]
    with col_f1:
        ini_sel = st.selectbox("📌 2. Iniciativa:", ["Ver Todo (Macronivel)"] + sorted(list(set(d['Iniciativa'] for d in datos_f0))))
        datos_f1 = datos_f0 if ini_sel == "Ver Todo (Macronivel)" else [d for d in datos_f0 if d['Iniciativa'] == ini_sel]
    with col_f2:
        est_sel = st.selectbox("🎨 3. Estado:", ["Todos los estados"] + sorted(list(set(d['Estado'] for d in datos_f1))))
        datos_f2 = datos_f1 if est_sel == "Todos los estados" else [d for d in datos_f1 if d['Estado'] == est_sel]
    with col_f3:
        hist_sel = st.selectbox("🔍 4. Historia Específica:", ["Todas las historias"] + sorted(list(set(f"{d['Historia_Corta']} ({d['Historia_Cod']})" for d in datos_f2))))
        df_final = datos_f2 if hist_sel == "Todas las historias" else [d for d in datos_f2 if d['Historia_Cod'] == hist_sel.split("(")[1].replace(")", "")]

    es_macronivel = (ini_sel == "Ver Todo (Macronivel)")
    col_ctrl1, col_ctrl2 = st.columns(2)
    with col_ctrl1: vista_simplificada = st.toggle("👁️ **Vista Simplificada** (Ocultar historias)", value=es_macronivel)
    with col_ctrl2: congelar_mapa = st.toggle("🛑 **Congelar Mapa** (Fijar cajas)", value=es_macronivel)

    col_grafo, col_lector = st.columns([7, 3])
    with col_grafo:
        if not df_final:
            st.warning("No hay datos para esta combinación.")
        else:
            peso_acciones = Counter([d['Acción Estratégica'] for d in df_final])
            peso_cambios = Counter([d['Cambio Esperado'] for d in df_final])
            borde_cambios_map = {d['Cambio Esperado']: d['Color_Borde'] for d in df_final}
            
            min_peso = 1
            max_peso = max(list(peso_acciones.values()) + list(peso_cambios.values()) + [1])
            def calc_fs(peso):
                if max_peso == min_peso: return 16
                return int(16 + ((peso - min_peso) / (max_peso - min_peso)) * 24)

            # ESTRUCTURA CORREGIDA: cdn_resources='remote' evita creación de directorios 'lib'
            net = Network(height='700px', width='100%', directed=True, bgcolor='#FFFFFF', font_color='#202124', cdn_resources='remote')
            net.set_options(f"""
            var options = {{ "nodes": {{ "margin": 12, "borderWidth": 4, "borderWidthSelected": 6 }}, "edges": {{ "smooth": {{ "type": "dynamic" }}, "width": 2.5 }}, "layout": {{ "hierarchical": {{ "enabled": true, "direction": "LR", "levelSeparation": {600 if vista_simplificada else 380}, "nodeSpacing": 120 }} }}, "physics": {{ "enabled": {"false" if congelar_mapa else "true"}, "solver": "hierarchicalRepulsion" }} }}
            """)

            for acc in set(d['Acción Estratégica'] for d in df_final):
                fs = calc_fs(peso_acciones[acc])
                net.add_node(acc, label=formatear_caja(acc, 35), shape='box', level=1 if vista_simplificada else 2, color={'border': '#003366', 'background': '#E3F2FD'}, font={'color': '#003366', 'face': 'sans-serif', 'bold': True, 'size': fs})
            
            for cam in set(d['Cambio Esperado'] for d in df_final):
                fs = calc_fs(peso_cambios[cam])
                color_letra_excel = borde_cambios_map.get(cam, '#BDBDBD')
                net.add_node(cam, label=formatear_caja(cam, 35), shape='box', level=2 if vista_simplificada else 3, color={'background': '#FFFFFF', 'border': color_letra_excel}, font={'color': '#212121', 'face': 'sans-serif', 'bold': True, 'size': fs})

            if vista_simplificada:
                rutas = defaultdict(list)
                for d in df_final: rutas[(d['Acción Estratégica'], d['Cambio Esperado'], COLORES_HEX_PUROS.get(d['Estado'], '#000'), d['Estado'])].append(d['Historia_Corta'])
                track = Counter()
                for (acc, cam, color, estado), h_list in rutas.items():
                    track[(acc, cam)] += 1
                    idx = track[(acc, cam)]
                    curva = {"type": "straight"} if idx == 1 else {"type": "curvedCW", "roundness": 0.15 * (idx // 2) * (1 if idx % 2 == 0 else -1)}
                    net.add_edge(acc, cam, color=color, width=1.5 + (len(h_list)*1.5), title=f"Estado: {estado}\nHistorias: {', '.join(h_list)}", smooth=curva)
            else:
                historias_unicas = set(d['Historia_Cod'] for d in df_final)
                peso_historias = Counter([d['Historia_Cod'] for d in df_final])
                for hist in historias_unicas:
                    hist_corta = [d['Historia_Corta'] for d in df_final if d['Historia_Cod'] == hist][0]
                    fs_h = calc_fs(peso_historias[hist])
                    net.add_node(hist, label=hist_corta, color={'background': '#FFFFFF', 'border': '#E0E0E0'}, shape='ellipse', level=1, font={'size': fs_h})
                for d in df_final:
                    net.add_edge(d['Historia_Cod'], d['Acción Estratégica'], color='#E0E0E0')
                    net.add_edge(d['Acción Estratégica'], d['Cambio Esperado'], color=COLORES_HEX_PUROS.get(d['Estado'], '#000'), title=f"{d['Estado']}")

            # ESTRUCTURA CORREGIDA: Generación in-memory (No se escribe en disco)
            html_source = net.generate_html()
            components.html(html_source, height=720)

    with col_lector:
        st.header("📖 Lector Analítico")
        h_leer = st.selectbox("📚 Leer narrativa:", ["Seleccionar..."] + sorted(list(set(f"{d['Historia_Corta']} ({d['Historia_Cod']})" for d in df_final))))
        if h_leer != "Seleccionar...":
            cod_real = h_leer.split("(")[1].replace(")", "")
            info_h = [d for d in datos_list if d['Historia_Cod'] == cod_real]
            if info_h:
                st.success(f"**Historia:** {info_h[0]['Historia_Corta']}")
                st.write(info_h[0]['Texto'])
                st.markdown("---")
                for d in info_h: st.markdown(f"- ➔ {d['Cambio Esperado']} (**{d['Estado']}**)")

# ==========================================
# PÁGINA 2 Y 3: LÓGICA DE ANALÍTICA Y ECOSISTEMA
# ==========================================
elif pagina_actual in ["📊 Analítica de Portafolio", "🧬 Patrones de Co-Ocurrencia"]:
    st.markdown("### 🎛️ Filtros Globales")
    col_fg1, col_fg2 = st.columns(2)
    with col_fg1:
        area_sel = st.selectbox("🎯 1. Agrupar por Área de Oportunidad:", ["Todas las Áreas"] + sorted(list(set(d['Área'] for d in datos_list))))
        datos_f1 = datos_list if area_sel == "Todas las Áreas" else [d for d in datos_list if d['Área'] == area_sel]
    with col_fg2:
        ini_sel = st.selectbox("🌍 2. Analizar Iniciativa Específica:", ["Todo el Portafolio"] + sorted(list(set(d['Iniciativa'] for d in datos_f1))))
        datos_ana = datos_f1 if ini_sel == "Todo el Portafolio" else [d for d in datos_f1 if d['Iniciativa'] == ini_sel]
    
    st.markdown("---")
    
    if pagina_actual == "📊 Analítica de Portafolio":
        
        def crear_grafico_ranking(lista_datos, key_obj, color_scale, titulo_eje):
            lista_elementos = [d[key_obj] for d in lista_datos]
            if not lista_elementos: return None
            conteo = Counter(lista_elementos)
            items = sorted(conteo.items(), key=lambda x: x[1], reverse=False)
            y_labels = ["<br>".join(textwrap.wrap(k, width=50)) for k, v in items]
            x_vals = [v for k, v in items]
            textos_completos = [k for k, v in items]
            fig = go.Figure(go.Bar(x=x_vals, y=y_labels, orientation='h', marker=dict(color=x_vals, colorscale=color_scale), customdata=textos_completos, hovertemplate="<b>%{customdata}</b><br>Frecuencia: %{x}<extra></extra>"))
            fig.update_layout(xaxis_title=titulo_eje, yaxis_title="", margin=dict(l=0, r=0, t=0, b=0), height=max(300, len(items)*40))
            return fig

        st.markdown("### 🏆 Enfoque Estratégico (Rankings Cruzados)")
        c_gen1, c_gen2 = st.columns(2)
        with c_gen1:
            fig1 = crear_grafico_ranking(datos_ana, 'Acción Estratégica', 'Blues', 'Número de Historias')
            if fig1: st.plotly_chart(fig1, use_container_width=True)
            else: st.info("Sin datos.")
        with c_gen2:
            fig2 = crear_grafico_ranking(datos_ana, 'Cambio Esperado', 'Teal', 'Número de Historias')
            if fig2: st.plotly_chart(fig2, use_container_width=True)
            else: st.info("Sin datos.")

        st.markdown("#### Logros en Terreno (Ejemplos + Señales de Buen Camino)")
        datos_logros = [d for d in datos_ana if d['Estado'] in ['Negro (Ejemplo de Cambio)', 'Rojo (Señal de Buen Camino)']]
        c_log1, c_log2 = st.columns(2)
        with c_log1:
            fig3 = crear_grafico_ranking(datos_logros, 'Acción Estratégica', 'Greens', 'Historias (Negras/Rojas)')
            if fig3: st.plotly_chart(fig3, use_container_width=True)
        with c_log2:
            fig4 = crear_grafico_ranking(datos_logros, 'Cambio Esperado', 'Greens', 'Historias (Negras/Rojas)')
            if fig4: st.plotly_chart(fig4, use_container_width=True)

        st.markdown("#### Intenciones a Futuro (Solo Naranjas)")
        datos_intenciones = [d for d in datos_ana if d['Estado'] == 'Naranja (Intención de Cambio)']
        c_int1, c_int2 = st.columns(2)
        with c_int1:
            fig5 = crear_grafico_ranking(datos_intenciones, 'Acción Estratégica', 'Oranges', 'Historias (Naranjas)')
            if fig5: st.plotly_chart(fig5, use_container_width=True)
        with c_int2:
            fig6 = crear_grafico_ranking(datos_intenciones, 'Cambio Esperado', 'Oranges', 'Historias (Naranjas)')
            if fig6: st.plotly_chart(fig6, use_container_width=True)

        st.markdown("---")
        st.markdown("### 🌡️ Termómetros de Eficacia (Las 7 Tortas)")
        if datos_ana:
            conteo_gen = Counter([d['Estado'] for d in datos_ana])
            fig_gen = go.Figure(go.Pie(labels=list(conteo_gen.keys()), values=list(conteo_gen.values()), hole=0.4, marker=dict(colors=[COLORES_ESTADO.get(k, '#000') for k in conteo_gen.keys()])))
            fig_gen.update_layout(title_text="<b>PROMEDIO GENERAL DEL SISTEMA</b>", margin=dict(l=0, r=0, t=40, b=0), height=350)
            st.plotly_chart(fig_gen, use_container_width=True)

            cols_pie = st.columns(3)
            for i, accion in enumerate(acciones_unicas):
                datos_acc = [d['Estado'] for d in datos_ana if d['Acción Estratégica'] == accion]
                if datos_acc:
                    conteo_acc = Counter(datos_acc)
                    fig_acc = go.Figure(go.Pie(labels=list(conteo_acc.keys()), values=list(conteo_acc.values()), hole=0.5, marker=dict(colors=[COLORES_ESTADO.get(k, '#000') for k in conteo_acc.keys()])))
                    fig_acc.update_layout(title_text=f"<span style='font-size:13px'><b>{accion[:40]}...</b></span>", showlegend=False, margin=dict(l=10, r=10, t=40, b=10), height=220)
                    with cols_pie[i % 3]:
                        st.plotly_chart(fig_acc, use_container_width=True)

        st.markdown("---")
        st.markdown("### 🗺️ Densidad de Impacto: Acción vs Cambio")
        if datos_ana:
            df_heat_temp = pd.DataFrame(datos_ana)
            df_heat_temp['Accion_Corta'] = df_heat_temp['Acción Estratégica'].map(dict_acciones)
            df_heat_temp['Cambio_Corto'] = df_heat_temp['Cambio Esperado'].map(dict_cambios)
            
            heat_df = pd.crosstab(df_heat_temp['Accion_Corta'], df_heat_temp['Cambio_Corto'])
            hover_text = []
            for accion_corta in heat_df.index:
                hover_row = []
                for cambio_corto in heat_df.columns:
                    acc_real = [k for k, v in dict_acciones.items() if v == accion_corta][0]
                    cam_real = [k for k, v in dict_cambios.items() if v == cambio_corto][0]
                    cant = heat_df.loc[accion_corta, cambio_corto]
                    hover_row.append(f"<b>Acción:</b> {acc_real}<br><b>Cambio:</b> {cam_real}<br><b>Conexiones:</b> {cant}")
                hover_text.append(hover_row)

            fig_heat = px.imshow(heat_df, color_continuous_scale='YlGnBu', text_auto=True, aspect="auto")
            fig_heat.update_traces(customdata=hover_text, hovertemplate="%{customdata}<extra></extra>")
            fig_heat.update_layout(margin=dict(l=0, r=0, t=10, b=0), xaxis_title="Cambios Esperados (C1, C2...)", yaxis_title="Acciones Estratégicas (A1, A2...)")
            st.plotly_chart(fig_heat, use_container_width=True)
            
            with st.expander("📚 Ver Leyenda de Códigos (A1, C1...)"):
                col_l1, col_l2 = st.columns(2)
                with col_l1:
                    for k,v in dict_acciones.items(): st.markdown(f"**{v}:** {k}")
                with col_l2:
                    for k,v in dict_cambios.items(): st.markdown(f"**{v}:** {k}")

    # ==========================================
    # PÁGINA 3: PATRONES DE CO-OCURRENCIA 
    # ==========================================
    elif pagina_actual == "🧬 Patrones de Co-Ocurrencia":

        st.info("💡 **Fórmulas de Éxito:** El algoritmo cruza EXCLUSIVAMENTE las conexiones que tienen evidencia de cambio (Negras) o son señales de buen camino (Rojas). Omitiendo intenciones futuras o bloqueos.", icon="🚀")
        
        datos_exitosos = [d for d in datos_ana if d['Estado'] in ['Negro (Ejemplo de Cambio)', 'Rojo (Señal de Buen Camino)']]

        def algoritmo_seguro(lista_dicts, key_agrupadora, key_objetivo, nombre_grupo, min_r, max_r):
            grupos = defaultdict(set)
            for d in lista_dicts:
                if d[key_objetivo]: grupos[d[key_agrupadora]].add(d[key_objetivo])
                
            nombres_grupos = list(grupos.keys())
            patrones_encontrados = set()
            
            for i in range(len(nombres_grupos)):
                for j in range(i + 1, len(nombres_grupos)):
                    interseccion = frozenset(grupos[nombres_grupos[i]].intersection(grupos[nombres_grupos[j]]))
                    if len(interseccion) >= min_r:
                        if len(interseccion) <= max_r:
                            patrones_encontrados.add(interseccion)
                        if len(interseccion) <= 8:
                            for r in range(min_r, len(interseccion)):
                                for combo in itertools.combinations(interseccion, r):
                                    patrones_encontrados.add(frozenset(combo))

            resultados = {}
            for pat in patrones_encontrados:
                apariciones = [g for g, items in grupos.items() if pat.issubset(items)]
                if len(apariciones) > 1:
                    resultados[pat] = apariciones
                    
            patrones_limpios = {}
            for pat, aps in resultados.items():
                es_redundante = False
                for pat_otro, aps_otro in resultados.items():
                    if pat != pat_otro and pat.issubset(pat_otro) and len(aps) == len(aps_otro):
                        es_redundante = True
                        break
                if not es_redundante:
                    patrones_limpios[tuple(sorted(list(pat)))] = aps
                    
            ordenados = sorted(patrones_limpios.items(), key=lambda x: (len(x[1]), len(x[0])), reverse=True)
            
            if not ordenados:
                st.write(f"No hay patrones exitosos recurrentes en {nombre_grupo}s.")
                return
                
            for combo, grupos_list in ordenados[:30]:
                freq = len(grupos_list)
                tamano = len(combo)
                lista_md = "\n".join([f"* {item}" for item in combo])
                donde_se_vio = ", ".join(sorted(grupos_list))
                st.success(f"**[Fórmula de {tamano} Nodos] - Funcionó en {freq} {nombre_grupo}s distintos:**\n\n{lista_md}", icon="🎯")
                with st.expander(f"👁️ Ver en cuáles {nombre_grupo}s funcionó esto"):
                    st.write(f"**{nombre_grupo}s:** {donde_se_vio}")

        def algoritmo_mixto_veloz(lista_dicts, min_r=2, max_r=20):
            grupos = defaultdict(set)
            for d in lista_dicts:
                if d['Acción Estratégica']: grupos[d['Iniciativa']].add(f"🟦 {d['Acción Estratégica']}")
                if d['Cambio Esperado']: grupos[d['Iniciativa']].add(f"🟩 {d['Cambio Esperado']}")
                
            nombres_grupos = list(grupos.keys())
            patrones_encontrados = set()
            
            for i in range(len(nombres_grupos)):
                for j in range(i + 1, len(nombres_grupos)):
                    interseccion = frozenset(grupos[nombres_grupos[i]].intersection(grupos[nombres_grupos[j]]))
                    if len(interseccion) >= min_r and any(c.startswith("🟦") for c in interseccion) and any(c.startswith("🟩") for c in interseccion):
                        if len(interseccion) <= max_r:
                            patrones_encontrados.add(interseccion)
                        if len(interseccion) <= 8:
                            for r in range(min_r, len(interseccion)):
                                for combo in itertools.combinations(interseccion, r):
                                    if any(c.startswith("🟦") for c in combo) and any(c.startswith("🟩") for c in combo):
                                        patrones_encontrados.add(frozenset(combo))
                                        
            resultados = {}
            for pat in patrones_encontrados:
                apariciones = [g for g, items in grupos.items() if pat.issubset(items)]
                if len(apariciones) > 1:
                    resultados[pat] = apariciones
                    
            patrones_limpios = {}
            for pat, aps in resultados.items():
                es_redundante = False
                for pat_otro, aps_otro in resultados.items():
                    if pat != pat_otro and pat.issubset(pat_otro) and len(aps) == len(aps_otro):
                        es_redundante = True
                        break
                if not es_redundante:
                    patrones_limpios[tuple(sorted(list(pat)))] = aps
                    
            ordenados = sorted(patrones_limpios.items(), key=lambda x: (len(x[1]), len(x[0])), reverse=True)
            
            if not ordenados:
                st.write("No hay patrones mixtos exitosos que se repitan en múltiples Iniciativas.")
                return
                
            for combo, grupos_list in ordenados[:30]:
                freq = len(grupos_list)
                tamano = len(combo)
                lista_md = "\n".join([f"  {item}" for item in sorted(list(combo))])
                donde_se_vio = ", ".join(sorted(grupos_list))
                st.warning(f"**[Ecosistema de {tamano} Nodos] - Compartido con éxito por {freq} Iniciativas:**\n\n{lista_md}", icon="🌍")
                with st.expander("👁️ Ver cuáles iniciativas comparten este modelo"):
                    st.write(f"**Iniciativas:** {donde_se_vio}")

        tab_c1, tab_c2, tab_c3 = st.tabs(["📌 En una misma HISTORIA", "📌 En una misma INICIATIVA", "🌐 Ecosistema MIXTO"])

        with tab_c1:
            cc1, cc2 = st.columns(2)
            with cc1:
                st.markdown("**Acciones aplicadas simultáneamente con éxito:**")
                algoritmo_seguro(datos_exitosos, 'Historia_Cod', 'Acción Estratégica', "Historia", 2, 6)
            with cc2:
                st.markdown("**Efectos logrados simultáneamente:**")
                algoritmo_seguro(datos_exitosos, 'Historia_Cod', 'Cambio Esperado', "Historia", 2, 14)

        with tab_c2:
            cc3, cc4 = st.columns(2)
            with cc3:
                st.markdown("**Acciones exitosas movilizadas por la misma Iniciativa:**")
                algoritmo_seguro(datos_exitosos, 'Iniciativa', 'Acción Estratégica', "Iniciativa", 2, 6)
            with cc4:
                st.markdown("**Cambios logrados por la misma Iniciativa:**")
                algoritmo_seguro(datos_exitosos, 'Iniciativa', 'Cambio Esperado', "Iniciativa", 2, 14)

        with tab_c3:
            st.markdown("**Combinación Exitosa de Acciones + Cambios en la misma Iniciativa:**")
            algoritmo_mixto_veloz(datos_exitosos, 2, 20)

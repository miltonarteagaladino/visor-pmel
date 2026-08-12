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
        "🧬 Patrones de Co-Ocurrencia",
        "📈 Reporte Ejecutivo MEL"
    ])
    st.markdown("---")

titulo_limpio = pagina_actual.split(" ", 1)[1] if " " in pagina_actual else pagina_actual
st.markdown(f"""
    <div class="corona-header">
        <h1 class="corona-title">{titulo_limpio}</h1>
        <p class="corona-subtitle">ESTRATEGIA 2030 | Fundación Corona</p>
    </div>
""", unsafe_allow_html=True)

# --- 2. DECODIFICADORES GLOBALES ---
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

# --- 3. EXTRACCIÓN ESTRUCTURAL JERÁRQUICA (CATÁLOGO MAESTRO) ---
@st.cache_data
def extraer_universo_y_conexiones(ruta_archivo):
    wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
    
    # 3.1 Cargar Matriz Portafolio (Regla de oro de los 44)
    catalogo_iniciativas = []
    map_portafolio = {}
    map_nombre = {}
    
    if "Matriz Portafolio" in wb.sheetnames:
        ws_port = wb["Matriz Portafolio"]
        header_row, col_acro, col_port, col_nom = -1, -1, -1, -1
        
        for r in range(1, 10):
            for c in range(1, ws_port.max_column + 1):
                val = str(ws_port.cell(row=r, column=c).value).strip()
                if val == "Nombre": col_acro = c; header_row = r
                elif val == "Oportunidad a la que más aporta": col_port = c
                elif val == "Mecanismo de acción": col_nom = c
            if col_acro != -1: break
                
        if col_acro != -1 and col_port != -1:
            for r in range(header_row + 1, ws_port.max_row + 1):
                acronimo = str(ws_port.cell(row=r, column=col_acro).value).strip()
                porta = str(ws_port.cell(row=r, column=col_port).value).strip()
                nom = str(ws_port.cell(row=r, column=col_nom).value).strip() if col_nom != -1 else acronimo
                if acronimo and acronimo not in ['None', 'nan', '']:
                    porta_limpio = porta if porta and porta != 'nan' else 'Sin Portafolio'
                    catalogo_iniciativas.append({
                        'Acrónimo': acronimo,
                        'Iniciativa': nom,
                        'Portafolio': porta_limpio
                    })
                    map_portafolio[acronimo] = porta_limpio
                    map_nombre[acronimo] = nom

    def get_portafolio_y_nombre(sheet_name):
        if sheet_name in map_portafolio: return map_portafolio[sheet_name], map_nombre.get(sheet_name, sheet_name)
        if '-' in sheet_name:
            sin_prefijo = sheet_name.split('-', 1)[1]
            if sin_prefijo in map_portafolio: return map_portafolio[sin_prefijo], map_nombre.get(sin_prefijo, sheet_name)
        return "Sin Portafolio", sheet_name

    # 3.2 Extraer Conexiones
    datos = []
    pestañas = [sht for sht in wb.sheetnames if sht.startswith('I-') or sht.startswith('M-') or sht.startswith('B-') or sht.startswith('H-')]
    
    for sheet_name in pestañas:
        ws = wb[sheet_name]
        porta_asignado, iniciativa_asignada = get_portafolio_y_nombre(sheet_name)
        
        fila_cabeceras = -1
        for r in range(1, min(30, ws.max_row + 1)):
            val = ws.cell(row=r, column=1).value
            if val and str(val).strip().lower().startswith("acciones est"):
                fila_cabeceras = r
                break
        if fila_cabeceras == -1: continue
            
        colores_cambios_font = {}
        # CANDADO MATEMÁTICO: Columna 3 a 17 (Exactamente 15 cambios máximos permitidos)
        for col in range(3, min(18, ws.max_column + 1)):
            val_cambio = ws.cell(row=fila_cabeceras, column=col).value
            if val_cambio and str(val_cambio).strip() not in ['None', '']:
                # Normalización implacable de texto
                c_texto = re.sub(r'\s+', ' ', str(val_cambio).replace('\n', ' ').strip())
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
            if val_str.lower().startswith('antecedentes') or val_str.lower().startswith('contexto'): break 
                
            nueva_accion = detectar_accion_oficial(val_str)
            if nueva_accion: accion_actual = nueva_accion
            if not accion_actual: continue 
            
            # CANDADO MATEMÁTICO EN EXTRACCIÓN (Columnas 3 a 17)
            for col in range(3, min(18, ws.max_column + 1)):
                val_cambio = ws.cell(row=fila_cabeceras, column=col).value
                if not val_cambio or str(val_cambio).strip() in ['None', '']: continue
                
                cambio_texto = re.sub(r'\s+', ' ', str(val_cambio).replace('\n', ' ').strip())
                
                val_conexion = ws.cell(row=row, column=col).value
                if val_conexion and isinstance(val_conexion, str) and '-' in val_conexion:
                    codigos = [cd.strip() for cd in re.split(r'[\n,;\s]+', val_conexion) if cd.strip() and '-' in cd]
                    for codigo in codigos:
                        cod_limpio = limpiar_codigo_historia(codigo)
                        c_corto = extraer_corazon_codigo(cod_limpio)
                        estado_nom = obtener_estado_por_sufijo(cod_limpio)
                        color_fuente = colores_cambios_font.get(cambio_texto, "#BDBDBD")
                        
                        datos.append({
                            'Portafolio': porta_asignado,
                            'Iniciativa': iniciativa_asignada, 
                            'Acción Estratégica': accion_actual,
                            'Cambio Esperado': cambio_texto, 
                            'Historia_Cod': c_corto, 
                            'Historia_Corta': acortar_codigo(c_corto),
                            'Texto': textos_hist.get(c_corto, "Narrativa no encontrada."), 
                            'Estado': estado_nom,
                            'Color_Borde': color_fuente,
                            'Área': extraer_area_codigo(c_corto)
                        })
                        
    return catalogo_iniciativas, datos

catalogo_iniciativas, datos_list = extraer_universo_y_conexiones("Matriz.xlsx")
df_catalogo = pd.DataFrame(catalogo_iniciativas)
df_conexiones = pd.DataFrame(datos_list)

# --- DIAGNÓSTICO EN BARRA LATERAL ---
with st.sidebar:
    st.markdown("### 🛠️ Diagnóstico General")
    st.write(f"**Universo Iniciativas:** {len(df_catalogo)}")
    st.write(f"**Conexiones:** {len(datos_list)}")
    st.write(f"**Historias:** {len(set(d['Historia_Cod'] for d in datos_list))}")
    st.write(f"**Acciones:** {len(set(d['Acción Estratégica'] for d in datos_list))}")
    st.write(f"**Cambios:** {len(set(d['Cambio Esperado'] for d in datos_list))}")
    st.markdown("---")

acciones_unicas = sorted(list(set(d['Acción Estratégica'] for d in datos_list)))
cambios_unicos = sorted(list(set(d['Cambio Esperado'] for d in datos_list)))
dict_acciones = {acc: f"A{i+1}" for i, acc in enumerate(acciones_unicas)}
dict_cambios = {cam: f"C{i+1}" for i, cam in enumerate(cambios_unicos)}

# COLORES GENERALES (Usados en todos lados excepto en MEL)
COLORES_ESTADO = {'Negro (Ejemplo de Cambio)': '#212121', 'Naranja (Intención de Cambio)': '#FF9800', 'Rojo (Señal de Buen Camino)': '#D32F2F', 'Morado (Efecto Estancado)': '#351C75'}
COLORES_HEX_PUROS = {'Negro (Ejemplo de Cambio)': '#212121', 'Naranja (Intención de Cambio)': '#FF9800', 'Rojo (Señal de Buen Camino)': '#D32F2F', 'Morado (Efecto Estancado)': '#7B1FA2'}

# COLORES ESPECÍFICOS PARA REPORTE MEL (Ajuste 1)
MEL_COLORS = {
    'Negro (Ejemplo de Cambio)': '#4CAF50', # Verde
    'Rojo (Señal de Buen Camino)': '#FF9800', # Naranja
    'Naranja (Intención de Cambio)': '#2196F3', # Azul
    'Morado (Efecto Estancado)': '#F44336' # Rojo
}

# ==========================================
# PÁGINA 1: MAPA SISTÉMICO 
# ==========================================
if pagina_actual == "🕸️ Mapa Sistémico (Redes)":
    st.markdown("### 🎛️ Filtros del Mapa (Clásico)")
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

            html_source = net.generate_html()
            components.html(html_source, height=720)

    with col_lector:
        st.header("📖 Lector Narrativo")
        h_leer = st.selectbox("📚 Seleccionar historia:", ["Seleccionar..."] + sorted(list(set(f"{d['Historia_Corta']} ({d['Historia_Cod']})" for d in df_final))))
        if h_leer != "Seleccionar...":
            cod_real = h_leer.split("(")[1].replace(")", "")
            info_h = [d for d in datos_list if d['Historia_Cod'] == cod_real]
            if info_h:
                st.success(f"**Historia:** {info_h[0]['Historia_Corta']}")
                st.write(info_h[0]['Texto'])
                st.markdown("---")
                for d in info_h: st.markdown(f"- ➔ {d['Cambio Esperado']} (**{d['Estado']}**)")

# ==========================================
# PÁGINA 2 Y 3: LÓGICA DE ANALÍTICA EXISTENTE
# ==========================================
elif pagina_actual in ["📊 Analítica de Portafolio", "🧬 Patrones de Co-Ocurrencia"]:
    st.markdown("### 🎛️ Filtros Globales (Jerarquía Estricta)")
    col_fg1, col_fg2 = st.columns(2)
    with col_fg1:
        portafolios_disp = sorted([p for p in df_catalogo['Portafolio'].unique() if p != ''])
        port_sel = st.selectbox("🎯 1. Portafolio:", ["Todos los Portafolios"] + portafolios_disp)
        cat_f1 = df_catalogo if port_sel == "Todos los Portafolios" else df_catalogo[df_catalogo['Portafolio'] == port_sel]
        datos_f1 = datos_list if port_sel == "Todos los Portafolios" else [d for d in datos_list if d.get('Portafolio') == port_sel]
    with col_fg2:
        inis_disp = sorted(cat_f1['Iniciativa'].unique())
        ini_sel = st.selectbox("🌍 2. Iniciativa:", ["Todo el Portafolio"] + inis_disp)
        cat_ana = cat_f1 if ini_sel == "Todo el Portafolio" else cat_f1[cat_f1['Iniciativa'] == ini_sel]
        datos_ana = datos_f1 if ini_sel == "Todo el Portafolio" else [d for d in datos_f1 if d.get('Iniciativa') == ini_sel]
    st.markdown("---")
    
    if pagina_actual == "📊 Analítica de Portafolio":
        
        st.markdown("#### 📊 Métricas de Validación del Portafolio")
        num_ini = len(cat_ana)
        num_hist = len(set(d['Historia_Cod'] for d in datos_ana))
        num_acc = len(set(d['Acción Estratégica'] for d in datos_ana))
        num_cam = len(set(d['Cambio Esperado'] for d in datos_ana))
        
        km1, km2, km3, km4 = st.columns(4)
        km1.metric("Iniciativas del Universo", num_ini)
        km2.metric("Historias Trazadas", num_hist)
        km3.metric("Nodos Únicos", f"{num_acc + num_cam}", f"{num_acc} Acciones | {num_cam} Cambios", delta_color="off")
        km4.metric("Conexiones Totales", len(datos_ana))
        st.markdown("---")

        if not datos_ana:
            st.warning("La iniciativa o portafolio seleccionado tiene 0 historias y 0 conexiones reportadas en la matriz.")
        else:
            def crear_grafico_ranking(lista_datos, key_obj, color_scale, titulo_eje):
                lista_elementos = [d[key_obj] for d in lista_datos]
                if not lista_elementos: return None
                conteo = Counter(lista_elementos)
                items = sorted(conteo.items(), key=lambda x: x[1], reverse=False)
                y_labels = ["<br>".join(textwrap.wrap(k, width=50)) for k, v in items]
                x_vals = [v for k, v in items]
                textos_completos = [k for k, v in items]
                fig = go.Figure(go.Bar(x=x_vals, y=y_labels, orientation='h', marker=dict(color=x_vals, colorscale=color_scale), customdata=textos_completos, hovertemplate="<b>%{customdata}</b><br>Frecuencia: %{x}<extra></extra>"))
                fig.update_layout(xaxis_title=titulo_eje, yaxis_title="", margin=dict(l=0, r=0, t=0, b=0), height=max(300, len(items)*40), template="plotly_white")
                return fig

            st.markdown("### 🏆 Enfoque Estratégico (Rankings Cruzados)")
            c_gen1, c_gen2 = st.columns(2)
            with c_gen1:
                fig1 = crear_grafico_ranking(datos_ana, 'Acción Estratégica', 'Blues', 'Número de Historias')
                if fig1: st.plotly_chart(fig1, use_container_width=True)
            with c_gen2:
                fig2 = crear_grafico_ranking(datos_ana, 'Cambio Esperado', 'Teal', 'Número de Historias')
                if fig2: st.plotly_chart(fig2, use_container_width=True)

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
            conteo_gen = Counter([d['Estado'] for d in datos_ana])
            fig_gen = go.Figure(go.Pie(labels=list(conteo_gen.keys()), values=list(conteo_gen.values()), hole=0.4, marker=dict(colors=[COLORES_ESTADO.get(k, '#000') for k in conteo_gen.keys()])))
            fig_gen.update_layout(title_text="<b>PROMEDIO GENERAL DEL SISTEMA</b>", margin=dict(l=0, r=0, t=40, b=0), height=350, template="plotly_white")
            st.plotly_chart(fig_gen, use_container_width=True)

            cols_pie = st.columns(3)
            for i, accion in enumerate(acciones_unicas):
                datos_acc = [d['Estado'] for d in datos_ana if d['Acción Estratégica'] == accion]
                if datos_acc:
                    conteo_acc = Counter(datos_acc)
                    fig_acc = go.Figure(go.Pie(labels=list(conteo_acc.keys()), values=list(conteo_acc.values()), hole=0.5, marker=dict(colors=[COLORES_ESTADO.get(k, '#000') for k in conteo_acc.keys()])))
                    fig_acc.update_layout(title_text=f"<span style='font-size:13px'><b>{accion[:40]}...</b></span>", showlegend=False, margin=dict(l=10, r=10, t=40, b=10), height=220, template="plotly_white")
                    with cols_pie[i % 3]:
                        st.plotly_chart(fig_acc, use_container_width=True)

            st.markdown("---")
            st.markdown("### 🗺️ Densidad de Impacto: Acción vs Cambio")
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
            fig_heat.update_layout(margin=dict(l=0, r=0, t=10, b=0), xaxis_title="Cambios Esperados (C1, C2...)", yaxis_title="Acciones Estratégicas (A1, A2...)", template="plotly_white")
            st.plotly_chart(fig_heat, use_container_width=True)

    # ==========================================
    # PÁGINA 3: PATRONES DE CO-OCURRENCIA 
    # ==========================================
    elif pagina_actual == "🧬 Patrones de Co-Ocurrencia":
        st.info("💡 **Fórmulas de Éxito:** El algoritmo cruza EXCLUSIVAMENTE las conexiones que tienen evidencia de cambio (Negras) o son señales de buen camino (Rojas).", icon="🚀")
        datos_exitosos = [d for d in datos_ana if d['Estado'] in ['Negro (Ejemplo de Cambio)', 'Rojo (Señal de Buen Camino)']]

        if not datos_exitosos:
            st.warning("El portafolio seleccionado no tiene conexiones exitosas.")
        else:
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
                            if len(interseccion) <= max_r: patrones_encontrados.add(interseccion)
                            if len(interseccion) <= 8:
                                for r in range(min_r, len(interseccion)):
                                    for combo in itertools.combinations(interseccion, r):
                                        patrones_encontrados.add(frozenset(combo))

                resultados = {}
                for pat in patrones_encontrados:
                    apariciones = [g for g, items in grupos.items() if pat.issubset(items)]
                    if len(apariciones) > 1: resultados[pat] = apariciones
                        
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
                            if len(interseccion) <= max_r: patrones_encontrados.add(interseccion)
                            if len(interseccion) <= 8:
                                for r in range(min_r, len(interseccion)):
                                    for combo in itertools.combinations(interseccion, r):
                                        if any(c.startswith("🟦") for c in combo) and any(c.startswith("🟩") for c in combo):
                                            patrones_encontrados.add(frozenset(combo))
                                            
                resultados = {}
                for pat in patrones_encontrados:
                    apariciones = [g for g, items in grupos.items() if pat.issubset(items)]
                    if len(apariciones) > 1: resultados[pat] = apariciones
                        
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

# ==========================================
# PÁGINA 4: REPORTE EJECUTIVO MEL
# ==========================================
elif pagina_actual == "📈 Reporte Ejecutivo MEL":
    st.markdown("### 📈 Reporte Ejecutivo MEL")
    st.caption("Visión macroestructural del Portafolio Institucional (Gráficos optimizados para exportación a PowerPoint)")
    
    df_conn = pd.DataFrame(datos_list)
    df_mel_resumen = df_catalogo.copy()
    
    if df_conn.empty: df_conn = pd.DataFrame(columns=['Iniciativa', 'Portafolio', 'Historia_Cod', 'Acción Estratégica', 'Cambio Esperado', 'Estado'])
        
    hists_por_ini = df_conn.groupby('Iniciativa')['Historia_Cod'].nunique().reset_index(name='Historias')
    conns_por_ini = df_conn.groupby('Iniciativa').size().reset_index(name='Conexiones')
    
    df_mel_resumen = df_mel_resumen.merge(hists_por_ini, on='Iniciativa', how='left')
    df_mel_resumen = df_mel_resumen.merge(conns_por_ini, on='Iniciativa', how='left')
    df_mel_resumen['Historias'] = df_mel_resumen['Historias'].fillna(0).astype(int)
    df_mel_resumen['Conexiones'] = df_mel_resumen['Conexiones'].fillna(0).astype(int)
    
    # ---------------- ANÁLISIS 1: DISTRIBUCIÓN ----------------
    st.markdown("#### 1. Distribución del Portafolio")
    tot_inis = len(df_mel_resumen)
    tot_hists = df_mel_resumen['Historias'].sum()
    
    dist_port = df_mel_resumen.groupby('Portafolio').agg(
        Iniciativas=('Iniciativa', 'nunique'),
        Historias=('Historias', 'sum'),
        Conexiones=('Conexiones', 'sum')
    ).reset_index()
    
    dist_port['% Iniciativas'] = ((dist_port['Iniciativas'] / max(1, tot_inis)) * 100).round(1).astype(str) + '%'
    dist_port['% Historias'] = ((dist_port['Historias'] / max(1, tot_hists)) * 100).round(1).astype(str) + '%'
    dist_port = dist_port[['Portafolio', 'Iniciativas', '% Iniciativas', 'Historias', '% Historias', 'Conexiones']]
    st.dataframe(dist_port, use_container_width=True, hide_index=True)
    st.markdown("---")

    # ---------------- ANÁLISIS 2: MADUREZ ----------------
    st.markdown("#### 2. Madurez de Cambio por Portafolio")
    if not df_conn.empty:
        madurez = df_conn.groupby(['Portafolio', 'Estado']).size().reset_index(name='Conexiones')
        totales = madurez.groupby('Portafolio')['Conexiones'].transform('sum')
        madurez['Porcentaje'] = (madurez['Conexiones'] / totales * 100).round(1)
        
        fig_madurez = px.bar(madurez, x='Portafolio', y='Conexiones', color='Estado', 
                             color_discrete_map=MEL_COLORS, text='Porcentaje',
                             title="Distribución de Estados de Conexión (Volumen y Porcentaje)")
        fig_madurez.update_traces(texttemplate='%{text}%')
        fig_madurez.update_layout(template="plotly_white", barmode='stack', legend_title="Estado de Madurez")
        st.plotly_chart(fig_madurez, use_container_width=True)
    else: st.info("No hay datos suficientes.")
    st.markdown("---")

    # ---------------- ANÁLISIS 3: INICIATIVAS LÍDERES ----------------
    st.markdown("#### 3. Iniciativas Líderes (Índice de Madurez)")
    if not df_conn.empty:
        df_conn['Es_Maduro'] = df_conn['Estado'].isin(['Negro (Ejemplo de Cambio)', 'Rojo (Señal de Buen Camino)'])
        df_conn['Es_Intencion'] = df_conn['Estado'] == 'Naranja (Intención de Cambio)'
        df_conn['Es_Reto'] = df_conn['Estado'] == 'Morado (Efecto Estancado)'
        
        # Agrupamos sumando y cruzamos con el catálogo completo para incluir los 0s
        madurez_ini = df_conn.groupby('Iniciativa')[['Es_Maduro', 'Es_Intencion', 'Es_Reto']].sum().reset_index()
        madurez_ini = pd.merge(df_catalogo[['Iniciativa']], madurez_ini, on='Iniciativa', how='left').fillna(0)
        
        c_l1, c_l2, c_l3 = st.columns(3)
        with c_l1:
            st.write("**TOP INICIATIVAS (Negros + Rojos)**")
            top1 = madurez_ini.sort_values('Es_Maduro', ascending=False).head(10)
            st.dataframe(top1[['Iniciativa', 'Es_Maduro']].rename(columns={'Es_Maduro': 'Logros'}), hide_index=True, use_container_width=True)
            fig1 = px.bar(top1, x='Es_Maduro', y='Iniciativa', orientation='h', color_discrete_sequence=['#4CAF50'], title="Mayor Evidencia de Cambio")
            fig1.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=0, r=0, t=30, b=0), height=300, template="plotly_white")
            st.plotly_chart(fig1, use_container_width=True)
            
        with c_l2:
            st.write("**MÁS INTENCIONES (Naranjas)**")
            top2 = madurez_ini.sort_values('Es_Intencion', ascending=False).head(10)
            st.dataframe(top2[['Iniciativa', 'Es_Intencion']].rename(columns={'Es_Intencion': 'Intenciones'}), hide_index=True, use_container_width=True)
            fig2 = px.bar(top2, x='Es_Intencion', y='Iniciativa', orientation='h', color_discrete_sequence=['#2196F3'], title="Fase de Intención")
            fig2.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=0, r=0, t=30, b=0), height=300, template="plotly_white")
            st.plotly_chart(fig2, use_container_width=True)
            
        with c_l3:
            st.write("**MÁS RETOS (Morados)**")
            top3 = madurez_ini.sort_values('Es_Reto', ascending=False).head(10)
            st.dataframe(top3[['Iniciativa', 'Es_Reto']].rename(columns={'Es_Reto': 'Retos'}), hide_index=True, use_container_width=True)
            fig3 = px.bar(top3, x='Es_Reto', y='Iniciativa', orientation='h', color_discrete_sequence=['#F44336'], title="Fase de Estancamiento")
            fig3.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(l=0, r=0, t=30, b=0), height=300, template="plotly_white")
            st.plotly_chart(fig3, use_container_width=True)
    st.markdown("---")

    # ---------------- ANÁLISIS 4: ROLES ESTRATÉGICOS ----------------
    st.markdown("#### 4. Roles Estratégicos (Acciones) por Portafolio")
    if not df_conn.empty:
        roles_port = pd.crosstab(df_conn['Portafolio'], df_conn['Acción Estratégica'], normalize='index') * 100
        st.write("**Tabla de Distribución Porcentual**")
        st.dataframe(roles_port.round(1).astype(str) + '%', use_container_width=True)
        
        roles_melt = roles_port.reset_index().melt(id_vars='Portafolio', var_name='Acción Estratégica', value_name='Porcentaje')
        fig_roles = px.bar(roles_melt, x='Portafolio', y='Porcentaje', color='Acción Estratégica',
                           title="Rol Predominante por Portafolio (100%)",
                           color_discrete_sequence=px.colors.qualitative.Prism, text_auto='.1f')
        fig_roles.update_layout(template="plotly_white", barmode='stack')
        st.plotly_chart(fig_roles, use_container_width=True)
    st.markdown("---")

    # ---------------- ANÁLISIS 5: CAMBIOS MOVILIZADOS ----------------
    st.markdown("#### 5. Cambios Estratégicos Movilizados")
    if not df_conn.empty:
        cam_port_nunique = df_conn.groupby('Portafolio')['Cambio Esperado'].nunique().reset_index(name='Diversidad de Cambios (Max 15)')
        st.write("**Diversidad de Cambios**")
        st.dataframe(cam_port_nunique, hide_index=True)
        
        st.write("**Mapa de Calor: Concentración de Cambios por Portafolio**")
        df_conn['Cambio_Corto'] = df_conn['Cambio Esperado'].map(dict_cambios)
        heat_df = pd.crosstab(df_conn['Portafolio'], df_conn['Cambio_Corto'])
        
        fig_heat_mel = px.imshow(heat_df, color_continuous_scale='Blues', text_auto=True, aspect="auto")
        fig_heat_mel.update_layout(template="plotly_white", xaxis_title="Cambios Esperados (C1 a C15)", yaxis_title="Portafolio")
        st.plotly_chart(fig_heat_mel, use_container_width=True)
    st.markdown("---")

    # ---------------- ANÁLISIS 6: ALINEACIÓN ESTRATÉGICA ----------------
    st.markdown("#### 6. Alineación Estratégica por Portafolio")
    st.caption("Una historia es desalineada si apunta a 'Cambio - no en estrategia'. Una iniciativa es desalineada si tiene al menos 1 historia desalineada.")
    if not df_conn.empty:
        df_conn['No_Alineada'] = df_conn['Cambio Esperado'] == 'Cambio - no en estrategia'
        
        # Merge conn alignment back to full catalog (0 connections = Aligned)
        ini_alin = df_conn.groupby('Iniciativa')['No_Alineada'].any().reset_index()
        df_aliniacion = pd.merge(df_catalogo, ini_alin, on='Iniciativa', how='left').fillna(False)
        
        for p in df_catalogo['Portafolio'].unique():
            if p == '': continue
            st.markdown(f"**Portafolio: {p}**")
            cp1, cp2 = st.columns(2)
            
            # Filtros por portafolio
            df_p_ini = df_aliniacion[df_aliniacion['Portafolio'] == p]
            df_p_hist = df_conn[df_conn['Portafolio'] == p]
            
            with cp1:
                if not df_p_ini.empty:
                    alineadas = (~df_p_ini['No_Alineada']).sum()
                    no_alineadas = df_p_ini['No_Alineada'].sum()
                    fig_a1 = px.pie(names=['Alineadas', 'No Alineadas'], values=[alineadas, no_alineadas],
                                    title="Iniciativas", color_discrete_sequence=['#4CAF50', '#D32F2F'], hole=0.4)
                    fig_a1.update_layout(template="plotly_white", height=300, margin=dict(l=0, r=0, t=30, b=0))
                    st.plotly_chart(fig_a1, use_container_width=True)
                else: st.write("Sin iniciativas")
                
            with cp2:
                if not df_p_hist.empty:
                    h_alin_df = df_p_hist.groupby('Historia_Cod')['No_Alineada'].any()
                    h_alin_count = (~h_alin_df).sum()
                    h_no_alin_count = h_alin_df.sum()
                    fig_a2 = px.pie(names=['Alineadas', 'No Alineadas'], values=[h_alin_count, h_no_alin_count],
                                    title="Historias", color_discrete_sequence=['#4CAF50', '#D32F2F'], hole=0.4)
                    fig_a2.update_layout(template="plotly_white", height=300, margin=dict(l=0, r=0, t=30, b=0))
                    st.plotly_chart(fig_a2, use_container_width=True)
                else: st.write("Sin historias")
            st.markdown("<br>", unsafe_allow_html=True)

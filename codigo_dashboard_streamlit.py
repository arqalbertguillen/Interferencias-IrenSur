import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime, timedelta, timezone
import io
import streamlit.components.v1 as components
import glob
import os

st.set_page_config(page_title="Dashboard BIM - IREN SUR", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stApp { background-color: #1b202b; color: #e2e8f0; }
    .header-container { background-color: #242b38; padding: 20px 30px; border-radius: 8px; border: 1px solid #f7931e; margin-bottom: 25px; }
    .header-top-text { color: #f7931e; font-size: 12px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 5px; }
    .header-title { color: #ffffff; font-size: 32px; font-weight: 800; margin-bottom: 0px; display: flex; align-items: center; }
    .badge { background-color: #2d88ff; color: white; font-size: 14px; padding: 4px 10px; border-radius: 6px; margin-left: 15px; font-weight: bold; }
    .header-subtitle { color: #94a3b8; font-size: 15px; margin-top: 5px; margin-bottom: 20px; }
    .meta-grid { display: flex; justify-content: space-between; border-top: 1px solid #334155; padding-top: 15px; }
    .meta-box { display: flex; flex-direction: column; }
    .meta-title { color: #64748b; font-size: 11px; font-weight: bold; margin-bottom: 3px; text-transform: uppercase; }
    .meta-value { color: #f8fafc; font-size: 14px; font-weight: 600; }
    .chart-container { background-color: #242b38; padding: 20px; border-radius: 10px; border: 1px solid #334155; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; height: 100%; }
    h1, h2, h3, h4 { color: #f7931e !important; }
    div[data-testid="stMetricValue"] { color: #ffffff; }
    div[data-testid="stMetricLabel"] { color: #94a3b8; }
    
    @media print { 
        header, footer { display: none !important; } 
        section[data-testid="stSidebar"] { display: none !important; }
        div[data-testid="stFileUploader"] { display: none !important; }
        div[data-testid="stAlert"] { display: none !important; }
        button { display: none !important; }
        div[data-testid="stMultiSelect"] { display: none !important; }
        div[role="tablist"] { display: none !important; } 
        .hide-print { display: none !important; } 
        .stApp { background-color: white !important; } 
        * { color: black !important; } 
        .header-container { background-color: white !important; border: 2px solid black !important; }
        .header-title, .header-subtitle, .meta-title, .meta-value, .header-top-text { color: black !important; }
        
        /* NUEVAS REGLAS: Fuerzan a la web a expandirse completa hacia abajo para el PDF */
        html, body, .stApp, .main, .block-container, div[data-testid="stVerticalBlock"], div[data-testid="stTabs"], div[role="tabpanel"] {
            height: auto !important;
            overflow: visible !important;
            display: block !important;
            position: relative !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] { 
            border: 1px solid #ccc !important; background-color: white !important; 
            height: auto !important; overflow: visible !important; display: block !important;
        }
        table { page-break-inside: auto !important; }
        tr { page-break-inside: avoid !important; page-break-after: auto !important; }
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data(file):
    df_raw = pd.read_excel(file, header=None)
    current_test_name = "Grupo Desconocido"
    current_tol = "-"
    data_rows, headers = [], []
    expecting_tol_values = False
    tol_idx = -1
    
    for index, row in df_raw.iterrows():
        row_strs = [str(x).strip() for x in row]
        if expecting_tol_values:
            if tol_idx != -1 and tol_idx < len(row_strs):
                val = row_strs[tol_idx]
                if val and val != 'nan': current_tol = val
            expecting_tol_values = False
            
        if 'Tolerancia' in row_strs:
            posibles = [x for x in row_strs if x not in ['nan', '', 'Tolerancia']]
            if posibles: current_test_name = posibles[0]
            if 'Tolerancia' in row_strs:
                tol_idx = row_strs.index('Tolerancia')
                expecting_tol_values = True
            
        elif 'Nombre de conflicto' in row_strs or 'Estado' in row_strs:
            if not headers:
                raw_h = [x if x not in ['nan', ''] else f"Col_{i}" for i, x in enumerate(row_strs)]
                seen = set()
                for h in raw_h:
                    val = h
                    c = 1
                    while val in seen:
                        val = f"{h}_{c}"
                        c += 1
                    seen.add(val)
                    headers.append(val)
                headers.append('Test') 
                headers.append('Tolerancia_Grupo') 
        elif any(estado in row_strs for estado in ['Activo', 'Resuelto', 'Aprobado', 'Nuevo', 'Revisado']):
            if headers: 
                row_list = list(row)
                row_list.append(current_test_name)
                row_list.append(current_tol)
                data_rows.append(row_list)
                
    if data_rows and headers:
        clean_rows = [r + [None]*(len(headers)-len(r)) if len(r) < len(headers) else r[:len(headers)] for r in data_rows]
        df = pd.DataFrame(clean_rows, columns=headers)
    else:
        df = pd.read_excel(file, header=7)
        if 'Test' not in df.columns: df['Test'] = 'Todos los conflictos'
            
    df.columns = df.columns.str.replace('Personalizar ', '', case=False).str.strip()
    mask = df.astype(str).apply(lambda x: x.str.contains('Armadura estructural', case=False, na=False)).any(axis=1)
    return df[~mask]

with st.sidebar:
    st.markdown("### Exportar Dashboard")
    components.html("""
        <button onclick="window.parent.print()" style="background-color:#f7931e; color:white; padding:10px 15px; border:none; border-radius:5px; cursor:pointer; width:100%; font-weight:bold;">
            🖨️ Guardar como PDF
        </button>
    """, height=50)

uploaded_files = st.file_uploader("📂 [Revisión Interna] Arrastra Excels aquí para previsualizar. Si se deja vacío, cargarán los Datos de la Nube.", type=['xlsx'], accept_multiple_files=True)

archivos_a_procesar = []
origen_datos = ""

if uploaded_files:
    archivos_a_procesar = uploaded_files
    origen_datos = "Archivos Locales (Previsualización)"
else:
    archivos_a_procesar = sorted(glob.glob("*.xlsx"))
    origen_datos = "Datos de la Nube"

if archivos_a_procesar:
    num_archivos = len(archivos_a_procesar)
    version_str = f"V{num_archivos:02d}"

    all_dfs = []
    for f in archivos_a_procesar:
        df_part = load_data(f)
        all_dfs.append(df_part)
    df_timeline = pd.concat(all_dfs, ignore_index=True)
    
    ultimo_archivo = archivos_a_procesar[-1]
    df_dashboard = load_data(ultimo_archivo)

    zona_peru = timezone(timedelta(hours=-5))
    fecha_hoy = datetime.now(zona_peru).strftime("%d/%m/%Y")
    
    st.markdown(f"""
    <div class="header-container">
        <div class="header-top-text">GESTIÓN BIM • CONTROL DE DISEÑO</div>
        <div class="header-title">Hospital IREN SUR Permanente <span class="badge">{version_str}</span></div>
        <div class="header-subtitle">Dashboard de Control Gerencial — Seguimiento de Interferencias</div>
        <div class="meta-grid">
            <div class="meta-box"><span class="meta-title">ELABORADO POR</span><span class="meta-value">Coord. BIM Diseño • Albert Guillen</span></div>
            <div class="meta-box"><span class="meta-title">VALIDADO POR</span><span class="meta-value">KP-07 • Jhonathan Seminario</span></div>
            <div class="meta-box"><span class="meta-title">ORIGEN DE DATOS</span><span class="meta-value" style="color:#4ade80;">{origen_datos}</span></div>
            <div class="meta-box"><span class="meta-title">FECHA DE CORRIDA</span><span class="meta-value">{fecha_hoy}</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.info("💡 **Para exportar a PDF:** Presiona `Ctrl + P` en tu teclado. El reporte se auto-formateará de color blanco para ocultar menús e imprimir limpio.")

    if not df_dashboard.empty:
        estado_col = next((col for col in df_dashboard.columns if 'estado' in col.lower()), None)
        
        if estado_col:
            df_dashboard[estado_col] = df_dashboard[estado_col].astype(str).str.strip()
            if estado_col in df_timeline.columns:
                df_timeline[estado_col] = df_timeline[estado_col].astype(str).str.strip()
            
            st.markdown('<h3 class="hide-print">🔍 Filtros del Proyecto</h3>', unsafe_allow_html=True)
            
            c_ed = next((col for col in df_dashboard.columns if 'edificio' in str(col).lower()), None)
            c_ni = next((col for col in df_dashboard.columns if 'nivel' in str(col).lower()), None)
            c_zo = next((col for col in df_dashboard.columns if 'zona' in str(col).lower()), None)
            c_am = next((col for col in df_dashboard.columns if 'ambiente' in str(col).lower()), None)
            
            f1, f2, f3 = st.columns(3)
            f4, f5 = st.columns([2, 1]) 
            
            df_dash_filt = df_dashboard.copy()
            df_time_filt = df_timeline.copy()
            
            if c_ed:
                sel_ed = f1.multiselect("Edificio", sorted([str(x) for x in df_dash_filt[c_ed].dropna().unique() if str(x)!='nan']))
                if sel_ed: 
                    df_dash_filt = df_dash_filt[df_dash_filt[c_ed].astype(str).isin(sel_ed)]
                    df_time_filt = df_time_filt[df_time_filt[c_ed].astype(str).isin(sel_ed)]
            if c_ni:
                sel_ni = f2.multiselect("Nivel", sorted([str(x) for x in df_dash_filt[c_ni].dropna().unique() if str(x)!='nan']))
                if sel_ni: 
                    df_dash_filt = df_dash_filt[df_dash_filt[c_ni].astype(str).isin(sel_ni)]
                    df_time_filt = df_time_filt[df_time_filt[c_ni].astype(str).isin(sel_ni)]
            if c_zo:
                sel_zo = f3.multiselect("Zona (UPSS)", sorted([str(x) for x in df_dash_filt[c_zo].dropna().unique() if str(x)!='nan']))
                if sel_zo: 
                    df_dash_filt = df_dash_filt[df_dash_filt[c_zo].astype(str).isin(sel_zo)]
                    df_time_filt = df_time_filt[df_time_filt[c_zo].astype(str).isin(sel_zo)]
            if c_am:
                sel_am = f4.multiselect("Ambiente", sorted([str(x) for x in df_dash_filt[c_am].dropna().unique() if str(x)!='nan']))
                if sel_am: 
                    df_dash_filt = df_dash_filt[df_dash_filt[c_am].astype(str).isin(sel_am)]
                    df_time_filt = df_time_filt[df_time_filt[c_am].astype(str).isin(sel_am)]
                
            sel_test = f5.multiselect("Grupo de Clash (VS)", sorted(list(df_dash_filt['Test'].unique())))
            if sel_test: 
                df_dash_filt = df_dash_filt[df_dash_filt['Test'].isin(sel_test)]
                df_time_filt = df_time_filt[df_time_filt['Test'].isin(sel_test)]

            tab_dash, tab_timeline, tab_details = st.tabs(["📊 Dashboard General", "⏱️ Línea de Tiempo", "🔎 Detalle de Clashes e IDs"])
            
            with tab_dash:
                tot = len(df_dash_filt)
                res = len(df_dash_filt[df_dash_filt[estado_col].isin(['Resuelto', 'Aprobado'])])
                avance_pct = (res / tot * 100) if tot > 0 else 0
                
                st.markdown("---")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("TOTAL CLASHES", tot)
                col2.metric("PENDIENTES", tot - res)
                col3.metric("RESUELTOS / APROB.", res)
                col4.metric("AVANCE GLOBAL (%)", f"{avance_pct:.1f} %")
                
                if tot > 0:
                    st.markdown("---")
                    
                    # 1. PRIMERO PROCESAMOS LA DATA
                    df_dash_filt['Es_Resuelto'] = df_dash_filt[estado_col].isin(['Resuelto', 'Aprobado'])
                    agrupado = df_dash_filt.groupby(['Test', 'Tolerancia_Grupo']).agg(Total=('Test', 'count'), Resueltos=('Es_Resuelto', 'sum')).reset_index()
                    agrupado['Orden_CL'] = agrupado['Test'].apply(lambda x: 1 if "(CL)" in str(x) else 0)
                    agrupado = agrupado.sort_values(by=['Orden_CL', 'Test']).drop(columns=['Orden_CL'])
                    agrupado['Total / Resueltos'] = agrupado['Total'].astype(str) + " / " + agrupado['Resueltos'].astype(str)
                    lista_clashes = agrupado[['Test', 'Tolerancia_Grupo', 'Total / Resueltos']].rename(columns={'Test': 'Grupo de Clash (VS)', 'Tolerancia_Grupo': 'Tolerancia'})
                    
                    # 2. CALCULAMOS EL NÚMERO DE GRUPOS
                    num_grupos = len(agrupado)
                    
                    # 3. LUEGO IMPRIMIMOS EL TÍTULO CON EL NÚMERO
                    st.markdown(f"### 📋 Lista de Interferencias &nbsp;<span style='font-size:16px; color:#f7931e; background-color:rgba(247,147,30,0.1); padding:4px 10px; border-radius:6px; vertical-align: middle;'>{num_grupos} Grupos</span>", unsafe_allow_html=True)
                    
                    def color_semaforo_lista(val):
                        try:
                            t_str, r_str = val.split(' / ')
                            t, r = int(t_str), int(r_str)
                            if t == 0: return ''
                            pct = r / t
                            if pct >= 0.85: return 'background-color: rgba(34, 197, 94, 0.2); color: #4ade80; font-weight: bold; text-align: center;'
                            elif pct >= 0.50: return 'background-color: rgba(249, 115, 22, 0.2); color: #fb923c; font-weight: bold; text-align: center;'
                            else: return 'background-color: rgba(239, 68, 68, 0.2); color: #f87171; font-weight: bold; text-align: center;'
                        except: return ''
                    
                    try: 
                        styled_lista = lista_clashes.style.hide(axis='index').map(color_semaforo_lista, subset=['Total / Resueltos'])
                    except AttributeError: 
                        styled_lista = lista_clashes.style.hide_index().applymap(color_semaforo_lista, subset=['Total / Resueltos'])
                    
                    st.table(styled_lista)
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer: lista_clashes.to_excel(writer, sheet_name='Lista', index=False)
                    st.download_button("📥 Descargar Tabla en Excel", data=buffer.getvalue(), file_name="Lista_Interferencias.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                    st.markdown("---")
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        with st.container(border=True):
                            st.markdown("#### Clashes por Zona (UPSS)")
                            if c_zo:
                                zona_data = df_dash_filt[c_zo].value_counts().reset_index()
                                zona_data.columns = ['Zona', 'Cantidad']
                                fig_z = px.bar(zona_data, x='Cantidad', y='Zona', orientation='h', color='Cantidad', color_continuous_scale='Reds')
                                fig_z.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", yaxis={'categoryorder':'total ascending'}, margin=dict(t=10, b=10, l=10, r=10))
                                st.plotly_chart(fig_z, use_container_width=True)
                            else: st.info("No se encontró columna de Zonas.")
                    
                    with c2:
                        with st.container(border=True):
                            st.markdown("#### Estado General")
                            fig_pie = px.pie(df_dash_filt, names=estado_col, hole=0.4, color=estado_col, color_discrete_map={'Resuelto':'#4ade80', 'Aprobado':'#22c55e', 'Activo':'#ef4444', 'Nuevo':'#f97316', 'Revisado':'#3b82f6'})
                            fig_pie.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", margin=dict(t=10, b=10, l=10, r=10))
                            st.plotly_chart(fig_pie, use_container_width=True)

                    c3, c4 = st.columns(2)
                    with c3:
                        with st.container(border=True):
                            st.markdown("#### Clashes por Nivel")
                            if c_ni:
                                ni_data = df_dash_filt[c_ni].value_counts().reset_index()
                                ni_data.columns = ['Nivel', 'Cantidad']
                                fig_n = px.bar(ni_data, x='Nivel', y='Cantidad', color='Cantidad', color_continuous_scale='Reds')
                                fig_n.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", xaxis_type='category', margin=dict(t=10, b=10, l=10, r=10))
                                st.plotly_chart(fig_n, use_container_width=True)
                            else: st.info("No se encontró columna de Niveles.")
                    
                    with c4:
                        with st.container(border=True):
                            st.markdown("#### Top 15 Ambientes Críticos")
                            if c_am:
                                amb_data = df_dash_filt[c_am].value_counts().head(15).reset_index()
                                amb_data.columns = ['Ambiente', 'Cantidad']
                                fig_a = px.bar(amb_data, x='Ambiente', y='Cantidad', color='Cantidad', color_continuous_scale='Oranges')
                                fig_a.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white", xaxis_tickangle=-45, xaxis_type='category', margin=dict(t=10, b=10, l=10, r=10))
                                st.plotly_chart(fig_a, use_container_width=True)
                            else: st.info("No se encontró columna de Ambientes.")

            with tab_timeline:
                st.markdown("### ⏱️ Evolución de Interferencias")
                date_col = next((col for col in df_time_filt.columns if 'fecha' in str(col).lower()), None)
                if date_col:
                    if 'ID de elemento' in df_time_filt.columns:
                        df_time_clean = df_time_filt.drop_duplicates(subset=['Test', 'Nombre de conflicto', date_col, estado_col, 'ID de elemento'])
                    else:
                        df_time_clean = df_time_filt.drop_duplicates(subset=['Test', 'Nombre de conflicto', date_col, estado_col])
                        
                    df_time_clean['Fecha_Limpia'] = pd.to_datetime(df_time_clean[date_col], errors='coerce').dt.strftime('%d/%m/%Y')
                    timeline_data = df_time_clean.groupby(['Fecha_Limpia', estado_col]).size().reset_index(name='Cantidad')
                    
                    fig_time = px.line(timeline_data, x='Fecha_Limpia', y='Cantidad', color=estado_col, markers=True)
                    fig_time.update_xaxes(type='category', categoryorder='category ascending')
                    fig_time.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="white")
                    st.plotly_chart(fig_time, use_container_width=True)
                else:
                    st.info("No se encontró una columna de Fecha.")

            with tab_details:
                st.markdown("### 🔎 Detalle de Clashes e IDs")
                cols_to_show = [c for c in ['Test', 'Nombre de conflicto', estado_col, 'ID de elemento', c_zo, c_ni, c_am] if c in df_dash_filt.columns]
                if not cols_to_show: cols_to_show = df_dash_filt.columns
                st.dataframe(df_dash_filt[cols_to_show], use_container_width=True)
else:
    st.info("No se encontraron archivos de Navisworks. Sube los reportes desde la interfaz o al repositorio en la nube.")

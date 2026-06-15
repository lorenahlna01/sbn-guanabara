import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import io
import urllib.request

# ==========================================
# ⚙️ CONFIGURAÇÕES (PREENCHA DEPOIS)
# ==========================================
# Quando tiver o Forms, cole o link de exportação aqui.
URL_FORMS_RESPOSTAS = "" 
URL_DRIVE_GERAL = "https://drive.google.com/drive/my-drive"

# --- IMPORTAÇÃO SEGURA DO FOLIUM COM CLUSTER ---
try:
    import folium
    from folium.plugins import MarkerCluster
    from streamlit_folium import st_folium
    HAS_FOLIUM = True
except ImportError:
    HAS_FOLIUM = False

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Guanabara Verde | Gestão", 
    page_icon="🌿", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- ESTILO CSS ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        .main { background-color: #020617; color: #F8FAFC; font-family: 'Inter', sans-serif; }
        [data-testid="stSidebar"] { background-color: #0F172A; border-right: 1px solid #1E293B; }
        .glass-card { background: rgba(30, 41, 59, 0.45); border: 1px solid rgba(255, 255, 255, 0.08); border-radius: 16px; padding: 22px; margin-bottom: 20px; backdrop-filter: blur(12px); }
        .kpi-title { font-size: 11px; color: #94A3B8; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
        .kpi-value { font-size: 32px; color: #10B981; font-weight: 800; margin-top: 5px; }
        .logo-text { font-size: 20px; font-weight: 800; color: #10B981; text-align: center; margin-bottom: 25px; }
        .stDataFrame { border-radius: 12px; overflow: hidden; border: 1px solid #1E293B; }
    </style>
""", unsafe_allow_html=True)

# --- MAPAS E COORDENADAS ---
COLOR_MAP = {
    "Agricultura Urbana e Periurbana": "#10B981", 
    "Aquicultura Urbana e Periurbana": "#3B82F6",      
    "Conforto Térmico Urbano": "#F59E0B",            
    "Não Especificado": "#64748B"
}

COORDS_MAP = {
    "rio de janeiro": [-22.9068, -43.1729], "niterói": [-22.8858, -43.1153],
    "magé": [-22.6514, -43.0401], "itaboraí": [-22.7486, -42.8594],
    "duque de caxias": [-22.7856, -43.3115], "são gonçalo": [-22.8269, -43.0539],
    "cachoeiras de macacu": [-22.4633, -42.6542], "seropédica": [-22.7441, -43.7121],
    "guapimirim": [-22.5364, -42.9818], "maricá": [-22.9194, -42.8181]
}

def obter_coordenadas(local):
    local_clean = str(local).strip().lower()
    for nome, lat_lon in COORDS_MAP.items():
        if nome in local_clean: return lat_lon
    return [-22.84, -43.15] 

# ==========================================
# 2. CARREGAMENTO E SANEAMENTO SEVERO
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    url = "https://docs.google.com/spreadsheets/d/1H5TMFJvuDpX9EWr-qVJO2vQPTEGB82We/export?format=xlsx"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            xls = pd.ExcelFile(io.BytesIO(response.read()))
        
        target_sheet = [s for s in xls.sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
        df_t = pd.read_excel(xls, sheet_name=target_sheet)
        df_t.columns = df_t.columns.str.strip()
        
        # Garante que todas as colunas necessárias existam
        colunas_obrig = [
            "KM Previsto", "Nº de Participantes", "Carga Horária (h)", "Local", 
            "Subtema", "Turma", "Status", "Tipo de Atividade", "Tipo Veículo", 
            "Deslocamento", "Hospedagem", "Público-Alvo", "Link Evidências", "Observações"
        ]
        for col in colunas_obrig:
            if col not in df_t.columns: 
                df_t[col] = "0" if "Nº" in col or "KM" in col or "Carga" in col else "Não Informado"

        # 1. PADRONIZAÇÃO DE EIXOS
        def map_eixo(val):
            val = str(val).strip().lower()
            if '1' in val or 'agricultura' in val or 'agroecologia' in val: return "Agricultura Urbana e Periurbana"
            if '2' in val or 'aquicultura' in val: return "Aquicultura Urbana e Periurbana"
            if '3' in val or 'térmico' in val or 'termico' in val: return "Conforto Térmico Urbano"
            return "Não Especificado"
        df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"

        # 2. PADRONIZAÇÃO DE TIPO DE ATIVIDADE
        def map_atividade(val):
            val = str(val).lower()
            if 'prática' in val or 'pratica' in val or 'oficina' in val: return 'Prática'
            if 'visita' in val or 'campo' in val: return 'Visita Técnica'
            return 'Teórica'
        df_t["Tipo de Atividade"] = df_t["Tipo de Atividade"].apply(map_atividade)

        # 3. LIMPEZA DE POLOS E STATUS PARA NÃO INFLAR OS NÚMEROS
        df_t["Local"] = df_t["Local"].astype(str).str.strip().str.title()
        df_t["Status"] = df_t["Status"].astype(str).str.strip().str.title()

        # 4. LIMPEZA NUMÉRICA
        df_t["Carga Horária (h)"] = pd.to_numeric(df_t["Carga Horária (h)"].astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
        df_t["Nº de Participantes"] = pd.to_numeric(df_t["Nº de Participantes"], errors='coerce').fillna(0)
        df_t["KM Previsto"] = pd.to_numeric(df_t["KM Previsto"], errors='coerce').fillna(0)
        
        # Lê todas as linhas válidas com base no número da turma
        df_t = df_t[df_t["Turma"].astype(str) != "nan"]
        
        return df_t, "✅ Conectado via Google Sheets"
    except Exception as e:
        return pd.DataFrame({
            "Eixo Temático": ["Erro"], "Local": ["Erro"], "Subtema": ["Erro"], 
            "Nº de Participantes": [0], "Carga Horária (h)": [0], "Status": ["Planejada"], 
            "Tipo de Atividade": ["Teórica"], "Turma": [1], "Público-Alvo": [""], 
            "Tipo Veículo": [""], "KM Previsto": [0], "Deslocamento": [""], 
            "Hospedagem": [""], "Observações": [""]
        }), f"❌ Falha de Conexão: {e}"

df_f, status_conexao = load_data()

# ==========================================
# 3. SIDEBAR E NAVEGAÇÃO
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("NAVEGAÇÃO", [
        "📌 VISÃO GERAL", 
        "📋 GESTÃO DE TURMAS", 
        "🗺️ MAPA DE ATUAÇÃO", 
        "🚐 LOGÍSTICA", 
        "📊 INDICADORES E METAS"
    ])
    st.markdown("---")
    st.info(status_conexao)
    if st.button("🔄 Atualizar Painel"):
        st.cache_data.clear()
        st.rerun()

# ==========================================
# PÁGINA 1: VISÃO GERAL
# ==========================================
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    
    total_ch = int(df_f["Carga Horária (h)"].sum())
    total_vagas = int(df_f["Nº de Participantes"].sum())
    polos_unicos = len(df_f["Local"].unique())
    
    # Cálculo real de progresso
    atividades_total = len(df_f)
    atividades_concluidas = len(df_f[df_f["Status"].str.contains("Concluíd", case=False, na=False)])
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária Global</div><div class="kpi-value">{total_ch}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Vagas (Planilha)</div><div class="kpi-value">{total_vagas}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Polos de Atuação</div><div class="kpi-value" style="color: #3B82F6;">{polos_unicos}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Progresso (Concluídas)</div><div class="kpi-value" style="color: #F59E0B;">{atividades_concluidas} / {atividades_total}</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        df_comp = df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index()
        fig_comp = go.Figure(go.Bar(x=df_comp["Eixo Temático"], y=df_comp["Nº de Participantes"], marker_color="#10B981"))
        fig_comp.update_layout(template="plotly_dark", title="Vagas por Eixo Temático", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_comp, use_container_width=True)
        
    with col_b:
        df_ativ = df_f.groupby("Tipo de Atividade")["Carga Horária (h)"].sum().reset_index()
        fig_ativ = px.pie(df_ativ, values="Carga Horária (h)", names="Tipo de Atividade", hole=0.4,
                          title="Equilíbrio Metodológico (CH por Categoria)", template="plotly_dark",
                          color_discrete_map={"Teórica": "#3B82F6", "Prática": "#10B981", "Visita Técnica": "#F59E0B"})
        fig_ativ.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_ativ, use_container_width=True)

# ==========================================
# PÁGINA 2: GESTÃO E FILTROS EM CASCATA
# ==========================================
elif menu == "📋 GESTÃO DE TURMAS":
    st.title("📋 Gestão e Filtros de Turmas")
    st.markdown("### 🔍 Filtros em Cascata")
    
    df_p2 = df_f.copy()
    col1, col2, col3 = st.columns(3)
    
    with col1:
        opcoes_eixo = sorted(df_p2["Eixo Temático"].dropna().unique())
        f_eixo = st.multiselect("Macrotema (Eixo)", options=opcoes_eixo)
        if f_eixo: df_p2 = df_p2[df_p2["Eixo Temático"].isin(f_eixo)]
            
    with col2:
        opcoes_tipo = sorted(df_p2["Tipo de Atividade"].dropna().unique())
        f_tipo = st.multiselect("Tipo de Atividade", options=opcoes_tipo)
        if f_tipo: df_p2 = df_p2[df_p2["Tipo de Atividade"].isin(f_tipo)]
            
    with col3:
        opcoes_polo = sorted(df_p2["Local"].dropna().unique())
        f_polo = st.multiselect("Polo (Local)", options=opcoes_polo)
        if f_polo: df_p2 = df_p2[df_p2["Local"].isin(f_polo)]
            
    st.markdown("---")
    if not df_p2.empty:
        st.success(f"✅ {len(df_p2)} atividades encontradas.")
        colunas_desejadas = ["Turma", "Eixo Temático", "Subtema", "Tipo de Atividade", "Local", "Público-Alvo", "Carga Horária (h)", "Status"]
        colunas_visiveis = [col for col in colunas_desejadas if col in df_p2.columns]
        st.dataframe(df_p2[colunas_visiveis], use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ Nenhuma atividade encontrada com essa combinação. Tente limpar algumas seleções.")

# ==========================================
# PÁGINA 3: MAPA DE ATUAÇÃO COM CLUSTER
# ==========================================
elif menu == "🗺️ MAPA DE ATUAÇÃO":
    st.title("🗺️ Mapeamento Territorial e Oficinas")
    st.markdown(f"Visualização de todas as **{len(df_f)} atividades**. Clique nos círculos com números para expandir as atividades agrupadas no mesmo polo.")
    
    st.markdown("""
        <div style="display:flex; gap: 20px; font-size:14px; color:#94A3B8; margin-bottom:15px;">
            <div>📖 Teóricas</div>
            <div>🔧 <b>Oficinas/Práticas</b></div>
            <div>🚌 Visitas Técnicas</div>
        </div>
    """, unsafe_allow_html=True)
    
    if HAS_FOLIUM:
        m = folium.Map(location=[-22.82, -43.12], zoom_start=9, tiles="CartoDB dark_matter")
        marker_cluster = MarkerCluster().add_to(m)
        
        for _, r in df_f.iterrows():
            lat, lon = obter_coordenadas(r['Local'])
            cor = COLOR_MAP.get(r['Eixo Temático'], "gray")
            tipo = str(r['Tipo de Atividade'])
            
            if tipo == 'Prática': icone_fa = "wrench"
            elif tipo == 'Visita Técnica': icone_fa = "bus"
            else: icone_fa = "book"
                
            obs = r.get('Observações', '')
            obs_texto = obs if str(obs).strip() and str(obs) != "nan" else "Nenhuma observação registrada."
            
            popup_html = f"""
            <div style='font-family: sans-serif; font-size:13px; min-width: 280px;'>
                <h4 style='color: {cor}; margin-bottom: 8px; margin-top: 0;'>{r['Subtema']}</h4>
                <span style='background-color:#1E293B; color:white; padding:3px 6px; border-radius:4px; font-size:11px; font-weight:bold;'>{tipo.upper()}</span><br><br>
                <b>📍 Polo:</b> {r['Local']}<br>
                <b>🌿 Eixo:</b> {r['Eixo Temático']}<br>
                <b>👥 Turma/Público:</b> Turma {r['Turma']} - {r['Público-Alvo']}<br>
                <b>⏱️ Carga Horária:</b> {r['Carga Horária (h)']}h<br>
                <b>📊 Status:</b> {r['Status']}<br>
                <hr style='margin: 8px 0; border: 0.5px solid #ccc;'>
                <span style='font-size:11.5px; color:#555;'><b>Obs:</b> {obs_texto}</span>
            </div>
            """
            
            folium.Marker(
                [lat, lon], 
                popup=folium.Popup(popup_html, max_width=350), 
                tooltip=f"CLIQUE AQUI - {tipo}: {r['Subtema']}", 
                icon=folium.Icon(color="white", icon_color=cor, icon=icone_fa, prefix='fa')
            ).add_to(marker_cluster)
            
        st_folium(m, width="100%", height=600)
    else:
        st.error("Biblioteca Folium não instalada. Exibindo mapa nativo simples.")
        st.map(pd.DataFrame({"lat": df_f["Local"].map(lambda x: obter_coordenadas(x)[0]), "lon": df_f["Local"].map(lambda x: obter_coordenadas(x)[1])}))

# ==========================================
# PÁGINA 4: LOGÍSTICA
# ==========================================
elif menu == "🚐 LOGÍSTICA":
    st.title("🚐 Controle Logístico (Deslocamento e Hospedagem)")
    st.markdown("Monitoramento de atividades com necessidades logísticas ativas.")
    
    df_log = df_f[(df_f["KM Previsto"] > 0) | (df_f["Deslocamento"].str.upper().str.contains("X|SIM|S", na=False)) | (df_f["Hospedagem"].str.upper().str.contains("X|SIM|ALOJAMENTO", na=False))]
    
    col1, col2 = st.columns(2)
    with col1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Demandas de Transporte (Total KM)</div><div class="kpi-value">{int(df_f["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with col2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Atividades com Logística Ativa</div><div class="kpi-value">{len(df_log)}</div></div>', unsafe_allow_html=True)
    
    colunas_logistica_desejadas = ["Local", "Público-Alvo", "Subtema", "Tipo de Atividade", "Tipo Veículo", "KM Previsto", "Deslocamento", "Hospedagem", "Observações"]
    colunas_log_visiveis = [col for col in colunas_logistica_desejadas if col in df_log.columns]
    
    st.dataframe(df_log[colunas_log_visiveis], use_container_width=True, hide_index=True)

# ==========================================
# PÁGINA 5: INDICADORES E METAS (GAC/BID)
# ==========================================
elif menu == "📊 INDICADORES E METAS":
    st.title("📊 Indicadores de Resultados (Tabela 1 - BID)")
    st.markdown("Painel estruturado conforme os Instrumentos de Coleta do Plano de Trabalho.")
    
    tab1, tab2 = st.tabs(["📉 Métricas Quantitativas (Inscrições e Projetos)", "🧠 Métricas Qualitativas (Avaliações PAP)"])
    
    with tab1:
        st.subheader("Dimensões de Participação e Equidade")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            st.markdown("### Participação (%)")
            st.info("Instrumento: Listas de Presença\n\n*Aguardando consolidação.*")
            st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=0, title={'text': "Taxa de Presença"}, gauge={'axis': {'range': [None, 100]}})), use_container_width=True)
            
        with c2:
            st.markdown("### Equidade de Gênero")
            st.info("Instrumento: Ficha de Inscrição\n\n*Meta Mínima: 40% Feminino*")
            st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=0, title={'text': "Mulheres Participantes"}, gauge={'axis': {'range': [None, 100]}, 'threshold': {'line': {'color': "red", 'width': 4}, 'thickness': 0.75, 'value': 40}})), use_container_width=True)
            
        with c3:
            st.markdown("### Juventude (18-35 anos)")
            st.info("Instrumento: Ficha de Inscrição\n\n*Meta Mínima: 20% Jovens*")
            st.plotly_chart(go.Figure(go.Indicator(mode="gauge+number", value=0, title={'text': "Participação Jovem"}, gauge={'axis': {'range': [None, 100]}, 'threshold': {'line': {'color': "red", 'width': 4}, 'thickness': 0.75, 'value': 20}})), use_container_width=True)

        st.markdown("---")
        st.subheader("Dimensão: Resultados das Oficinas")
        st.markdown("**Indicador:** Projetos conceituais desenvolvidos | **Instrumento:** Relatórios das Oficinas")
        st.warning("Consolidação do Portfólio de SBN - Nenhuma evidência de projeto inserida ainda.")

    with tab2:
        st.subheader("Dimensões Pedagógicas e de Apropriação")
        st.markdown("A coleta destes dados ocorre durante e no pós-capacitação.")
        
        col_q1, col_q2 = st.columns(2)
        with col_q1:
            st.markdown("""
            <div class="glass-card">
                <h4 style='color:#10B981;'>🧠 Engajamento e Aprendizagem</h4>
                <ul style='color:#94A3B8; font-size:14px;'>
                    <li><b>Nível de Engajamento:</b> Aguardando registro qualitativo das dinâmicas.</li>
                    <li><b>Apropriação dos Conteúdos:</b> Aguardando evidências de aplicação prática.</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        with col_q2:
            st.markdown("""
            <div class="glass-card">
                <h4 style='color:#3B82F6;'>🤝 Qualidade, Aplicabilidade e Inclusão</h4>
                <ul style='color:#94A3B8; font-size:14px;'>
                    <li><b>Avaliação da Metodologia:</b> Aguardando formulários de satisfação (Pós-capacitação).</li>
                    <li><b>Aplicabilidade Territorial:</b> Aguardando rodas de conversa.</li>
                    <li><b>Ambiente Seguro e Inclusivo:</b> Avaliação sobre respeito e escuta.</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Realização SEAS-RJ & BID</p>", unsafe_allow_html=True)

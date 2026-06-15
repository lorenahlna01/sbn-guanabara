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
    page_title="Guanabara Verde | Gestão Estratégica", 
    page_icon="🌿", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- ESTILO CSS EXECUTIVO ---
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
# 2. CARREGAMENTO DA NOVA PLANILHA
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    url = "https://docs.google.com/spreadsheets/d/1qBJ-Dk_AEvZx8zPg5y2fsDV7VvO_arNy/export?format=xlsx"
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
            "Deslocamento", "Hospedagem", "Público-Alvo", "Link Evidências", "Observações", "Data"
        ]
        for col in colunas_obrig:
            if col not in df_t.columns: 
                df_t[col] = "0" if "Nº" in col or "KM" in col or "Carga" in col else "Não Informado"

        # PADRONIZAÇÃO DE EIXOS
        def map_eixo(val):
            val = str(val).strip().lower()
            if '1' in val or 'agricultura' in val or 'agroecologia' in val: return "Agricultura Urbana e Periurbana"
            if '2' in val or 'aquicultura' in val: return "Aquicultura Urbana e Periurbana"
            if '3' in val or 'térmico' in val or 'termico' in val: return "Conforto Térmico Urbano"
            return "Não Especificado"
        df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"

        # PADRONIZAÇÃO DE TIPO DE ATIVIDADE
        def map_atividade(val):
            val = str(val).lower()
            if 'prática' in val or 'pratica' in val or 'oficina' in val: return 'Prática'
            if 'visita' in val or 'campo' in val: return 'Visita Técnica'
            return 'Teórica'
        df_t["Tipo de Atividade"] = df_t["Tipo de Atividade"].apply(map_atividade)

        # LIMPEZA
        df_t["Local"] = df_t["Local"].astype(str).str.strip().str.title()
        df_t["Status"] = df_t["Status"].astype(str).str.strip().str.title()
        df_t["Carga Horária (h)"] = pd.to_numeric(df_t["Carga Horária (h)"].astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
        df_t["Nº de Participantes"] = pd.to_numeric(df_t["Nº de Participantes"], errors='coerce').fillna(0)
        df_t["KM Previsto"] = pd.to_numeric(df_t["KM Previsto"], errors='coerce').fillna(0)
        
        df_t = df_t[df_t["Turma"].astype(str) != "nan"]
        
        return df_t, "✅ Conectado à Nova Planilha via Google Sheets"
    except Exception as e:
        return pd.DataFrame({
            "Eixo Temático": ["Erro"], "Local": ["Erro"], "Subtema": ["Erro"], 
            "Nº de Participantes": [0], "Carga Horária (h)": [0], "Status": ["Planejada"], 
            "Tipo de Atividade": ["Teórica"], "Turma": [1], "Público-Alvo": [""], 
            "Tipo Veículo": [""], "KM Previsto": [0], "Deslocamento": [""], 
            "Hospedagem": [""], "Observações": [""], "Data": [""]
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
        "📊 INDICADORES E METAS",
        "📈 ANÁLISE AVANÇADA"
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
    atividades_total = len(df_f)
    atividades_concluidas = len(df_f[df_f["Status"].str.contains("Concluíd", case=False, na=False)])
    
    # KPIs principais (Polos removido)
    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária Global</div><div class="kpi-value">{total_ch}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Vagas (Planilha)</div><div class="kpi-value">{total_vagas}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Atividades Concluídas</div><div class="kpi-value" style="color: #F59E0B;">{atividades_concluidas} / {atividades_total}</div></div>', unsafe_allow_html=True)
    
    # Radar de Próximas Atividades
    st.markdown("### 🗓️ Radar: Próximas Atividades Planejadas")
    if "Data" in df_f.columns:
        df_radar = df_f[df_f["Status"].str.contains("Planejada", case=False, na=False)].copy()
        df_radar["Data"] = pd.to_datetime(df_radar["Data"], errors="coerce")
        df_radar = df_radar.dropna(subset=["Data"]).sort_values("Data").head(5)
        
        if not df_radar.empty:
            st.dataframe(df_radar[["Data", "Local", "Subtema", "Tipo de Atividade"]], use_container_width=True, hide_index=True)
        else:
            st.info("Não há datas futuras preenchidas corretamente na planilha no momento.")
    else:
        st.info("A coluna 'Data' precisa estar preenchida para ativar o radar.")

    st.markdown("---")
    
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
# PÁGINA 2: GESTÃO E EXPORTAÇÃO
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
        opcoes_pub = sorted(df_p2["Público-Alvo"].dropna().unique())
        f_pub = st.multiselect("Público-Alvo", options=opcoes_pub)
        if f_pub: df_p2 = df_p2[df_p2["Público-Alvo"].isin(f_pub)]
            
    st.markdown("---")
    if not df_p2.empty:
        st.success(f"✅ {len(df_p2)} atividades encontradas.")
        colunas_desejadas = ["Turma", "Eixo Temático", "Subtema", "Tipo de Atividade", "Local", "Público-Alvo", "Carga Horária (h)", "Status", "Data"]
        colunas_visiveis = [col for col in colunas_desejadas if col in df_p2.columns]
        
        st.dataframe(df_p2[colunas_visiveis], use_container_width=True, hide_index=True)
        
        # Botão de Exportação
        csv = df_p2[colunas_visiveis].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Baixar Tabela Filtrada (CSV)",
            data=csv,
            file_name='atividades_filtradas.csv',
            mime='text/csv',
        )
    else:
        st.warning("⚠️ Nenhuma atividade encontrada com essa combinação.")

# ==========================================
# PÁGINA 3: MAPA DE ATUAÇÃO (CAMPO)
# ==========================================
elif menu == "🗺️ MAPA DE ATUAÇÃO":
    st.title("🗺️ Mapeamento Territorial de Campo")
    
    # Exibe apenas atividades práticas e visitas técnicas
    df_mapa = df_f[df_f["Tipo de Atividade"].isin(['Prática', 'Visita Técnica'])].copy()
    
    opcoes_mapa = ["Todas as Atividades de Campo"] + sorted(df_mapa["Tipo de Atividade"].unique().tolist())
    filtro_mapa = st.selectbox("🔍 Selecione o que deseja visualizar no mapa:", options=opcoes_mapa)
    
    if filtro_mapa != "Todas as Atividades de Campo":
        df_mapa = df_mapa[df_mapa["Tipo de Atividade"] == filtro_mapa]

    st.markdown(f"Visualizando **{len(df_mapa)} atividades** (Aulas Teóricas ocultadas). Clique nos círculos com números para expandir.")
    
    st.markdown("""
        <div style="display:flex; gap: 20px; font-size:14px; color:#94A3B8; margin-bottom:15px;">
            <div>🔧 <b>Oficinas/Práticas</b></div>
            <div>🚌 <b>Visitas Técnicas</b></div>
        </div>
    """, unsafe_allow_html=True)
    
    if HAS_FOLIUM:
        m = folium.Map(location=[-22.82, -43.12], zoom_start=9, tiles="CartoDB dark_matter")
        marker_cluster = MarkerCluster().add_to(m)
        
        for _, r in df_mapa.iterrows():
            lat, lon = obter_coordenadas(r['Local'])
            cor = COLOR_MAP.get(r['Eixo Temático'], "gray")
            tipo = str(r['Tipo de Atividade'])
            
            icone_fa = "wrench" if tipo == 'Prática' else "bus"
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
        st.error("Biblioteca Folium não instalada.")

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
    
    # Eixo Temático adicionado no início
    colunas_logistica_desejadas = ["Eixo Temático", "Local", "Público-Alvo", "Subtema", "Tipo de Atividade", "Tipo Veículo", "KM Previsto", "Deslocamento", "Hospedagem", "Observações"]
    colunas_log_visiveis = [col for col in colunas_logistica_desejadas if col in df_log.columns]
    
    st.dataframe(df_log[colunas_log_visiveis], use_container_width=True, hide_index=True)

# ==========================================
# PÁGINA 5: INDICADORES E METAS
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

# ==========================================
# PÁGINA 6: ANÁLISE AVANÇADA
# ==========================================
elif menu == "📈 ANÁLISE AVANÇADA":
    st.title("📈 Análise Avançada e Inteligência de Dados")
    st.markdown("Cruzamentos estratégicos para tomada de decisão e relatórios executivos (BID).")

    st.markdown("---")
    st.subheader("🔥 Distribuição de Esforço (Carga Horária)")
    st.markdown("Identifique rapidamente onde o maior volume de horas do projeto está concentrado.")
    
    if not df_f.empty:
        pivot_df = df_f.pivot_table(
            index="Eixo Temático", 
            columns="Público-Alvo", 
            values="Carga Horária (h)", 
            aggfunc="sum"
        ).fillna(0)
        
        fig_heat = px.imshow(
            pivot_df, 
            text_auto=True, 
            aspect="auto", 
            color_continuous_scale="Greens", 
            template="plotly_dark",
            labels=dict(color="Horas Totais")
        )
        fig_heat.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("---")
    st.subheader("🚐 Eficiência Logística Sustentável (KM / Participante)")
    st.markdown("Avalia o impacto logístico. Atividades no topo da lista possuem **alto custo de deslocamento para baixo volume de pessoas**.")
    
    df_eff = df_f[df_f["KM Previsto"] > 0].copy()
    if not df_eff.empty:
        df_eff["KM por Participante"] = np.where(
            df_eff["Nº de Participantes"] > 0, 
            (df_eff["KM Previsto"] / df_eff["Nº de Participantes"]).round(1), 
            0
        )
        
        df_eff_top = df_eff.sort_values("KM por Participante", ascending=False).head(5)
        colunas_eff = ["Eixo Temático", "Local", "Subtema", "KM Previsto", "Nº de Participantes", "KM por Participante"]
        
        col_view = [col for col in colunas_eff if col in df_eff_top.columns]
        st.dataframe(df_eff_top[col_view], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma atividade com quilometragem prevista para calcular eficiência.")

    st.markdown("---")
    st.subheader("⚠️ Diagnóstico Operacional Automático")
    
    riscos = 0
    turmas_ociosas = df_f[df_f["Nº de Participantes"] < 10]
    if len(turmas_ociosas) > 0:
        st.warning(f"**Risco de Ociosidade:** Existem {len(turmas_ociosas)} atividades planejadas com menos de 10 participantes.")
        riscos += 1
        
    ch_zerada = df_f[df_f["Carga Horária (h)"] <= 0]
    if len(ch_zerada) > 0:
        st.error(f"**Dados Incompletos:** Foram encontradas {len(ch_zerada)} atividades com Carga Horária zerada ou não preenchida na planilha.")
        riscos += 1
        
    if riscos == 0:
        st.success("✅ **Excelente!** Nenhum risco operacional ou falha de preenchimento crítico foi detectado nos dados atuais.")

st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Realização SEAS-RJ & BID</p>", unsafe_allow_html=True)

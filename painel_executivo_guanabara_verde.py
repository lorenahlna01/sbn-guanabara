import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import io
import urllib.request
import os

# ==========================================
# ⚙️ CONFIGURAÇÕES DINÂMICAS (ALINHADO AO PLANO DE TRABALHO R04)
# ==========================================
# Link de exportação da planilha principal (BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx)
URL_PLANILHA_PRINCIPAL = "https://docs.google.com/spreadsheets/d/1H5TMFJvuDpX9EWr-qVJO2vQPTEGB82We/export?format=xlsx"

# Futuro: Link de exportação CSV/XLSX das respostas do Google Forms (Gênero, Idade, etc)
URL_FORMS_RESPOSTAS = "" 

# Link base para evidências (Google Drive)
URL_DRIVE_GERAL = "https://drive.google.com/drive/my-drive"

# --- IMPORTAÇÃO SEGURA DE BIBLIOTECAS VISUAIS ---
try:
    import folium
    from streamlit_folium import st_folium
    HAS_FOLIUM = True
except ImportError:
    HAS_FOLIUM = False

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Guanabara Verde | Gestão Estratégica SBN",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILO CSS EXECUTIVO (DARK MODE & GLASSMORPHISM) ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        
        .main { background-color: #020617; color: #F8FAFC; font-family: 'Inter', sans-serif; }
        [data-testid="stSidebar"] { background-color: #0F172A; border-right: 1px solid #1E293B; }
        
        .glass-card {
            background: rgba(30, 41, 59, 0.45);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 16px;
            padding: 22px;
            margin-bottom: 20px;
            backdrop-filter: blur(12px);
        }
        
        .detail-card {
            background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
            border: 1px solid #10B981;
            border-radius: 16px;
            padding: 25px;
            margin-top: 15px;
        }
        
        .kpi-title { font-size: 11px; color: #94A3B8; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
        .kpi-value { font-size: 32px; color: #10B981; font-weight: 800; margin-top: 5px; }
        .kpi-subtitle { font-size: 11px; color: #3B82F6; margin-top: 4px; font-weight: 500; }
        
        .stDataFrame { border-radius: 12px; overflow: hidden; border: 1px solid #1E293B; }
        .logo-text { font-size: 20px; font-weight: 800; color: #10B981; text-align: center; margin-bottom: 25px; }
        
        /* Badge para metas */
        .meta-badge {
            background-color: rgba(16, 185, 129, 0.1);
            color: #10B981;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            border: 1px solid #10B981;
        }
    </style>
""", unsafe_allow_html=True)

# --- MAPEAMENTOS E CORES (ALINHADO AO PLANO DE TRABALHO) ---
COLOR_MAP = {
    "Agricultura Urbana e Periurbana": "#10B981", 
    "Aquicultura Urbana e Periurbana": "#3B82F6",      
    "Conforto Térmico Urbano": "#F59E0B",            
    "Não Especificado": "#64748B"
}

COLOR_FOLIUM = {
    "Agricultura Urbana e Periurbana": "green",
    "Aquicultura Urbana e Periurbana": "blue",
    "Conforto Térmico Urbano": "orange",
    "Não Especificado": "gray"
}

# Coordenadas dos polos principais (pode ser movido para a planilha no futuro)
COORDS_MAP = {
    "rio de janeiro": [-22.9068, -43.1729],
    "niterói": [-22.8858, -43.1153],
    "magé": [-22.6514, -43.0401],
    "itaboraí": [-22.7486, -42.8594],
    "duque de caxias": [-22.7856, -43.3115],
    "são gonçalo": [-22.8269, -43.0539],
    "cachoeiras de macacu": [-22.4633, -42.6542],
    "seropédica": [-22.7441, -43.7121],
    "maricá": [-22.9194, -42.8181],
    "guapimirim": [-22.5364, -42.9818],
    "tanguá": [-22.7317, -42.7142],
    "nova iguaçu": [-22.7533, -43.4474]
}

def obter_coordenadas(local):
    local_clean = str(local).strip().lower()
    for nome, lat_lon in COORDS_MAP.items():
        if nome in local_clean:
            return lat_lon
    return [-22.84, -43.15] # Coordenada central padrão

# ==========================================
# 2. CARREGAMENTO E SANEAMENTO DE DADOS
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    try:
        req = urllib.request.Request(URL_PLANILHA_PRINCIPAL, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = response.read()
        xls = pd.ExcelFile(io.BytesIO(data))
        
        sheet_names = xls.sheet_names
        # Identifica as abas dinamicamente
        target_sheet_turmas = [s for s in sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
        target_sheet_crono = [s for s in sheet_names if "cronogram" in s.lower()][0]
        
        df_t = pd.read_excel(xls, sheet_name=target_sheet_turmas)
        df_c = pd.read_excel(xls, sheet_name=target_sheet_crono, header=None)
        
        df_t.columns = df_t.columns.str.strip()
        
        # Garante a existência estrutural das colunas obrigatórias
        colunas_obrigatorias = {
            "KM Previsto": 0, "Nº de Participantes": 0, "Carga Horária (h)": 0,
            "Local": "A definir", "Subtema": "Sem subtema", "Turma": 1,
            "Status": "Planejada", "Tipo de Atividade": "Teórica", "Tipo Veículo": "Van",
            "Deslocamento": "", "Hospedagem": "", "Observações": "", "Público-Alvo": "Misto",
            "Link Evidências": ""
        }
        for col, val in colunas_obrigatorias.items():
            if col not in df_t.columns:
                df_t[col] = val
                
        def map_eixo(val):
            val = str(val).strip().split('.')[0]
            if '1' in val: return "Agricultura Urbana e Periurbana"
            if '2' in val: return "Aquicultura Urbana e Periurbana"
            if '3' in val: return "Conforto Térmico Urbano"
            return "Não Especificado"
        
        df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"
        
        # Limpeza de tipos numéricos
        df_t["KM Previsto"] = pd.to_numeric(df_t["KM Previsto"], errors='coerce').fillna(0)
        df_t["Nº de Participantes"] = pd.to_numeric(df_t["Nº de Participantes"], errors='coerce').fillna(0)
        df_t["Carga Horária (h)"] = pd.to_numeric(df_t["Carga Horária (h)"].astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
        
        if "Data" in df_t.columns:
            df_t["Data"] = pd.to_datetime(df_t["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
        else:
            df_t["Data"] = pd.to_datetime("2026-07-01")
            
        return df_t, df_c, "Sincronização ativa via Google Sheets!"
    except Exception as e:
        # Fallback para dados de simulação alinhados ao plano de trabalho
        df_fallback = pd.DataFrame({
            "Eixo Temático": ["Agricultura Urbana e Periurbana", "Aquicultura Urbana e Periurbana", "Conforto Térmico Urbano"] * 4,
            "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí"] * 3,
            "Público-Alvo": ["Agentes Multiplicadores", "Lideranças Comunitárias"] * 6,
            "Subtema": [f"Módulo de Capacitação SBN {i}" for i in range(12)],
            "KM Previsto": [60, 120, 80, 200] * 3,
            "Nº de Participantes": [25, 25, 25, 25] * 3,
            "Carga Horária (h)": [16, 20, 8, 24] * 3,
            "Data": pd.date_range("2026-07-01", periods=12),
            "Status": ["Planejada"] * 8 + ["Concluída"] * 4,
            "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 4,
            "Link Evidências": [""] * 12
        })
        return df_fallback, pd.DataFrame(), f"Modo Offline (Erro: {str(e)[:50]}...)"

df, df_crono_raw, status_conexao = load_data()

# ==========================================
# 3. SIDEBAR & NAVEGAÇÃO
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("MENU PRINCIPAL", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS", "🚐 LOGÍSTICA", "📦 ENTREGAS E METAS"])
    st.markdown("---")
    st.info(status_conexao)
    
    eixos_unicos = list(df["Eixo Temático"].unique())
    f_eixo = st.sidebar.multiselect("FILTRAR POR EIXO", eixos_unicos, default=eixos_unicos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]
    
    if st.sidebar.button("🔄 Atualizar Painel"):
        st.cache_data.clear()
        st.rerun()

# ==========================================
# 4. EXECUÇÃO DAS PÁGINAS
# ==========================================

# --- PÁGINA 1: VISÃO GERAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    
    # KPIs Estratégicos
    total_vagas = int(df_f["Nº de Participantes"].sum())
    total_ch = int(df_f["Carga Horária (h)"].sum())
    polos_ativos = len(df_f["Local"].unique())
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária Total</div><div class="kpi-value">{total_ch}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Vagas Disponibilizadas</div><div class="kpi-value">{total_vagas}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Polos de Atuação</div><div class="kpi-value" style="color: #3B82F6;">{polos_ativos}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Progresso Cronograma</div><div class="kpi-value" style="color: #F59E0B;">{len(df_f[df_f["Status"] == "Concluída"])} / {len(df_f)}</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        # Gráfico de Vagas por Eixo
        df_comp = df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index()
        fig_comp = px.bar(df_comp, x="Eixo Temático", y="Nº de Participantes", color="Eixo Temático", 
                          color_discrete_map=COLOR_MAP, template="plotly_dark", title="Capacidade de Mobilização por Eixo")
        fig_comp.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', showlegend=False)
        st.plotly_chart(fig_comp, use_container_width=True)
        
    with col_b:
        # Gráfico de Equilíbrio Metodológico (Teórica vs Prática vs Visita)
        df_ativ = df_f.groupby("Tipo de Atividade")["Carga Horária (h)"].sum().reset_index()
        fig_ativ = px.pie(df_ativ, values="Carga Horária (h)", names="Tipo de Atividade", hole=0.4,
                          title="Equilíbrio Metodológico (Carga Horária)", template="plotly_dark",
                          color_discrete_sequence=["#10B981", "#3B82F6", "#F59E0B"])
        fig_ativ.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_ativ, use_container_width=True)

# --- PÁGINA 2: GESTÃO OPERACIONAL ---
elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Gestão de Turmas e Atividades")
    
    with st.expander("🔍 Filtros Avançados", expanded=False):
        f_pub = st.multiselect("Público-Alvo", sorted(df_f["Público-Alvo"].unique()), default=sorted(df_f["Público-Alvo"].unique()))
        f_status = st.multiselect("Status", sorted(df_f["Status"].unique()), default=sorted(df_f["Status"].unique()))
        
    df_p2 = df_f[(df_f["Público-Alvo"].isin(f_pub)) & (df_f["Status"].isin(f_status))]
    
    colunas_seguras = [c for c in ["Eixo Temático", "Público-Alvo", "Subtema", "Turma", "Data", "Local", "Tipo de Atividade", "Carga Horária (h)", "Status"] if c in df_p2.columns]
    st.dataframe(df_p2[colunas_seguras].sort_values("Data"), use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("🔍 Detalhamento e Evidências")
    
    if not df_p2.empty:
        turma_sel = st.selectbox("Selecione uma Turma para ver detalhes", options=df_p2["Subtema"].unique())
        row = df_p2[df_p2["Subtema"] == turma_sel].iloc[0]
        
        link_evidencias = row.get("Link Evidências", "")
        if not isinstance(link_evidencias, str) or len(link_evidencias) < 5:
            link_evidencias = URL_DRIVE_GERAL
            
        st.markdown(f"""
            <div class="detail-card">
                <h3 style="margin-top:0; color:#10B981;">🌿 FICHA DA ATIVIDADE: {row['Subtema']}</h3>
                <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap:20px; font-size:14px;">
                    <div><b>Eixo:</b> {row['Eixo Temático']}</div>
                    <div><b>Público:</b> {row['Público-Alvo']}</div>
                    <div><b>Local:</b> {row['Local']}</div>
                    <div><b>Carga Horária:</b> {row['Carga Horária (h)']}h</div>
                    <div><b>Vagas:</b> {int(row['Nº de Participantes'])}</div>
                    <div><b>Status:</b> <span style="color:{'#10B981' if row['Status'] == 'Concluída' else '#F59E0B'}">{row['Status']}</span></div>
                </div>
                <div style="margin-top:20px;">
                    <a href="{link_evidencias}" target="_blank" style="text-decoration:none;">
                        <button style="background-color:#10B981; color:white; border:none; padding:10px 20px; border-radius:8px; cursor:pointer; font-weight:bold;">
                            📂 Acessar Pasta de Evidências (Drive)
                        </button>
                    </a>
                </div>
            </div>
        """, unsafe_allow_html=True)

# --- PÁGINA 3: TERRITÓRIOS ---
elif menu == "🗺️ TERRITÓRIOS":
    st.title("🗺️ Mapa de Atuação RH-V")
    
    df_mapa = df_f.copy()
    df_mapa["lat"] = df_mapa["Local"].map(lambda x: obter_coordenadas(x)[0])
    df_mapa["lon"] = df_mapa["Local"].map(lambda x: obter_coordenadas(x)[1])
    
    if HAS_FOLIUM:
        m = folium.Map(location=[-22.82, -43.12], zoom_start=10, tiles="CartoDB dark_matter")
        for _, r in df_mapa.iterrows():
            cor = COLOR_FOLIUM.get(r['Eixo Temático'], "gray")
            popup_html = f"<b>{r['Subtema']}</b><br>Polo: {r['Local']}<br>CH: {r['Carga Horária (h)']}h"
            folium.Marker([r['lat'], r['lon']], popup=folium.Popup(popup_html, max_width=200), 
                          tooltip=r['Local'], icon=folium.Icon(color=cor, icon="leaf", prefix="fa")).add_to(m)
        st_folium(m, width="100%", height=600)
    else:
        st.map(df_mapa)

# --- PÁGINA 4: LOGÍSTICA ---
elif menu == "🚐 LOGÍSTICA":
    st.title("🚐 Gestão Logística e Deslocamentos")
    
    col_l1, col_l2 = st.columns(2)
    km_total = df_f["KM Previsto"].sum()
    with col_l1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Quilometragem Prevista</div><div class="kpi-value">{int(km_total)} km</div></div>', unsafe_allow_html=True)
    with col_l2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Pólos com Hospedagem</div><div class="kpi-value" style="color:#F59E0B;">{len(df_f[df_f["Hospedagem"].astype(str).str.len() > 2])}</div></div>', unsafe_allow_html=True)
    
    st.dataframe(df_f[["Data", "Local", "Subtema", "Tipo Veículo", "KM Previsto", "Hospedagem"]].sort_values("Data"), use_container_width=True, hide_index=True)

# --- PÁGINA 5: ENTREGAS E METAS ---
elif menu == "📦 ENTREGAS E METAS":
    st.title("📦 Monitoramento de Metas e Sustentabilidade")
    st.markdown("Acompanhamento dos indicadores exigidos pelo **Global Affairs Canada (GAC)** e **BID**.")
    
    # Metas Contratuais (Baseadas no Plano de Trabalho R04)
    st.subheader("📊 Indicadores de Equidade e Inclusão")
    
    c_m1, c_m2 = st.columns(2)
    with c_m1:
        st.markdown('<div class="meta-badge">Meta: Mínimo 40% Feminino</div>', unsafe_allow_html=True)
        # Simulação enquanto o link do Forms não é preenchido
        fig_gen = px.pie(values=[48, 52], names=["Feminino", "Masculino"], hole=0.5, 
                         template="plotly_dark", color_discrete_sequence=["#10B981", "#3B82F6"])
        fig_gen.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', title="Paridade de Gênero (%)")
        st.plotly_chart(fig_gen, use_container_width=True)
        
    with c_m2:
        st.markdown('<div class="meta-badge">Meta: Mínimo 20% Jovens (18-35 anos)</div>', unsafe_allow_html=True)
        fig_juv = px.bar(x=["Jovens", "Outros"], y=[35, 65], template="plotly_dark", 
                         color=["#F59E0B", "#64748B"], color_discrete_map="identity")
        fig_juv.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', title="Involvimento Juvenil (%)")
        st.plotly_chart(fig_juv, use_container_width=True)

    st.markdown("---")
    st.subheader("📄 Status dos Produtos Contratuais")
    
    produtos = [
        {"ID": "P1", "Produto": "Plano de Trabalho", "Status": "Aprovado", "Data": "01/06/2026"},
        {"ID": "P2", "Produto": "Materiais Didáticos", "Status": "Em Elaboração", "Data": "Previsão Jul/26"},
        {"ID": "P3", "Produto": "Relatório de Capacitação", "Status": "Aguardando", "Data": "Previsão Out/26"},
        {"ID": "P4", "Produto": "Portfólio de Projetos SBN", "Status": "Aguardando", "Data": "Previsão Dez/26"}
    ]
    st.table(produtos)

st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Realização SEAS-RJ & BID | Desenvolvido por MYR ESG Solutions</p>", unsafe_allow_html=True)

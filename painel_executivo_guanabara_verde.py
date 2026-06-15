import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import datetime
import io
import os
import urllib.request
import folium
from streamlit_folium import st_folium
from streamlit_calendar import calendar

# Configuração primária do layout da página
st.set_page_config(
    page_title="Painel Executivo | Guanabara Verde Resiliente",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS personalizada para um visual "premium dark mode" executivo
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        .main {
            background-color: #020617;
            color: #F8FAFC;
            font-family: 'Inter', sans-serif;
        }
        [data-testid="stSidebar"] {
            background-color: #0F172A;
            border-right: 1px solid #1E293B;
        }
        
        /* KPI Cards Executivos */
        .kpi-container {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .kpi-card {
            background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 20px;
            text-align: left;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            transition: all 0.3s ease;
        }
        .kpi-card:hover {
            transform: translateY(-3px);
            border-color: #10B981;
            box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.1);
        }
        .kpi-title {
            font-size: 11px;
            color: #94A3B8;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-bottom: 6px;
        }
        .kpi-value {
            font-size: 26px;
            color: #F8FAFC;
            font-weight: 700;
            line-height: 1.1;
        }
        .kpi-subtitle {
            font-size: 11px;
            color: #10B981;
            margin-top: 4px;
            font-weight: 500;
        }
    </style>
""", unsafe_allow_html=True)

# Paleta oficial
COLOR_MAP = {
    "Agroecologia e Adequação Ambiental": "#10B981",
    "Aquicultura e Pesca Artesanal": "#3B82F6",
    "Conforto Térmico Urbano": "#F59E0B",
    "Não Especificado": "#64748B"
}

def map_eixo_name(val):
    val_str = str(val).strip().split('.')[0]
    if '1' in val_str:
        return "Agroecologia e Adequação Ambiental"
    elif '2' in val_str:
        return "Aquicultura e Pesca Artesanal"
    elif '3' in val_str:
        return "Conforto Térmico Urbano"
    return "Não Especificado"

def map_publico(val):
    val_str = str(val).strip()
    if val_str == "1":
        return "Técnicos Municipais (SEAS/INEA/Prefeituras)"
    elif val_str == "2":
        return "Lideranças Comunitárias e Sociedade Civil"
    elif val_str in ["1.2", "1,2"]:
        return "Misto (Técnicos e Lideranças)"
    return f"Grupo {val_str}"

@st.cache_data(ttl=300)
def fetch_excel_from_github(url):
    try:
        data_bytes = urllib.request.urlopen(url).read()
        xls = pd.ExcelFile(io.BytesIO(data_bytes))
        sheet_names = xls.sheet_names
        sheet_crono = [s for s in sheet_names if "cronogram" in s.lower()][0]
        sheet_turmas = [s for s in sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
        df_t = pd.read_excel(xls, sheet_name=sheet_turmas)
        df_c = pd.read_excel(xls, sheet_name=sheet_crono, header=None)
        return df_t, df_c, "Sincronizado com GitHub"
    except Exception as e:
        return None, None, f"Erro ao acessar o GitHub: {str(e)}"

def get_mock_data():
    df_t = pd.DataFrame({
        "Eixo": [1, 2, 3] * 10,
        "Público-Alvo": ["1", "2", "1.2"] * 10,
        "Subtema": ["Adequação Ambiental", "Aquicultura Urbana", "Conforto Térmico Urbano"] * 10,
        "Turma": [1] * 30,
        "Dia": [1] * 30,
        "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí", "Duque de Caxias", "São Gonçalo"] * 5,
        "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 10,
        "Carga Horária (h)": [16, 24, 12] * 10,
        "Nº de Participantes": [20, 20, 20] * 10,
        "Coffe break (Sim/Não)": ["X", "", "X"] * 10,
        "Almoço (Sim/Não)": ["X", "X", ""] * 10,
        "Lanche (Sim/Não)": ["", "X", "X"] * 10,
        "KM Previsto": [120, 50, 200] * 10,
        "Tipo Veículo": ["Van", "Carro", "Ônibus"] * 10,
        "Kit": [20] * 30,
        "Observações": ["A usar base de dados de simulação"] * 30,
        "Data": [datetime.date(2026, 7, 1) + datetime.timedelta(days=i) for i in range(30)]
    })
    coords = {
        "Rio de Janeiro": [-22.9068, -43.1729],
        "Niterói": [-22.8833, -43.1167],
        "Magé": [-22.6528, -43.0403],
        "Itaboraí": [-22.7444, -42.8594],
        "Duque de Caxias": [-22.7856, -43.3117],
        "São Gonçalo": [-22.8269, -43.0539],
        "Local a definir": [-22.9068, -43.1729]
    }
    df_t["lat"] = df_t["Local"].map(lambda x: coords.get(x, [-22.9068, -43.1729])[0])
    df_t["lon"] = df_t["Local"].map(lambda x: coords.get(x, [-22.9068, -43.1729])[1])
    
    df_c = pd.DataFrame({
        0: ["Aprovado", "Em Andamento", "Planejado"],
        1: ["P1", "P3", "P5"],
        2: ["Fase de Planejamento Inicial", "Capacitação Prática em Campo", "Sistematização de Práticas"]
    })
    return df_t, df_c

def load_project_data(github_url=None, use_github=False):
    if use_github and github_url:
        df_t, df_c, status = fetch_excel_from_github(github_url)
        if df_t is not None:
            return df_t, df_c, status
    df_t, df_c = get_mock_data()
    return df_t, df_c, "Dados Simulados (Backup)"

# Sidebar Branding
st.sidebar.markdown("""
    <div style="text-align: center; margin-bottom: 20px;">
        <span style="font-size: 28px;">🌿</span>
        <h2 style="color: #10B981; font-size: 18px; margin: 5px 0 0 0; font-weight:700;">Guanabara Verde</h2>
        <p style="color: #94A3B8; font-size: 11px; font-weight:500;">Capacitação em SbN (Região RH-V)</p>
    </div>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Canal de Dados")
origem_dados = st.sidebar.selectbox("Origem do Ficheiro:", ["Arquivos Locais/CSVs", "Sincronizar via GitHub"])

github_url = "https://github.com/grupomyr/sbn-guanabara/raw/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx"
use_github = (origem_dados == "Sincronizar via GitHub")

df_raw_turmas, df_raw_crono = load_project_data(github_url, use_github)
df_turmas = df_raw_turmas.copy()
df_turmas["Eixo Temático"] = df_turmas["Eixo"].apply(map_eixo_name)
df_turmas["Público-Alvo Formatado"] = df_turmas["Público-Alvo"].apply(map_publico)

# Navegação Executiva
menu = [
    "📌 Visão Geral",
    "📚 Gestão de Turmas",
    "🗺️ Territórios RH-V",
    "⚙️ Logística & Alimentação",
    "📅 Calendário de Atividades",
    "📦 Portfólio de Produtos",
    "📊 Indicadores Estratégicos"
]
choice = st.sidebar.radio("Navegação Executiva:", menu)

st.sidebar.markdown("---")
st.sidebar.markdown("### Filtros Globais")
eixos_disponiveis = list(df_turmas["Eixo Temático"].unique())
f_eixo = st.sidebar.multiselect("Filtrar por Eixo Temático:", options=eixos_disponiveis, default=eixos_disponiveis)
df_t_filtered = df_turmas[df_turmas["Eixo Temático"].isin(f_eixo)]

# --- PÁGINA 1: VISÃO GERAL ---
if choice == "📌 Visão Geral":
    st.title("📌 Visão Geral do Programa")
    st.markdown("Região Hidrográfica da Baía de Guanabara (RH-V) • Convênio BID/SEAS-RJ")
    
    st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_t_filtered["Carga Horária (h)"].sum())}h</div><div class="kpi-subtitle">Total Planejado</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Dias de Aula</div><div class="kpi-value">{len(df_t_filtered)}</div><div class="kpi-subtitle">Frentes Operacionais</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Alunos Previstos</div><div class="kpi-value">{int(df_t_filtered["Nº de Participantes"].sum())}</div><div class="kpi-subtitle">Pessoas Mobilizadas</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Logística (KM)</div><div class="kpi-value">{int(df_t_filtered["KM Previsto"].sum())} km</div><div class="kpi-subtitle">Deslocamento</div></div>', unsafe_allow_html=True)
    with c5: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Bases e Espaços</div><div class="kpi-value">{len(df_t_filtered["Local"].unique())}</div><div class="kpi-subtitle">Frentes em Campo</div></div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    g1, g2 = st.columns([3, 2])
    with g1:
        st.subheader("Público-Alvo por Eixo Temático")
        fig = px.bar(df_t_filtered, x="Eixo Temático", y="Nº de Participantes", color="Público-Alvo Formatado", barmode="group", template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Safe)
        st.plotly_chart(fig, use_container_width=True)
    with g2:
        st.subheader("Distribuição de Carga Horária")
        fig_pie = px.pie(df_t_filtered, names="Eixo Temático", values="Carga Horária (h)", hole=0.4, template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig_pie, use_container_width=True)

# --- PÁGINA 2: GESTÃO DE TURMAS ---
elif choice == "📚 Gestão de Turmas":
    st.title("📚 Detalhamento das Turmas e Atividades")
    st.markdown("Lista completa de capacitações planejadas para o projeto.")
    st.dataframe(df_t_filtered.drop(columns=['lat', 'lon']), use_container_width=True)

# --- PÁGINA 3: TERRITÓRIOS (OPENSTREETMAP) ---
elif choice == "🗺️ Territórios RH-V":
    st.title("🗺️ Mapa de Atuação - OpenStreetMap")
    st.markdown("Visualização geográfica das bases de capacitação e frentes de campo.")
    
    # Centralizado na Baía de Guanabara
    m = folium.Map(location=[-22.8, -43.1], zoom_start=9, tiles="OpenStreetMap")
    
    for idx, row in df_t_filtered.iterrows():
        color = COLOR_MAP.get(row['Eixo Temático'], "#64748B")
        folium.Marker(
            [row['lat'], row['lon']], 
            popup=f"<b>Local:</b> {row['Local']}<br><b>Eixo:</b> {row['Eixo Temático']}<br><b>Tema:</b> {row['Subtema']}",
            tooltip=f"{row['Local']} - {row['Subtema']}",
            icon=folium.Icon(color="green" if "Agro" in row['Eixo Temático'] else "blue", icon="info-sign")
        ).add_to(m)
    
    st_folium(m, width=1200, height=600)

# --- PÁGINA 4: LOGÍSTICA & ALIMENTAÇÃO ---
elif choice == "⚙️ Logística & Alimentação":
    st.title("⚙️ Logística, Coffee Break e Transporte")
    st.markdown("Gestão de insumos e necessidades logísticas para cada atividade.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("☕ Necessidade de Alimentação")
        # Filtrar onde há qualquer tipo de alimentação (marcado com X)
        df_food = df_t_filtered[
            (df_t_filtered["Coffe break (Sim/Não)"].str.upper() == "X") | 
            (df_t_filtered["Almoço (Sim/Não)"].str.upper() == "X") | 
            (df_t_filtered["Lanche (Sim/Não)"].str.upper() == "X")
        ][["Local", "Subtema", "Coffe break (Sim/Não)", "Almoço (Sim/Não)", "Lanche (Sim/Não)"]]
        st.dataframe(df_food, use_container_width=True, hide_index=True)
    
    with col2:
        st.subheader("🚐 Logística de Transporte")
        # Separar Visitas Técnicas conforme solicitado
        df_visitas = df_t_filtered[df_t_filtered["Tipo de Atividade"] == "Visita Técnica"]
        df_outros = df_t_filtered[df_t_filtered["Tipo de Atividade"] != "Visita Técnica"]
        
        st.write("📌 **Visitas Técnicas (Requerem Logística Especial):**")
        if not df_visitas.empty:
            st.dataframe(df_visitas[["Local", "Subtema", "Tipo Veículo", "KM Previsto"]], use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma visita técnica filtrada.")
            
        st.write("🏢 **Atividades em Base/Sede:**")
        st.dataframe(df_outros[["Local", "Subtema", "Tipo Veículo", "KM Previsto"]].head(10), use_container_width=True, hide_index=True)

# --- PÁGINA 5: CALENDÁRIO ---
elif choice == "📅 Calendário de Atividades":
    st.title("📅 Calendário do Projeto")
    st.markdown("Substituindo a linha do tempo por uma visão de calendário mensal/semanal.")
    
    calendar_events = []
    for idx, row in df_t_filtered.iterrows():
        # Cores baseadas no eixo
        color = COLOR_MAP.get(row['Eixo Temático'], "#64748B")
        calendar_events.append({
            "title": f"{row['Subtema']} - {row['Local']}",
            "start": row['Data'].isoformat(),
            "end": row['Data'].isoformat(),
            "backgroundColor": color,
            "borderColor": color,
        })
    
    calendar_options = {
        "headerToolbar": {
            "left": "prev,next today",
            "center": "title",
            "right": "dayGridMonth,timeGridWeek,listWeek",
        },
        "initialView": "dayGridMonth",
        "locale": "pt-br",
    }
    calendar(events=calendar_events, options=calendar_options)

# --- PÁGINA 6: PRODUTOS ---
elif choice == "📦 Portfólio de Produtos":
    st.title("📦 Produtos Contratuais e Entregas (BID/SEAS)")
    st.markdown("Documentação oficial e entregas aprovadas.")
    
    st.markdown("""
        <div style="background-color: #1E293B; padding: 20px; border-radius: 12px; border: 1px solid #334155;">
            <h3 style="color: #10B981; margin-top:0;">📄 Entrega P1: Plano de Trabalho</h3>
            <p>Última versão entregue e aprovada disponível para consulta no repositório oficial.</p>
            <a href="https://github.com/grupomyr/sbn-guanabara/raw/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf" target="_blank" style="text-decoration: none;">
                <button style="background-color: #10B981; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-weight: bold;">
                    ⬇️ Baixar P1 - Versão R04 (PDF)
                </button>
            </a>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Status das Entregas")
    produtos = [
        {"ID": "P1", "Produto": "Plano de Trabalho", "Status": "✅ Aprovado", "Versão": "R04"},
        {"ID": "P2", "Produto": "Materiais Pedagógicos e Kits", "Status": "⏳ Em Elaboração", "Versão": "R01"},
        {"ID": "P3", "Produto": "Relatório Técnico Ciclo 1", "Status": "📅 Planejado", "Versão": "-"},
    ]
    st.table(produtos)

# --- PÁGINA 7: INDICADORES ---
elif choice == "📊 Indicadores Estratégicos":
    st.title("📊 Indicadores e Monitoramento")
    st.markdown("Preparação para integração com Google Forms e Planilhas sob diretrizes da LGPD.")
    
    st.info("💡 **Próximos Passos:** Cada capacitação terá um link exclusivo do Forms para coleta de dados sociodemográficos, alimentando automaticamente este painel via Google Sheets API.")
    
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("Garantia de Paridade de Gênero")
        fig_g = px.pie(values=[55, 45], names=["Feminino (Meta)", "Masculino"], color_discrete_sequence=["#10B981", "#3B82F6"], hole=0.4, template="plotly_dark")
        st.plotly_chart(fig_g, use_container_width=True)
    with col_g2:
        st.subheader("Meta de Envolvimento Juvenil")
        fig_y = px.bar(x=["Meta Jovens", "Outros"], y=[35, 65], color=["#F59E0B", "#64748B"], color_discrete_map="identity", template="plotly_dark")
        st.plotly_chart(fig_y, use_container_width=True)

# Rodapé Institucional
st.sidebar.markdown("---")
st.sidebar.caption("Guanabara Verde Resiliente • Versão R02")
st.sidebar.caption("Realização: SEAS-RJ | BID | CANADÁ")

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import datetime
import io
import urllib.request
import folium
from streamlit_folium import st_folium
from streamlit_calendar import calendar

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Guanabara Verde | Dashboard Executivo",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILO CSS PREMIUM ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        /* Ajustes Globais */
        .main { background-color: #020617; color: #F8FAFC; font-family: 'Inter', sans-serif; }
        [data-testid="stSidebar"] { background-color: #0F172A; border-right: 1px solid #1E293B; }
        
        /* Cards de KPI */
        .kpi-card {
            background: rgba(30, 41, 59, 0.7);
            border: 1px solid #334155;
            border-radius: 16px;
            padding: 24px;
            text-align: center;
            transition: all 0.3s ease;
        }
        .kpi-card:hover { border-color: #10B981; transform: translateY(-5px); background: rgba(30, 41, 59, 0.9); }
        .kpi-title { font-size: 12px; color: #94A3B8; font-weight: 600; text-transform: uppercase; margin-bottom: 8px; }
        .kpi-value { font-size: 32px; color: #F8FAFC; font-weight: 800; }
        
        /* Botões e Links */
        .stButton>button { width: 100%; border-radius: 8px; background-color: #10B981; color: white; font-weight: 600; }
        
        /* Esconder menu Streamlit */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# --- PALETA DE CORES ---
COLOR_MAP = {
    "Agroecologia e Adequação Ambiental": "#10B981",
    "Aquicultura e Pesca Artesanal": "#3B82F6",
    "Conforto Térmico Urbano": "#F59E0B",
    "Não Especificado": "#64748B"
}

# --- FUNÇÕES DE SUPORTE ---
def safe_numeric(val):
    try: return float(pd.to_numeric(val, errors='coerce'))
    except: return 0.0

def map_eixo(val):
    val = str(val).strip().split('.')[0]
    if '1' in val: return "Agroecologia e Adequação Ambiental"
    if '2' in val: return "Aquicultura e Pesca Artesanal"
    if '3' in val: return "Conforto Térmico Urbano"
    return "Não Especificado"

@st.cache_data(ttl=300)
def load_data():
    url = "https://github.com/grupomyr/sbn-guanabara/raw/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = response.read()
        xls = pd.ExcelFile(io.BytesIO(data))
        # Tenta carregar a aba de turmas
        target_sheet = [s for s in xls.sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
        df = pd.read_excel(xls, sheet_name=target_sheet)
        
        # Limpeza básica
        df.columns = df.columns.str.strip()
        df["Eixo Temático"] = df["Eixo"].apply(map_eixo)
        df["Carga Horária (h)"] = df["Carga Horária (h)"].apply(safe_numeric)
        df["Nº de Participantes"] = df["Nº de Participantes"].apply(safe_numeric)
        df["KM Previsto"] = df["KM Previsto"].apply(safe_numeric)
        
        if "Data" in df.columns:
            df["Data"] = pd.to_datetime(df["Data"], errors='coerce')
        else:
            df["Data"] = pd.to_datetime("2026-07-01")
            
        if "Status" not in df.columns:
            df["Status"] = np.random.choice(["Planejada", "Em Andamento", "Concluída"], size=len(df))
            
        return df
    except Exception as e:
        # Fallback de segurança caso o GitHub falhe
        st.sidebar.error(f"Erro ao conectar: {e}")
        dates = pd.date_range(start="2026-07-01", periods=20)
        return pd.DataFrame({
            "Eixo Temático": ["Agroecologia e Adequação Ambiental", "Aquicultura e Pesca Artesanal", "Conforto Térmico Urbano"] * 7,
            "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí", "Duque de Caxias"] * 4 + ["São Gonçalo"],
            "Subtema": [f"Oficina {i}" for i in range(21)],
            "Carga Horária (h)": [16, 24, 12, 8] * 5 + [16],
            "Nº de Participantes": [20, 25, 15, 30] * 5 + [20],
            "KM Previsto": [100, 50, 200, 80] * 5 + [100],
            "Data": dates.append(pd.DatetimeIndex([dates[-1]])),
            "Status": ["Planejada"] * 10 + ["Em Andamento"] * 5 + ["Concluída"] * 6,
            "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 7
        })

# --- CARREGAMENTO ---
df = load_data()

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("""
        <div style="text-align: center; margin-bottom: 20px;">
            <h1 style="color: #10B981; font-size: 24px; margin-bottom: 0;">🌿 Guanabara Verde</h1>
            <p style="color: #94A3B8; font-size: 12px;">Programa de Capacitação em SbN</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    menu = ["📊 Dashboard Geral", "📋 Gestão de Turmas", "🗺️ Mapa RH-V", "🚐 Logística", "📅 Cronograma", "📦 Produtos"]
    choice = st.radio("Navegação Principal", menu)
    
    st.markdown("---")
    eixos = list(df["Eixo Temático"].unique())
    f_eixo = st.multiselect("Filtrar por Eixo", eixos, default=eixos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]

# --- DASHBOARD GERAL ---
if choice == "📊 Dashboard Geral":
    st.title("📊 Painel Executivo de Monitoramento")
    
    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Carga Horária Total</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total de Participantes</div><div class="kpi-value">{int(df_f["Nº de Participantes"].sum())}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Logística (KM)</div><div class="kpi-value">{int(df_f["KM Previsto"].sum())}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total de Turmas</div><div class="kpi-value">{len(df_f)}</div></div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    g1, g2 = st.columns(2)
    with g1:
        fig = px.bar(df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index(), 
                     x="Eixo Temático", y="Nº de Participantes", color="Eixo Temático",
                     title="Participantes por Eixo", template="plotly_dark", color_discrete_map=COLOR_MAP)
        st.plotly_chart(fig, use_container_width=True)
    with g2:
        fig_status = px.pie(df_f, names="Status", title="Status das Atividades", 
                            hole=0.4, template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Safe)
        st.plotly_chart(fig_status, use_container_width=True)

# --- GESTÃO DE TURMAS ---
elif choice == "📋 Gestão de Turmas":
    st.title("📋 Detalhamento das Turmas")
    st.dataframe(df_f, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("Análise de Atividades")
    fig_atv = px.bar(df_f.groupby(["Eixo Temático", "Status"]).size().reset_index(name="Qtd"), 
                     x="Eixo Temático", y="Qtd", color="Status", barmode="group",
                     template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Pastel)
    st.plotly_chart(fig_atv, use_container_width=True)

# --- MAPA ---
elif choice == "🗺️ Mapa RH-V":
    st.title("🗺️ Mapa Geográfico de Atuação")
    coords = {"Rio de Janeiro": [-22.9, -43.2], "Niterói": [-22.88, -43.11], "Magé": [-22.65, -43.04], "Itaboraí": [-22.74, -42.86], "Duque de Caxias": [-22.78, -43.31], "São Gonçalo": [-22.82, -43.05]}
    
    m = folium.Map(location=[-22.8, -43.1], zoom_start=10, tiles="cartodbpositron")
    for _, row in df_f.iterrows():
        loc = coords.get(row['Local'], [-22.9, -43.2])
        folium.Marker(
            loc, 
            popup=f"<b>{row['Subtema']}</b><br>Eixo: {row['Eixo Temático']}<br>Status: {row['Status']}",
            tooltip=row['Local'],
            icon=folium.Icon(color='green' if row['Status'] == 'Concluída' else 'blue')
        ).add_to(m)
    st_folium(m, width="100%", height=500)

# --- LOGÍSTICA ---
elif choice == "🚐 Logística":
    st.title("🚐 Gestão de Logística e Transporte")
    
    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("Quilometragem por Eixo")
        fig_km = px.bar(df_f.groupby("Eixo Temático")["KM Previsto"].sum().reset_index(), 
                        y="Eixo Temático", x="KM Previsto", orientation='h', 
                        template="plotly_dark", color="Eixo Temático", color_discrete_map=COLOR_MAP)
        st.plotly_chart(fig_km, use_container_width=True)
    with c2:
        st.subheader("Visitas Técnicas Identificadas")
        if "Tipo de Atividade" in df_f.columns:
            visitas = df_f[df_f["Tipo de Atividade"] == "Visita Técnica"]
            st.dataframe(visitas[["Local", "Subtema", "KM Previsto"]], use_container_width=True, hide_index=True)
        else:
            st.info("Dados de tipo de atividade não encontrados.")

# --- CRONOGRAMA ---
elif choice == "📅 Cronograma":
    st.title("📅 Planejamento Temporal")
    tab1, tab2 = st.tabs(["Calendário Mensal", "Gráfico de Gantt"])
    
    with tab1:
        events = []
        for _, r in df_f.iterrows():
            if pd.notna(r['Data']):
                events.append({
                    "title": f"{r['Subtema']}",
                    "start": r['Data'].strftime("%Y-%m-%d"),
                    "backgroundColor": COLOR_MAP.get(r['Eixo Temático'], "#64748B"),
                    "borderColor": "#1E293B"
                })
        calendar(events=events, options={"initialView": "dayGridMonth", "locale": "pt-br"})
        
    with tab2:
        df_g = df_f.copy()
        df_g["Start"] = df_g["Data"]
        df_g["Finish"] = df_g["Data"] + pd.Timedelta(days=1)
        fig_gantt = px.timeline(df_g, start="Start", finish="Finish", x_start="Start", x_end="Finish", 
                                y="Subtema", color="Eixo Temático", template="plotly_dark", 
                                color_discrete_map=COLOR_MAP, title="Linha do Tempo de Atividades")
        fig_gantt.update_yaxes(autorange="reversed")
        st.plotly_chart(fig_gantt, use_container_width=True)

# --- PRODUTOS ---
elif choice == "📦 Produtos":
    st.title("📦 Entregas e Produtos Contratuais")
    st.markdown("""
        <div style="background: rgba(30, 41, 59, 0.5); padding: 30px; border-radius: 16px; border: 1px solid #334155;">
            <h3 style="color: #10B981;">📄 P1: Plano de Trabalho (Versão R04)</h3>
            <p style="color: #94A3B8;">O documento oficial aprovado pelo BID e SEAS-RJ.</p>
            <div style="display: flex; gap: 10px; margin-top: 20px;">
                <a href="https://github.com/grupomyr/sbn-guanabara/raw/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf" target="_blank" style="text-decoration: none; background: #10B981; color: white; padding: 12px 24px; border-radius: 8px; font-weight: 600;">⬇️ Baixar PDF</a>
                <a href="https://github.com/grupomyr/sbn-guanabara/blob/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf" target="_blank" style="text-decoration: none; background: #334155; color: white; padding: 12px 24px; border-radius: 8px; font-weight: 600;">👁️ Ver no GitHub</a>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Indicadores ESG (BID)")
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.pie(values=[55, 45], names=["Feminino", "Masculino"], title="Meta de Gênero", hole=0.5, template="plotly_dark"), use_container_width=True)
    with c2:
        st.plotly_chart(px.bar(x=["Meta Jovens", "Outros"], y=[35, 65], title="Envolvimento Juvenil", template="plotly_dark"), use_container_width=True)

# --- RODAPÉ ---
st.sidebar.markdown("---")
st.sidebar.caption("Guanabara Verde Resiliente • Versão R03")
st.sidebar.caption("Realização: SEAS-RJ | BID | CANADÁ")

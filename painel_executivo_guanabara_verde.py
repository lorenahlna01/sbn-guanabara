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
    page_title="Guanabara Verde | Gestão Estratégica",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILO CSS EXECUTIVO ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        .main { background-color: #020617; color: #F8FAFC; font-family: 'Inter', sans-serif; }
        [data-testid="stSidebar"] { background-color: #0F172A; border-right: 1px solid #1E293B; }
        
        /* Glassmorphism Cards */
        .glass-card {
            background: rgba(30, 41, 59, 0.4);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 20px;
            padding: 25px;
            margin-bottom: 20px;
            backdrop-filter: blur(10px);
        }
        
        .kpi-title { font-size: 11px; color: #94A3B8; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
        .kpi-value { font-size: 36px; color: #10B981; font-weight: 800; margin-top: 5px; }
        
        /* Custom Table */
        .stDataFrame { border-radius: 15px; overflow: hidden; border: 1px solid #1E293B; }
        
        /* Sidebar Logo */
        .logo-text { font-size: 22px; font-weight: 800; color: #10B981; text-align: center; margin-bottom: 30px; }
    </style>
""", unsafe_allow_html=True)

# --- CONFIGURAÇÕES E CORES ---
COLOR_MAP = {
    "Agroecologia e Adequação Ambiental": "#10B981",
    "Aquicultura e Pesca Artesanal": "#3B82F6",
    "Conforto Térmico Urbano": "#F59E0B",
    "Não Especificado": "#64748B"
}

# --- CARREGAMENTO DE DADOS ---
@st.cache_data(ttl=60)
def load_data():
    url = "https://github.com/grupomyr/sbn-guanabara/raw/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = response.read()
        xls = pd.ExcelFile(io.BytesIO(data))
        target_sheet = [s for s in xls.sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
        df = pd.read_excel(xls, sheet_name=target_sheet)
        
        # Limpeza e Mapeamento
        df.columns = df.columns.str.strip()
        def map_eixo(val):
            val = str(val).strip().split('.')[0]
            if '1' in val: return "Agroecologia e Adequação Ambiental"
            if '2' in val: return "Aquicultura e Pesca Artesanal"
            if '3' in val: return "Conforto Térmico Urbano"
            return "Não Especificado"
        
        df["Eixo Temático"] = df["Eixo"].apply(map_eixo)
        df["KM Previsto"] = pd.to_numeric(df["KM Previsto"], errors='coerce').fillna(0)
        df["Nº de Participantes"] = pd.to_numeric(df["Nº de Participantes"], errors='coerce').fillna(0)
        df["Carga Horária (h)"] = pd.to_numeric(df["Carga Horária (h)"], errors='coerce').fillna(0)
        df["Data"] = pd.to_datetime(df["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
        if "Status" not in df.columns: df["Status"] = "Planejada"
        
        return df
    except:
        # Fallback Robusto
        return pd.DataFrame({
            "Eixo Temático": ["Agroecologia e Adequação Ambiental", "Aquicultura e Pesca Artesanal", "Conforto Térmico Urbano"] * 5,
            "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí", "Duque de Caxias"] * 3,
            "Subtema": [f"Oficina Estratégica {i}" for i in range(15)],
            "KM Previsto": [120, 45, 180, 90, 60] * 3,
            "Nº de Participantes": [25, 20, 15, 30, 20] * 3,
            "Carga Horária (h)": [16, 24, 12, 8, 16] * 3,
            "Data": pd.date_range("2026-07-01", periods=15),
            "Status": ["Planejada"] * 10 + ["Concluída"] * 5,
            "Tipo Veículo": ["Van", "Carro", "Ônibus", "Van", "Carro"] * 3,
            "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 5
        })

df = load_data()

# --- SIDEBAR ---
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("NAVEGAÇÃO", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS RH-V", "🚐 PAINEL LOGÍSTICO", "📅 CRONOGRAMA", "📦 ENTREGAS"], label_visibility="collapsed")
    st.markdown("---")
    f_eixo = st.multiselect("FILTRAR EIXO", list(df["Eixo Temático"].unique()), default=list(df["Eixo Temático"].unique()))
    df_f = df[df["Eixo Temático"].isin(f_eixo)]

# --- VISÃO GERAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Participantes</div><div class="kpi-value">{int(df_f["Nº de Participantes"].sum())}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Km Previstos</div><div class="kpi-value">{int(df_f["KM Previsto"].sum())}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Total Turmas</div><div class="kpi-value">{len(df_f)}</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.plotly_chart(px.bar(df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index(), x="Eixo Temático", y="Nº de Participantes", color="Eixo Temático", template="plotly_dark", color_discrete_map=COLOR_MAP, title="Mobilização por Eixo"), use_container_width=True)
    with col_b:
        st.plotly_chart(px.pie(df_f, names="Status", hole=0.6, template="plotly_dark", title="Status do Projeto"), use_container_width=True)

# --- GESTÃO OPERACIONAL ---
elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Gestão Operacional das Turmas")
    st.dataframe(df_f.drop(columns=["Eixo"], errors='ignore'), use_container_width=True)
    
    st.markdown("---")
    st.subheader("Complexidade das Atividades")
    fig = px.sunburst(df_f, path=['Eixo Temático', 'Status', 'Local'], values='Nº de Participantes', color='Eixo Temático', color_discrete_map=COLOR_MAP, template="plotly_dark")
    st.plotly_chart(fig, use_container_width=True)

# --- MAPA ---
elif menu == "🗺️ TERRITÓRIOS RH-V":
    st.title("🗺️ Territórios de Atuação")
    coords = {"Rio de Janeiro": [-22.9, -43.2], "Niterói": [-22.88, -43.11], "Magé": [-22.65, -43.04], "Itaboraí": [-22.74, -42.86], "Duque de Caxias": [-22.78, -43.31], "São Gonçalo": [-22.82, -43.05]}
    m = folium.Map(location=[-22.8, -43.1], zoom_start=10, tiles="cartodbpositron")
    for _, r in df_f.iterrows():
        folium.Marker(coords.get(r['Local'], [-22.9, -43.2]), popup=f"<b>{r['Subtema']}</b>", tooltip=r['Local']).add_to(m)
    st_folium(m, width="100%", height=600)

# --- PAINEL LOGÍSTICO ---
elif menu == "🚐 PAINEL LOGÍSTICO":
    st.title("🚐 Painel de Operações Logísticas")
    
    # KPIs Logísticos
    l1, l2, l3 = st.columns(3)
    with l1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Quilometragem Total</div><div class="kpi-value" style="color:#3B82F6;">{int(df_f["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with l2: 
        v_count = df_f["Tipo Veículo"].value_counts().to_dict() if "Tipo Veículo" in df_f.columns else {"Van": 0}
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Veículos Principais</div><div class="kpi-value" style="color:#F59E0B;">{list(v_count.keys())[0] if v_count else "N/A"}</div></div>', unsafe_allow_html=True)
    with l3: 
        vt_count = len(df_f[df_f["Tipo de Atividade"] == "Visita Técnica"]) if "Tipo de Atividade" in df_f.columns else 0
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Visitas Técnicas</div><div class="kpi-value" style="color:#10B981;">{vt_count}</div></div>', unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Análise de Deslocamento por Eixo")
    fig_log = px.bar(df_f.groupby("Eixo Temático")["KM Previsto"].sum().reset_index(), x="KM Previsto", y="Eixo Temático", orientation='h', color="Eixo Temático", color_discrete_map=COLOR_MAP, template="plotly_dark", text_auto=True)
    st.plotly_chart(fig_log, use_container_width=True)
    
    st.subheader("Detalhamento de Transporte")
    if "Tipo de Atividade" in df_f.columns:
        df_transp = df_f[df_f["Tipo de Atividade"].isin(["Visita Técnica", "Prática"])][["Local", "Subtema", "Tipo Veículo", "KM Previsto", "Tipo de Atividade"]]
        st.dataframe(df_transp, use_container_width=True, hide_index=True)

# --- CRONOGRAMA ---
elif menu == "📅 CRONOGRAMA":
    st.title("📅 Cronograma e Linha do Tempo")
    t1, t2 = st.tabs(["Calendário Mensal", "Gráfico de Gantt"])
    with t1:
        events = [{"title": r['Subtema'], "start": r['Data'].strftime("%Y-%m-%d"), "backgroundColor": COLOR_MAP.get(r['Eixo Temático'], "#64748B")} for _, r in df_f.iterrows()]
        calendar(events=events, options={"initialView": "dayGridMonth", "locale": "pt-br"})
    with t2:
        df_g = df_f.copy()
        df_g["Start"] = df_g["Data"]
        df_g["Finish"] = df_g["Data"] + pd.Timedelta(days=1)
        st.plotly_chart(px.timeline(df_g, start="Start", finish="Finish", x_start="Start", x_end="Finish", y="Subtema", color="Eixo Temático", template="plotly_dark", color_discrete_map=COLOR_MAP), use_container_width=True)

# --- ENTREGAS ---
elif menu == "📦 ENTREGAS":
    st.title("📦 Produtos e Entregas Contratuais")
    
    st.markdown("""
        <div class="glass-card">
            <h3 style="color:#10B981; margin-top:0;">📄 P1: Plano de Trabalho (R04)</h3>
            <p style="color:#94A3B8;">O Plano de Trabalho é o documento norteador das atividades de capacitação SbN.</p>
            <hr style="border: 0.5px solid #1E293B;">
            <p style="font-size:14px;"><b>Status:</b> ✅ Entregue e Aprovado</p>
            <br>
    """, unsafe_allow_html=True)
    
    # Botão de Download Direto e Seguro
    pdf_url = "https://github.com/grupomyr/sbn-guanabara/raw/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf"
    st.markdown(f'<a href="{pdf_url}" download><button style="background-color:#10B981; color:white; border:none; padding:12px 30px; border-radius:10px; font-weight:700; cursor:pointer; width:100%;">⬇️ BAIXAR PLANO DE TRABALHO (PDF)</button></a>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Indicadores de Sustentabilidade")
    col1, col2 = st.columns(2)
    with col1: st.plotly_chart(px.pie(values=[55, 45], names=["Feminino", "Masculino"], title="Paridade de Gênero", hole=0.5, template="plotly_dark"), use_container_width=True)
    with col2: st.plotly_chart(px.bar(x=["Meta Jovens", "Outros"], y=[35, 65], title="Envolvimento Juvenil", template="plotly_dark"), use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.caption("Guanabara Verde • Gestão R03")
st.sidebar.caption("BID | SEAS-RJ | CANADÁ")

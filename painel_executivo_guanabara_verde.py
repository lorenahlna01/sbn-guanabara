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

# Estilização CSS personalizada
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        .main { background-color: #020617; color: #F8FAFC; font-family: 'Inter', sans-serif; }
        [data-testid="stSidebar"] { background-color: #0F172A; border-right: 1px solid #1E293B; }
        .kpi-container { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
        .kpi-card { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid #334155; border-radius: 12px; padding: 20px; text-align: left; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); transition: all 0.3s ease; }
        .kpi-card:hover { transform: translateY(-3px); border-color: #10B981; box-shadow: 0 10px 15px -3px rgba(16, 185, 129, 0.1); }
        .kpi-title { font-size: 11px; color: #94A3B8; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 6px; }
        .kpi-value { font-size: 26px; color: #F8FAFC; font-weight: 700; line-height: 1.1; }
        .kpi-subtitle { font-size: 11px; color: #10B981; margin-top: 4px; font-weight: 500; }
    </style>
""", unsafe_allow_html=True)

COLOR_MAP = {
    "Agroecologia e Adequação Ambiental": "#10B981",
    "Aquicultura e Pesca Artesanal": "#3B82F6",
    "Conforto Térmico Urbano": "#F59E0B",
    "Não Especificado": "#64748B"
}

def map_eixo_name(val):
    val_str = str(val).strip().split('.')[0]
    if '1' in val_str: return "Agroecologia e Adequação Ambiental"
    elif '2' in val_str: return "Aquicultura e Pesca Artesanal"
    elif '3' in val_str: return "Conforto Térmico Urbano"
    return "Não Especificado"

def map_publico(val):
    val_str = str(val).strip()
    if val_str == "1": return "Técnicos Municipais"
    elif val_str == "2": return "Lideranças Comunitárias"
    elif val_str in ["1.2", "1,2"]: return "Misto"
    return f"Grupo {val_str}"

@st.cache_data(ttl=300)
def fetch_excel_from_github(url):
    try:
        # Adicionando headers para evitar bloqueios de segurança
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        data_bytes = urllib.request.urlopen(req).read()
        xls = pd.ExcelFile(io.BytesIO(data_bytes))
        sheet_names = xls.sheet_names
        sheet_turmas = [s for s in sheet_names if "turma" in s.lower() or "gest" in s.lower()]
        if not sheet_turmas: return None
        df_t = pd.read_excel(xls, sheet_name=sheet_turmas[0])
        return df_t
    except Exception as e:
        return None

def get_mock_data():
    df_t = pd.DataFrame({
        "Eixo": [1, 2, 3] * 10,
        "Público-Alvo": ["1", "2", "1.2"] * 10,
        "Subtema": ["Adequação Ambiental", "Aquicultura Urbana", "Conforto Térmico Urbano"] * 10,
        "Turma": [1] * 30,
        "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí", "Duque de Caxias", "São Gonçalo"] * 5,
        "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 10,
        "Carga Horária (h)": [16, 24, 12] * 10,
        "Nº de Participantes": [20, 20, 20] * 10,
        "KM Previsto": [120, 50, 200] * 10,
        "Tipo Veículo": ["Van", "Carro", "Ônibus"] * 10,
        "Data": [datetime.date(2026, 7, 1) + datetime.timedelta(days=i) for i in range(30)],
        "Status": ["Concluída", "Planejada", "Em Andamento"] * 10
    })
    return df_t

# --- CARREGAMENTO ---
GITHUB_XLSX_URL = "https://github.com/grupomyr/sbn-guanabara/raw/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx"
df_raw = fetch_excel_from_github(GITHUB_XLSX_URL)
if df_raw is None:
    st.sidebar.warning("⚠️ Usando dados simulados (Verifique o link do GitHub)")
    df_raw = get_mock_data()

df_turmas = df_raw.copy()
df_turmas["Eixo Temático"] = df_turmas["Eixo"].apply(map_eixo_name)
df_turmas["Público-Alvo Formatado"] = df_turmas["Público-Alvo"].apply(map_publico)
if "Status" not in df_turmas.columns: df_turmas["Status"] = "Planejada"

# Coordenadas
coords = {
    "Rio de Janeiro": [-22.9068, -43.1729], "Niterói": [-22.8833, -43.1167],
    "Magé": [-22.6528, -43.0403], "Itaboraí": [-22.7444, -42.8594],
    "Duque de Caxias": [-22.7856, -43.3117], "São Gonçalo": [-22.8269, -43.0539]
}
df_turmas["lat"] = df_turmas["Local"].map(lambda x: coords.get(x, [-22.9068, -43.1729])[0])
df_turmas["lon"] = df_turmas["Local"].map(lambda x: coords.get(x, [-22.9068, -43.1729])[1])

# Sidebar
st.sidebar.markdown('<div style="text-align: center;"><span style="font-size: 28px;">🌿</span><h2 style="color: #10B981; font-size: 18px; font-weight:700;">Guanabara Verde</h2></div>', unsafe_allow_html=True)
st.sidebar.markdown("---")
menu = ["📌 Visão Geral", "📚 Gestão de Turmas", "🗺️ Territórios RH-V", "⚙️ Logística & Transporte", "📅 Cronograma & Gantt", "📦 Portfólio de Produtos", "📊 Indicadores"]
choice = st.sidebar.radio("Navegação:", menu)
st.sidebar.markdown("---")
eixos = list(df_turmas["Eixo Temático"].unique())
f_eixo = st.sidebar.multiselect("Filtrar Eixo:", eixos, default=eixos)
df_filtered = df_turmas[df_turmas["Eixo Temático"].isin(f_eixo)]

# --- PÁGINAS ---
if choice == "📌 Visão Geral":
    st.title("📌 Visão Geral do Programa")
    st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_filtered["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Alunos</div><div class="kpi-value">{int(df_filtered["Nº de Participantes"].sum())}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Logística</div><div class="kpi-value">{int(df_filtered["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><div class="kpi-title">Turmas</div><div class="kpi-value">{len(df_filtered)}</div></div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.plotly_chart(px.bar(df_filtered, x="Eixo Temático", y="Nº de Participantes", color="Público-Alvo Formatado", barmode="group", template="plotly_dark"), use_container_width=True)

elif choice == "📚 Gestão de Turmas":
    st.title("📚 Gestão de Turmas")
    st.dataframe(df_filtered.drop(columns=['lat', 'lon'], errors='ignore'), use_container_width=True)
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Atividades por Eixo")
        fig1 = px.bar(df_filtered.groupby("Eixo Temático").size().reset_name(name="Qtd"), x="Eixo Temático", y="Qtd", template="plotly_dark", color="Eixo Temático", color_discrete_map=COLOR_MAP)
        st.plotly_chart(fig1, use_container_width=True)
    with col2:
        st.subheader("Status de Conclusão")
        fig2 = px.pie(df_filtered, names="Status", template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig2, use_container_width=True)

elif choice == "🗺️ Territórios RH-V":
    st.title("🗺️ Mapa de Atuação")
    m = folium.Map(location=[-22.8, -43.1], zoom_start=9, tiles="OpenStreetMap")
    for _, row in df_filtered.iterrows():
        folium.Marker([row['lat'], row['lon']], popup=f"<b>{row['Local']}</b><br>{row['Subtema']}<br>{row['Eixo Temático']}", tooltip=row['Local']).add_to(m)
    st_folium(m, width=1200, height=500)

elif choice == "⚙️ Logística & Transporte":
    st.title("🚐 Logística & Transporte")
    df_visitas = df_filtered[df_filtered["Tipo de Atividade"] == "Visita Técnica"]
    st.subheader("📌 Visitas Técnicas")
    st.dataframe(df_visitas[["Local", "Subtema", "Tipo Veículo", "KM Previsto"]], use_container_width=True)
    st.subheader("📊 Quilometragem por Eixo")
    fig_log = px.bar(df_filtered.groupby("Eixo Temático")["KM Previsto"].sum().reset_index(), y="Eixo Temático", x="KM Previsto", orientation='h', template="plotly_dark", color="Eixo Temático", color_discrete_map=COLOR_MAP)
    st.plotly_chart(fig_log, use_container_width=True)

elif choice == "📅 Cronograma & Gantt":
    st.title("📅 Cronograma & Gantt")
    tab1, tab2 = st.tabs(["Calendário", "Gráfico de Gantt"])
    with tab1:
        events = [{"title": f"{r['Subtema']} ({r['Local']})", "start": r['Data'].isoformat(), "backgroundColor": COLOR_MAP.get(r['Eixo Temático'], "#64748B")} for _, r in df_filtered.iterrows() if pd.notna(r['Data'])]
        calendar(events=events, options={"initialView": "dayGridMonth", "locale": "pt-br"})
    with tab2:
        if "Data" in df_filtered.columns:
            df_gantt = df_filtered.copy()
            df_gantt["Start"] = pd.to_datetime(df_gantt["Data"])
            df_gantt["Finish"] = df_gantt["Start"] + pd.Timedelta(days=1)
            fig_gantt = px.timeline(df_gantt, start="Start", finish="Finish", x_start="Start", x_end="Finish", y="Subtema", color="Eixo Temático", template="plotly_dark", color_discrete_map=COLOR_MAP)
            fig_gantt.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_gantt, use_container_width=True)

elif choice == "📦 Portfólio de Produtos":
    st.title("📦 Portfólio de Produtos")
    st.markdown("""
        <div style="background-color: #1E293B; padding: 20px; border-radius: 12px; border: 1px solid #334155;">
            <h3 style="color: #10B981; margin-top:0;">📄 Entrega P1: Plano de Trabalho</h3>
            <p>Clique abaixo para acessar o documento oficial no GitHub:</p>
            <a href="https://github.com/grupomyr/sbn-guanabara/blob/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf" target="_blank">
                <button style="background-color: #10B981; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-weight: bold;">
                    👁️ Ver P1 - Versão R04 (GitHub)
                </button>
            </a>
            <br><br>
            <iframe src="https://github.com/grupomyr/sbn-guanabara/raw/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf" width="100%" height="600px"></iframe>
        </div>
    """, unsafe_allow_html=True)

elif choice == "📊 Indicadores":
    st.title("📊 Indicadores")
    st.plotly_chart(px.pie(values=[55, 45], names=["Feminino", "Masculino"], title="Paridade de Gênero", hole=0.4, template="plotly_dark"))

st.sidebar.markdown("---")
st.sidebar.caption("Guanabara Verde • Versão R02")

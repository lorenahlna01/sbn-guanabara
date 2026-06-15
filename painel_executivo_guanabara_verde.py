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
import os

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
        .logo-text { font-size: 22px; font-weight: 800; color: #10B981; text-align: center; margin-bottom: 30px; }
        
        /* Ajuste para imagem na sidebar */
        .sidebar-footer {
            position: fixed;
            bottom: 20px;
            left: 20px;
            width: 260px;
        }
    </style>
""", unsafe_allow_html=True)

# --- CONFIGURAÇÕES E CORES ---
COLOR_MAP = {
    "Agroecologia e Adequação Ambiental": "#10B981",
    "Aquicultura e Pesca Artesanal": "#3B82F6",
    "Conforto Térmico Urbano": "#F59E0B",
    "Não Especificado": "#64748B"
}

# --- CARREGAMENTO DE DADOS (GOOGLE SHEETS) ---
@st.cache_data(ttl=60)
def load_data():
    sheet_id = "1vRtjdB6OY9Rei_gh-0b8TR0SJWZNv0GWL2uD0_0refkI5HU7mJCXVKEARgkNGbOvw"
    csv_url = f"https://docs.google.com/spreadsheets/d/e/2PACX-{sheet_id}/pub?output=csv"
    
    try:
        df = pd.read_csv(csv_url)
        df.columns = df.columns.str.strip()
        
        def map_eixo(val):
            val = str(val).strip().split('.')[0]
            if '1' in val: return "Agroecologia e Adequação Ambiental"
            if '2' in val: return "Aquicultura e Pesca Artesanal"
            if '3' in val: return "Conforto Térmico Urbano"
            return "Não Especificado"
        
        df["Eixo Temático"] = df["Eixo"].apply(map_eixo) if "Eixo" in df.columns else "Não Especificado"
        
        for col in ["KM Previsto", "Nº de Participantes", "Carga Horária (h)"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        if "Data" in df.columns:
            df["Data"] = pd.to_datetime(df["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
        else:
            df["Data"] = pd.to_datetime("2026-07-01")
            
        if "Status" not in df.columns: df["Status"] = "Planejada"
        
        return df
    except:
        return pd.DataFrame({"Eixo Temático": ["Erro"], "Local": ["Erro"], "Subtema": ["Erro"], "KM Previsto": [0], "Nº de Participantes": [0], "Carga Horária (h)": [0], "Data": [pd.to_datetime("2026-07-01")], "Status": ["Erro"]})

df = load_data()

# --- SIDEBAR ---
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("NAVEGAÇÃO", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS RH-V", "🚐 PAINEL LOGÍSTICO", "📅 CRONOGRAMA", "📦 ENTREGAS"], label_visibility="collapsed")
    st.markdown("---")
    eixos_unicos = list(df["Eixo Temático"].unique())
    f_eixo = st.multiselect("FILTRAR EIXO", eixos_unicos, default=eixos_unicos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]
    
    if st.button("🔄 Atualizar Dados"):
        st.cache_data.clear()
        st.rerun()
    
    # --- RODAPÉ COM RÉGUA DE LOGOS ---
    st.markdown("---")
    logo_path = "image_93c707.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        # Tenta carregar do GitHub se não estiver local
        github_logo = "https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/image_93c707.png"
        try:
            st.image(github_logo, use_container_width=True)
        except:
            st.caption("Realização: SEAS-RJ | BID | CANADÁ")

# --- CONTEÚDO PRINCIPAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Painel de Monitoramento SbN")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Participantes</div><div class="kpi-value">{int(df_f["Nº de Participantes"].sum())}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Km Logística</div><div class="kpi-value">{int(df_f["KM Previsto"].sum())}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Turmas</div><div class="kpi-value">{len(df_f)}</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a: st.plotly_chart(px.bar(df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index(), x="Eixo Temático", y="Nº de Participantes", color="Eixo Temático", template="plotly_dark", color_discrete_map=COLOR_MAP, title="Participantes por Eixo"), use_container_width=True)
    with col_b: st.plotly_chart(px.pie(df_f, names="Status", hole=0.6, template="plotly_dark", title="Status das Atividades"), use_container_width=True)

elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Detalhamento das Capacitações")
    st.dataframe(df_f.drop(columns=["Eixo"], errors='ignore'), use_container_width=True, hide_index=True)
    st.plotly_chart(px.bar(df_f.groupby("Local").size().reset_index(name="Qtd"), x="Local", y="Qtd", template="plotly_dark", color_discrete_sequence=["#10B981"], title="Distribuição Geográfica"), use_container_width=True)

elif menu == "🗺️ TERRITÓRIOS RH-V":
    st.title("🗺️ Territórios de Atuação")
    coords = {"Rio de Janeiro": [-22.9, -43.2], "Niterói": [-22.88, -43.11], "Magé": [-22.65, -43.04], "Itaboraí": [-22.74, -42.86], "Duque de Caxias": [-22.78, -43.31], "São Gonçalo": [-22.82, -43.05]}
    m = folium.Map(location=[-22.8, -43.1], zoom_start=10, tiles="cartodbpositron")
    for _, r in df_f.iterrows():
        if r['Local'] in coords:
            folium.Marker(coords[r['Local']], popup=f"<b>{r['Subtema']}</b><br>{r['Status']}", tooltip=r['Local']).add_to(m)
    st_folium(m, width="100%", height=600)

elif menu == "🚐 PAINEL LOGÍSTICO":
    st.title("🚐 Operações Logísticas e Transporte")
    l1, l2, l3 = st.columns(3)
    with l1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Total Km</div><div class="kpi-value" style="color:#3B82F6;">{int(df_f["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with l2: 
        v_count = df_f["Tipo Veículo"].value_counts().to_dict() if "Tipo Veículo" in df_f.columns else {}
        v_main = list(v_count.keys())[0] if v_count else "N/A"
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Veículo Principal</div><div class="kpi-value" style="color:#F59E0B;">{v_main}</div></div>', unsafe_allow_html=True)
    with l3: 
        vt_count = len(df_f[df_f["Tipo de Atividade"] == "Visita Técnica"]) if "Tipo de Atividade" in df_f.columns else 0
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Visitas Técnicas</div><div class="kpi-value" style="color:#10B981;">{vt_count}</div></div>', unsafe_allow_html=True)
    st.plotly_chart(px.bar(df_f.groupby("Eixo Temático")["KM Previsto"].sum().reset_index(), x="KM Previsto", y="Eixo Temático", orientation='h', color="Eixo Temático", color_discrete_map=COLOR_MAP, template="plotly_dark", text_auto=True, title="Km por Eixo"), use_container_width=True)
    st.dataframe(df_f[["Data", "Local", "Subtema", "Tipo Veículo", "KM Previsto"]].sort_values("Data"), use_container_width=True, hide_index=True)

elif menu == "📅 CRONOGRAMA":
    st.title("📅 Planejamento Temporal")
    t1, t2 = st.tabs(["Calendário", "Gantt"])
    with t1:
        events = [{"title": f"{r['Subtema']}", "start": r['Data'].strftime("%Y-%m-%d"), "backgroundColor": COLOR_MAP.get(r['Eixo Temático'], "#64748B")} for _, r in df_f.iterrows()]
        calendar(events=events, options={"initialView": "dayGridMonth", "locale": "pt-br"})
    with t2:
        df_g = df_f.copy()
        df_g["Start"], df_g["Finish"] = df_g["Data"], df_g["Data"] + pd.Timedelta(days=1)
        st.plotly_chart(px.timeline(df_g, start="Start", finish="Finish", x_start="Start", x_end="Finish", y="Subtema", color="Eixo Temático", template="plotly_dark", color_discrete_map=COLOR_MAP, title="Linha do Tempo"), use_container_width=True)

elif menu == "📦 ENTREGAS":
    st.title("📦 Portfólio de Produtos")
    st.markdown("""<div class="glass-card"><h3 style="color:#10B981; margin-top:0;">📄 P1: Plano de Trabalho (R04)</h3><p style="color:#94A3B8;">Documento oficial aprovado.</p><hr style="border: 0.5px solid #1E293B;"><p style="font-size:14px;"><b>Status:</b> ✅ Entregue e Aprovado</p><br>""", unsafe_allow_html=True)
    pdf_url = "https://github.com/grupomyr/sbn-guanabara/raw/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf"
    st.markdown(f'<a href="{pdf_url}" target="_blank"><button style="background-color:#10B981; color:white; border:none; padding:12px 30px; border-radius:10px; font-weight:700; cursor:pointer; width:100%;">⬇️ BAIXAR PLANO DE TRABALHO (PDF)</button></a>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(px.pie(values=[55, 45], names=["Feminino", "Masculino"], title="Meta Gênero", hole=0.5, template="plotly_dark"), use_container_width=True)
    with c2: st.plotly_chart(px.bar(x=["Meta Jovens", "Outros"], y=[35, 65], title="Engajamento Juvenil", template="plotly_dark"), use_container_width=True)

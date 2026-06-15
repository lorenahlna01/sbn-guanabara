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

# --- ESTILO CSS EXECUTIVO BRASILEIRO ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        
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

# --- CARREGAMENTO INTEGRADO DE AMBAS AS ABAS (EXCEL) ---
@st.cache_data(ttl=60)
def load_data():
    # URL de publicação do arquivo Excel inteiro (.xlsx) contendo as duas abas
    url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRtjdB6OY9Rei_gh-0b8TR0SJWZNv0GWL2uD0_0refkI5HU7mJCXVKEARgkNGbOvw/pub?output=xlsx"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = response.read()
        xls = pd.ExcelFile(io.BytesIO(data))
        
        # Identificação dinâmica das abas da planilha
        sheet_names = xls.sheet_names
        target_sheet_turmas = [s for s in sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
        target_sheet_crono = [s for s in sheet_names if "cronogram" in s.lower()][0]
        
        df_t = pd.read_excel(xls, sheet_name=target_sheet_turmas)
        df_c = pd.read_excel(xls, sheet_name=target_sheet_crono, header=None)
        
        # --- Limpeza e Mapeamento da aba Gestão de Turmas ---
        df_t.columns = df_t.columns.str.strip()
        
        def map_eixo(val):
            val = str(val).strip().split('.')[0]
            if '1' in val: return "Agroecologia e Adequação Ambiental"
            if '2' in val: return "Aquicultura e Pesca Artesanal"
            if '3' in val: return "Conforto Térmico Urbano"
            return "Não Especificado"
        
        df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"
        
        # Tratamento numérico seguro de colunas
        for col in ["KM Previsto", "Nº de Participantes", "Carga Horária (h)"]:
            if col in df_t.columns:
                df_t[col] = pd.to_numeric(df_t[col], errors='coerce').fillna(0)
                
        if "Data" in df_t.columns:
            df_t["Data"] = pd.to_datetime(df_t["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
        else:
            df_t["Data"] = pd.to_datetime("2026-07-01")
            
        if "Status" not in df_t.columns: 
            df_t["Status"] = "Planejada"
            
        return df_t, df_c, "Sucesso: Conectado às Planilhas de Produção via Google Sheets!"
        
    except Exception as e:
        # Fallback Robusto de Segurança (Caso a planilha ou conexão falhem)
        df_t_fallback = pd.DataFrame({
            "Eixo Temático": ["Agroecologia e Adequação Ambiental", "Aquicultura e Pesca Artesanal", "Conforto Térmico Urbano"] * 5,
            "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí", "Duque de Caxias"] * 3,
            "Subtema": [f"Oficina Prática {i}" for i in range(15)],
            "KM Previsto": [120, 45, 180, 90, 60] * 3,
            "Nº de Participantes": [25, 20, 15, 30, 20] * 3,
            "Carga Horária (h)": [16, 24, 12, 8, 16] * 3,
            "Data": pd.date_range("2026-07-01", periods=15),
            "Status": ["Planejada"] * 10 + ["Concluída"] * 5,
            "Tipo Veículo": ["Van", "Ônibus", "Ônibus", "Van", "Carro"] * 3,
            "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 5,
            "Coffe break (Sim/Não)": ["X", "", "X"] * 5,
            "Almoço (Sim/Não)": ["X", "X", ""] * 5,
            "Lanche (Sim/Não)": ["", "X", "X"] * 5
        })
        
        df_c_fallback = pd.DataFrame({
            0: ["Status", "ID", "Atividade"],
            1: ["Aprovado", "P1", "Plano de Trabalho"],
            2: ["Aprovado", "P2", "Materiais Pedagógicos"]
        })
        
        return df_t_fallback, df_c_fallback, f"Aviso: Usando dados locais de contingência devido a uma falha de conexão ({str(e)})."

df, df_crono_raw, status_conexao = load_data()

# --- SIDEBAR ---
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("NAVEGAÇÃO", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS RH-V", "🚐 PAINEL LOGÍSTICO", "📅 CRONOGRAMA", "📦 ENTREGAS"], label_visibility="collapsed")
    st.markdown("---")
    
    # Exibe o status da conexão atual para transparência de carregamento
    st.info(status_conexao)
    
    eixos_unicos = list(df["Eixo Temático"].unique())
    f_eixo = st.sidebar.multiselect("FILTRAR EIXO", eixos_unicos, default=eixos_unicos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]
    
    if st.sidebar.button("🔄 Atualizar Planilha"):
        st.cache_data.clear()
        st.rerun()
        
    # --- RODAPÉ COM RÉGUA DE LOGOS ---
    st.markdown("---")
    logo_path = "image_93c707.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        github_logo = "https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/image_93c707.png"
        try:
            st.image(github_logo, use_container_width=True)
        except:
            st.caption("Realização: SEAS-RJ | BID | CANADÁ | INEA")

# --- VISÃO GERAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    st.markdown("Região Hidrográfica da Baia de Guanabara (RH-V) • Monitoramento Técnico e Financeiro")
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Participantes Previstos</div><div class="kpi-value">{int(df_f["Nº de Participantes"].sum())}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Km Previstos</div><div class="kpi-value">{int(df_f["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Total Turmas</div><div class="kpi-value">{len(df_f)}</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.plotly_chart(px.bar(df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index(), x="Eixo Temático", y="Nº de Participantes", color="Eixo Temático", template="plotly_dark", color_discrete_map=COLOR_MAP, title="Mobilização por Eixo Temático"), use_container_width=True)
    with col_b:
        st.plotly_chart(px.pie(df_f, names="Status", hole=0.6, template="plotly_dark", title="Status das Turmas Cadastradas"), use_container_width=True)

# --- GESTÃO OPERACIONAL ---
elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Gestão Operacional das Turmas")
    st.markdown("Acompanhamento e listagem das turmas de capacitação.")
    
    # Filtro dinâmico para garantir que colunas que não existem na planilha do Excel não quebrem a aplicação (Garante imunidade a KeyError)
    colunas_desejadas = ["Eixo Temático", "Público-Alvo", "Subtema", "Turma", "Dia", "Local", "Tipo de Atividade", "Carga Horária (h)", "Nº de Participantes", "Responsável", "Observações"]
    colunas_disponiveis = [col for col in colunas_desejadas if col in df_f.columns]
    
    st.dataframe(df_f[colunas_disponiveis], use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("Complexidade das Atividades por Território")
    
    # Sunburst simplificado sem caminhos excessivos para evitar colisões visuais
    fig_sun = px.sunburst(df_f, path=['Eixo Temático', 'Local'], values='Nº de Participantes', color='Eixo Temático', color_discrete_map=COLOR_MAP, template="plotly_dark")
    st.plotly_chart(fig_sun, use_container_width=True)

# --- MAPA (TERRITÓRIOS) ---
elif menu == "🗺️ TERRITÓRIOS RH-V":
    st.title("🗺️ Territórios de Atuação")
    st.markdown("Localização espacial das atividades na Bacia de Guanabara (OSM Nativo).")
    
    coords = {
        "Rio de Janeiro": [-22.9068, -43.1729],
        "Niterói": [-22.8858, -43.1153],
        "Magé": [-22.6514, -43.0401],
        "Itaboraí": [-22.7486, -42.8594],
        "Duque de Caxias": [-22.7856, -43.3115],
        "São Gonçalo": [-22.8269, -43.0539],
        "Cachoeiras de Macacu": [-22.4633, -42.6542],
        "Seropédica": [-22.7441, -43.7121],
        "Local a definir": [-22.7533, -43.4474]
    }
    
    # Criação do mapa folium com base nativa livre e limpa do OpenStreetMap (OSM)
    m = folium.Map(location=[-22.8, -43.1], zoom_start=10, tiles="OpenStreetMap")
    for _, r in df_f.iterrows():
        coord_loc = coords.get(r['Local'], [-22.9, -43.2])
        folium.Marker(coord_loc, popup=f"<b>{r['Subtema']}</b><br>{r['Local']}", tooltip=r['Local']).add_to(m)
    st_folium(m, width="100%", height=600)

# --- PAINEL LOGÍSTICO ---
elif menu == "🚐 PAINEL LOGÍSTICO":
    st.title("🚐 Painel de Operações Logísticas")
    st.markdown("Planejamento de transportes, deslocamento de pessoal e alimentação.")
    
    # Sanitização segura e contagem para Coffee breaks, Almoços e Lanches
    def obter_sim_count(series):
        return series.astype(str).str.upper().str.contains("X|S|SIM").sum()
        
    c_break = obter_sim_count(df_f["Coffe break (Sim/Não)"]) if "Coffe break (Sim/Não)" in df_f.columns else 0
    almoco = obter_sim_count(df_f["Almoço (Sim/Não)"]) if "Almoço (Sim/Não)" in df_f.columns else 0
    lanche = obter_sim_count(df_f["Lanche (Sim/Não)"]) if "Lanche (Sim/Não)" in df_f.columns else 0
    
    # Linha de KPIs Logísticos refinados
    l1, l2, l3, l4 = st.columns(4)
    with l1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Quilometragem Total</div><div class="kpi-value" style="color:#3B82F6;">{int(df_f["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with l2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Coffee Breaks Planejados</div><div class="kpi-value" style="color:#F59E0B;">{c_break}</div></div>', unsafe_allow_html=True)
    with l3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Almoços Requeridos</div><div class="kpi-value" style="color:#10B981;">{almoco}</div></div>', unsafe_allow_html=True)
    with l4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Lanches Solicitados</div><div class="kpi-value" style="color:#8B5CF6;">{lanche}</div></div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Visitas Técnicas de Campo Separadas e Detalhadas
    st.subheader("🚐 Visitas Técnicas e Atividades Práticas de Campo")
    if "Tipo de Atividade" in df_f.columns:
        df_visitas = df_f[df_f["Tipo de Atividade"].str.contains("visita|campo|prática", case=False, na=False)]
        col_visitas_dispo = [col for col in ["Data", "Local", "Subtema", "Tipo Veículo", "KM Previsto"] if col in df_visitas.columns]
        st.dataframe(df_visitas[col_visitas_dispo].sort_values("Data"), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma Visita Técnica catalogada no conjunto de dados ativo.")

# --- CRONOGRAMA ---
elif menu == "📅 CRONOGRAMA":
    st.title("📅 Cronograma e Linha do Tempo")
    st.markdown("Marcos contratuais, atividades semanais e planejamento pedagógico.")
    
    t1, t2 = st.tabs(["Planilha Calendário (Aba 1)", "Agenda de Atividades"])
    
    with t1:
        # Tratamento e exibição da Aba 1 do Cronograma (Células de Atividade mescladas)
        df_c_clean = df_crono_raw.copy()
        months = list(df_c_clean.iloc[0])
        current_month = ""
        for i in range(3, len(months)):
            if pd.notna(months[i]) and str(months[i]).strip() != "":
                current_month = str(months[i]).strip()
            months[i] = current_month
            
        weeks = list(df_c_clean.iloc[1])

        headers = []
        for i, (m, w) in enumerate(zip(months, weeks)):
            m_s, w_s = str(m).strip(), str(w).strip()
            if i >= 3 and m_s and w_s:
                headers.append(f"{m_s} - {w_s}")
            else:
                headers.append(w_s if w_s and w_s != "nan" else (m_s if m_s != "nan" else f"Col_{i}"))
                
        df_c_clean.columns = headers
        df_c_clean = df_c_clean.iloc[2:].copy()
        df_c_clean.rename(columns={df_c_clean.columns[0]: "Status", df_c_clean.columns[1]: "ID", df_c_clean.columns[2]: "Atividade"}, inplace=True)
        df_c_clean = df_c_clean[df_c_clean["Atividade"].notna() & (df_c_clean["Atividade"].str.strip() != "")]
        
        st.dataframe(df_c_clean, use_container_width=True, hide_index=True)
        
    with t2:
        # Agenda mensal visual de atividades cadastradas nas turmas
        events = [{"title": r['Subtema'], "start": r['Data'].strftime("%Y-%m-%d"), "backgroundColor": COLOR_MAP.get(r['Eixo Temático'], "#64748B")} for _, r in df_f.iterrows()]
        calendar(events=events, options={"initialView": "dayGridMonth", "locale": "pt-br"})

# --- ENTREGAS ---
elif menu == "📦 ENTREGAS":
    st.title("📦 Produtos e Entregas Contratuais")
    st.markdown("Verificação do progresso físico dos produtos contratados pelo BID/SEAS-RJ.")
    
    st.markdown("""
        <div class="glass-card">
            <h3 style="color:#10B981; margin-top:0;">📄 P1: Plano de Trabalho (Versão R04)</h3>
            <p style="color:#94A3B8;">O Plano de Trabalho norteia a metodologia de campo e a agenda dos ciclos de capacitações do programa de SbN.</p>
            <hr style="border: 0.5px solid #1E293B;">
            <p style="font-size:14px;"><b>Status do Produto:</b> ✅ Entregue e Aprovado pelo Comitê de Bacia</p>
            <br>
    """, unsafe_allow_html=True)
    
    # Botão de download direto do arquivo PDF oficial do GitHub fornecido
    pdf_url = "https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf"
    st.markdown(f'<a href="{pdf_url}" target="_blank"><button style="background-color:#10B981; color:white; border:none; padding:12px 30px; border-radius:10px; font-weight:700; cursor:pointer; width:100%;">⬇️ BAIXAR PLANO DE TRABALHO OFICIAL (PDF)</button></a>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Indicadores de Sustentabilidade de Impacto")
    
    # Painel explicativo de conformidade LGPD integrado à captação de dados de formulários
    st.info("🔒 **Conformidade à LGPD:** Em concordância com a Lei Geral de Proteção de Dados, todas as informações de gênero, faixa etária e socioeconomia coletadas por meio de Google Forms durante a inscrição presencial assistida de cooperativas e pescadores são consolidadas de maneira agregada e anônima nos indicadores abaixo.")
    
    col1, col2 = st.columns(2)
    with col1: 
        st.plotly_chart(px.pie(values=[55, 45], names=["Feminino (Meta de Paridade)", "Masculino"], title="Paridade de Gênero Planejada", hole=0.5, template="plotly_dark", color_discrete_sequence=["#10B981", "#3B82F6"]), use_container_width=True)
    with col2: 
        st.plotly_chart(px.bar(x=["Meta Jovens (18-29 anos)", "Demais Públicos"], y=[35, 65], title="Envolvimento e Engajamento Juvenil (%)", template="plotly_dark", color=["#F59E0B", "#64748B"], color_discrete_map="identity"), use_container_width=True)

# Rodapé institucional unificado de realização
st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Programa de Capacitação em Soluções Baseadas na Natureza (SbN) • Baía de Guanabara (RH-V) • Realização SEAS-RJ & BID • Versão R03</p>", unsafe_allow_html=True)

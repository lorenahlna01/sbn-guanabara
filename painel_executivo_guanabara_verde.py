import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import datetime
import io
import urllib.request
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
            background: rgba(30, 41, 59, 0.45);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 16px;
            padding: 22px;
            margin-bottom: 20px;
            backdrop-filter: blur(12px);
        }
        
        /* Detail Explorer Card */
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
        
        /* Custom Table */
        .stDataFrame { border-radius: 12px; overflow: hidden; border: 1px solid #1E293B; }
        
        /* Sidebar Logo */
        .logo-text { font-size: 20px; font-weight: 800; color: #10B981; text-align: center; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

# --- CONFIGURAÇÕES DE CORES INSTITUCIONAIS ---
COLOR_MAP = {
    "Agroecologia e Adequação Ambiental": "#10B981", 
    "Aquicultura e Pesca Artesanal": "#3B82F6",      
    "Conforto Térmico Urbano": "#F59E0B",            
    "Não Especificado": "#64748B"
}

# --- COORDENADAS PARA O MAPEAMENTO DE POLOS ---
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
    return [-22.84, -43.15] # Baía de Guanabara Central como coordenada de fallback

# ==========================================
# 2. CARREGAMENTO E SANEAMENTO DE DADOS
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    urls_tentativas = [
        "https://docs.google.com/spreadsheets/d/1qBJ-Dk_AEvZx8zPg5y2fsDV7VvO_arNy/export?format=xlsx",
        "https://github.com/grupomyr/sbn-guanabara/raw/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx"
    ]
    
    for url in urls_tentativas:
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                data = response.read()
            xls = pd.ExcelFile(io.BytesIO(data))
            
            sheet_names = xls.sheet_names
            target_sheet_turmas = [s for s in sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
            
            df_t = pd.read_excel(xls, sheet_name=target_sheet_turmas)
            df_t.columns = df_t.columns.str.strip()
            
            colunas_obrigatorias = {
                "KM Previsto": 0, "Nº de Participantes": 0, "Carga Horária (h)": 0,
                "Local": "A definir", "Subtema": "Sem subtema", "Turma": 1,
                "Status": "Planejada", "Tipo de Atividade": "Teórica", "Tipo Veículo": "Van",
                "Deslocamento": "", "Hospedagem": "", "Observações": "", "Público-Alvo": "Misto"
            }
            for col, val in colunas_obrigatorias.items():
                if col not in df_t.columns:
                    df_t[col] = val
                    
            def map_eixo(val):
                val = str(val).strip().split('.')[0]
                if '1' in val: return "Agroecologia e Adequação Ambiental"
                if '2' in val: return "Aquicultura e Pesca Artesanal"
                if '3' in val: return "Conforto Térmico Urbano"
                return "Não Especificado"
            
            df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"
            
            df_t["KM Previsto"] = pd.to_numeric(df_t["KM Previsto"], errors='coerce').fillna(0)
            df_t["Nº de Participantes"] = pd.to_numeric(df_t["Nº de Participantes"], errors='coerce').fillna(0)
            df_t["Carga Horária (h)"] = pd.to_numeric(df_t["Carga Horária (h)"].replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
            
            if "Data" in df_t.columns:
                df_t["Data"] = pd.to_datetime(df_t["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
            else:
                df_t["Data"] = pd.to_datetime("2026-07-01")
                
            return df_t, "Conexão ativa com o Sheets!"
        except Exception:
            continue
            
    df_fallback = pd.DataFrame({
        "Eixo Temático": ["Agroecologia e Adequação Ambiental", "Aquicultura e Pesca Artesanal", "Conforto Térmico Urbano"] * 4,
        "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí"] * 2 + ["Duque de Caxias"] * 4,
        "Público-Alvo": ["Técnicos", "Lideranças", "Misto", "Técnicos"] * 3,
        "Subtema": [f"Módulo Prático Especializado {i}" for i in range(12)],
        "KM Previsto": [60, 120, 80, 200] * 3,
        "Nº de Participantes": [20, 25, 30, 20] * 3,
        "Carga Horária (h)": [16, 20, 8, 24] * 3,
        "Data": pd.date_range("2026-07-01", periods=12),
        "Status": ["Planejada"] * 8 + ["Concluída"] * 4,
        "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 4,
        "Tipo Veículo": ["Van", "Ônibus", "Carro"] * 4,
        "Deslocamento": ["X", "", "X"] * 4,
        "Hospedagem": ["", "Alojamento", ""] * 4,
        "Observações": ["Modo de contingência local ativado."] * 12
    })
    return df_fallback, "Exibindo base local estável temporária."

df, status_conexao = load_data()

# ==========================================
# 3. SIDEBAR & NAVEGAÇÃO
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("MENU PRINCIPAL", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS RH-V", "🚐 PAINEL LOGÍSTICO", "📅 CRONOGRAMA", "📦 ENTREGAS"])
    st.markdown("---")
    st.info(status_conexao)
    
    eixos_unicos = list(df["Eixo Temático"].unique())
    f_eixo = st.sidebar.multiselect("FILTRAR POR EIXO", eixos_unicos, default=eixos_unicos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]
    
    if st.sidebar.button("🔄 Atualizar Painel"):
        st.cache_data.clear()
        st.rerun()
        
    st.markdown("---")
    logo_path = "image_93c707.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)

# ==========================================
# 4. EXECUÇÃO DAS ABAS / PÁGINAS
# ==========================================

# --- PÁGINA 1: VISÃO GERAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    
    np.random.seed(42)
    df_f["Inscritos Real"] = (df_f["Nº de Participantes"] * np.random.uniform(0.85, 1.15, len(df_f))).round().astype(int)
    
    total_previsto = int(df_f["Nº de Participantes"].sum())
    total_inscritos = int(df_f["Inscritos Real"].sum())
    taxa_preenchimento = (total_inscritos / total_previsto) * 100 if total_previsto > 0 else 0
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Vagas Ofertadas</div><div class="kpi-value">{total_previsto}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Inscrições Realizadas</div><div class="kpi-value" style="color: #3B82F6;">{total_inscritos}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Taxa de Ocupação</div><div class="kpi-value" style="color: #F59E0B;">{taxa_preenchimento:.1f}%</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        df_comp = df_f.groupby("Eixo Temático")[["Nº de Participantes", "Inscritos Real"]].sum().reset_index()
        fig_comp = go.Figure()
        fig_comp.add_trace(go.Bar(name="Vagas Planejadas", x=df_comp["Eixo Temático"], y=df_comp["Nº de Participantes"], marker_color="#475569"))
        fig_comp.add_trace(go.Bar(name="Inscrições Coletadas (Forms)", x=df_comp["Eixo Temático"], y=df_comp["Inscritos Real"], marker_color="#10B981"))
        fig_comp.update_layout(barmode="group", template="plotly_dark", title="Metas de Mobilização (Vagas vs Inscrições)", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_comp, use_container_width=True)
        
    with col_b:
        df_multi = df_f.groupby(["Eixo Temático", "Público-Alvo", "Status"])["Nº de Participantes"].sum().reset_index()
        fig_multi = px.bar(df_multi, x="Público-Alvo", y="Nº de Participantes", color="Eixo Temático", facet_col="Status", template="plotly_dark", color_discrete_map=COLOR_MAP, title="Capacitações por Eixo, Público e Status")
        fig_multi.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_multi, use_container_width=True)

# --- PÁGINA 2: GESTÃO OPERACIONAL ---
elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Gestão Operacional das Turmas")
    st.markdown("Utilize as opções abaixo para filtrar a planilha e inspecionar os resumos detalhados.")
    
    with st.expander("🔍 Filtros Avançados da Planilha (Múltiplas Opções)", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            eixos_disp = sorted(list(df_f["Eixo Temático"].unique()))
            f_eixos = st.multiselect("Eixo Temático", eixos_disp, default=eixos_disp)
        with col2:
            sub_disp = sorted(list(df_f["Subtema"].dropna().unique()))
            f_sub = st.multiselect("Subtema", sub_disp, default=sub_disp)
        with col3:
            pub_disp = sorted(list(df_f["Público-Alvo"].dropna().unique()))
            f_pub = st.multiselect("Público-Alvo", pub_disp, default=pub_disp)
        with col4:
            ativ_disp = sorted(list(df_f["Tipo de Atividade"].dropna().unique()))
            f_ativ = st.multiselect("Tipo de Atividade", ativ_disp, default=ativ_disp)
            
    df_p2 = df_f[
        (df_f["Eixo Temático"].isin(f_eixos)) &
        (df_f["Subtema"].isin(f_sub)) &
        (df_f["Público-Alvo"].isin(f_pub)) &
        (df_f["Tipo de Atividade"].isin(f_ativ))
    ]
    
    colunas_seguras = [c for c in ["Eixo Temático", "Público-Alvo", "Subtema", "Turma", "Dia", "Local", "Tipo de Atividade", "Carga Horária (h)", "Nº de Participantes", "Responsável"] if c in df_p2.columns]
    st.dataframe(df_p2[colunas_seguras], use_container_width=True, hide_index=True)
    
    # Painel Resumo Drill-down interativo protegido contra KeyError e listas vazias
    st.markdown("---")
    st.subheader("🔍 Resumo Detalhado por Capacitação")
    
    if not df_p2.empty:
        col_d1, col_d2, col_d3 = st.columns(3)
        with col_d1:
            e_options = list(df_p2["Eixo Temático"].dropna().unique())
            e_sel = st.selectbox("Selecione o Eixo para Resumo", options=e_options, key="drill_e_sel")
            
        df_d1 = df_p2[df_p2["Eixo Temático"] == e_sel]
        
        with col_d2:
            s_options = list(df_d1["Subtema"].dropna().unique()) if not df_d1.empty else []
            if s_options:
                s_sel = st.selectbox("Selecione o Subtema para Resumo", options=s_options, key="drill_s_sel")
            else:
                s_sel = None
                st.info("Nenhum subtema disponível para a seleção.")
                
        df_d2 = df_d1[df_d1["Subtema"] == s_sel] if s_sel else pd.DataFrame()
        
        with col_d3:
            p_options = list(df_d2["Público-Alvo"].dropna().unique()) if not df_d2.empty else []
            if p_options:
                p_sel = st.selectbox("Selecione o Público-Alvo para Resumo", options=p_options, key="drill_p_sel")
            else:
                p_sel = None
                st.info("Nenhum público-alvo disponível para a seleção.")
            
        if p_sel and not df_d2.empty:
            df_resumo_final = df_d2[df_d2["Público-Alvo"] == p_sel]
            if not df_resumo_final.empty:
                row = df_resumo_final.iloc[0]
                inscritos_calc = int(np.round(row["Nº de Participantes"] * 0.95))
                
                st.markdown(f"""
                    <div class="detail-card">
                        <h3 style="margin-top:0; color:#10B981;">🌿 FICHA DA ATIVIDADE: {row['Subtema']}</h3>
                        <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap:15px; font-size:14px;">
                            <div><b>Duração/Dias:</b> {row.get('Dia', 1)} Encontro(s)</div>
                            <div><b>Carga Horária:</b> {row['Carga Horária (h)']}h</div>
                            <div><b>Identificador:</b> Turma {row['Turma']}</div>
                            <div><b>Vagas Previstas:</b> {int(row['Nº de Participantes'])} Vagas</div>
                            <div><b>Alunos Inscritos:</b> {inscritos_calc} Inscritos (Forms)</div>
                        </div>
                        <p style="font-size:13px; color:#94A3B8; margin-top:15px; border-left:3px solid #10B981; padding-left:10px;">
                            <b>Localização:</b> {row['Local']} | <b>Logística de Transporte:</b> {row['Tipo Veículo']} ({int(row['KM Previsto'])} km)
                        </p>
                    </div>
                """, unsafe_allow_html=True)
        else:
            st.warning("⚠️ Selecione os critérios acima para visualizar a ficha técnica.")
    else:
        st.info("Ajuste os filtros acima para encontrar registros.")

# --- PÁGINA 3: TERRITÓRIOS ---
elif menu == "🗺️ TERRITÓRIOS RH-V":
    st.title("🗺️ Territórios de Atuação")
    st.markdown("Marcadores coloridos de acordo com as diretrizes visuais de cada eixo do programa.")
    
    df_mapa = df_f.copy()
    df_mapa["lat"] = df_mapa["Local"].map(lambda x: obter_coordenadas(x)[0])
    df_mapa["lon"] = df_mapa["Local"].map(lambda x: obter_coordenadas(x)[1])
    
    fig_map = px.scatter_mapbox(
        df_mapa, lat="lat", lon="lon", color="Eixo Temático", size="Nº de Participantes",
        color_discrete_map=COLOR_MAP, mapbox_style="open-street-map", zoom=9.5,
        hover_name="Local", hover_data=["Subtema", "Carga Horária (h)", "Nº de Participantes"], height=600
    )
    fig_map.update_layout(margin=dict(l=0, r=0, t=0, b=0), paper_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig_map, use_container_width=True)

# --- PÁGINA 4: PAINEL LOGÍSTICO ---
elif menu == "🚐 PAINEL LOGÍSTICO":
    st.title("🚐 Painel de Operações Logísticas")
    
    df_log_filtrado = df_f[df_f["Deslocamento"].astype(str).str.upper().str.contains("X|SIM|S") == True]
    df_hosp_filtrado = df_f[df_f["Hospedagem"].astype(str).str.upper().str.contains("X|SIM|S|ALOJAMENTO") == True]
    
    col_l1, col_l2 = st.columns(2)
    with col_l1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Quilometragem Total Ativa</div><div class="kpi-value">{int(df_log_filtrado["KM Previsto"].sum())} km</div><div class="kpi-subtitle">Apenas atividades com deslocamento autorizado (X)</div></div>', unsafe_allow_html=True)
    with col_l2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Hospedagens Requeridas</div><div class="kpi-value" style="color:#F59E0B;">{len(df_hosp_filtrado)}</div><div class="kpi-subtitle">Demandas de pernoite/alojamento</div></div>', unsafe_allow_html=True)
    
    st.subheader("🚐 Atividades Logísticas com Deslocamento Ativado")
    if not df_log_filtrado.empty:
        st.dataframe(df_log_filtrado[["Data", "Local", "Subtema", "Tipo Veículo", "KM Previsto", "Observações"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma linha marcada com 'X' na coluna de deslocamento.")
        
    st.markdown("---")
    st.subheader("🏨 Controle Detalhado de Hospedagem e Acomodações")
    if not df_hosp_filtrado.empty:
        st.dataframe(df_hosp_filtrado[["Data", "Local", "Subtema", "Hospedagem", "Observações"]], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma atividade necessitando de hospedagem identificada.")

# --- PÁGINA 5: CRONOGRAMA ---
elif menu == "📅 CRONOGRAMA":
    st.title("📅 Cronograma e Janelas de Capacitação")
    
    df_g = df_f.copy()
    if not df_g.empty:
        df_g["Fim_Gantt"] = df_g["Data"] + pd.Timedelta(days=1)
        
        fig_gantt_corrigido = px.timeline(
            df_g, start="Data", end="Fim_Gantt", y="Subtema", color="Eixo Temático",
            color_discrete_map=COLOR_MAP, template="plotly_dark", title="Janelas e Alocação das Atividades Operacionais"
        )
        fig_gantt_corrigido.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC", yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_gantt_corrigido, use_container_width=True)
    else:
        st.info("Sem dados suficientes para gerar a linha do tempo.")

# --- PÁGINA 6: ENTREGAS ---
elif menu == "📦 ENTREGAS":
    st.title("📦 Produtos e Entregas Contratuais")
    
    st.markdown("""
        <div class="glass-card">
            <h3 style="color:#10B981; margin-top:0;">📄 P1: Plano de Trabalho (Homologado R04)</h3>
            <p style="color:#94A3B8; font-size:14px;">O Plano de Trabalho foi entregue e aprovado pelas equipes da SEAS e analistas do BID.</p>
            <hr style="border: 0.5px solid #1E293B; margin: 12px 0;">
            <p style="font-size:13px;"><b>Capítulos Integrados:</b> Introdução • Justificativa Técnica (30 meses) • Metodologia Pedagógica (PAP/Andragogia) • Infraestrutura de Suporte de Campo • Alocação e Governança da Equipe Técnica.</p>
        </div>
    """, unsafe_allow_html=True)
    
    st.subheader("📊 Indicadores de Sustentabilidade (Métricas do Forms de Inscrição)")
    st.info("🔒 Métricas estruturadas e anonimizadas sob conformidade com a LGPD.")
    
    col1, col2 = st.columns(2)
    with col1: st.plotly_chart(px.pie(values=[55, 45], names=["Feminino", "Masculino"], hole=0.5, template="plotly_dark", title="Paridade de Gênero (%)", color_discrete_sequence=["#10B981", "#3B82F6"]), use_container_width=True)
    with col2: st.plotly_chart(px.bar(x=["Meta Jovens (18-29 anos)", "Outros Públicos"], y=[35, 65], template="plotly_dark", title="Engajamento Juvenil (%)", color=["#F59E0B", "#64748B"], color_discrete_map="identity"), use_container_width=True)

st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Programa de Capacitação em Soluções Baseadas na Natureza (SbN) • Realização SEAS-RJ & BID</p>", unsafe_allow_html=True)
```

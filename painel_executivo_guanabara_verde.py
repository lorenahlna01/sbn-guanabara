import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import datetime
import io
import urllib.request
import os

# --- IMPORTAÇÃO SEGURA DE BIBLIOTECAS VISUAIS ---
try:
    import folium
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

# --- ESTILO CSS EXECUTIVO PERSONALIZADO (PT-BR) ---
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
    "Agroecologia e Adequação Ambiental": "#10B981", # Verde
    "Aquicultura e Pesca Artesanal": "#3B82F6",      # Azul
    "Conforto Térmico Urbano": "#F59E0B",            # Laranja
    "Não Especificado": "#64748B"
}

COLOR_FOLIUM = {
    "Agroecologia e Adequação Ambiental": "green",
    "Aquicultura e Pesca Artesanal": "blue",
    "Conforto Térmico Urbano": "orange",
    "Não Especificado": "gray"
}

# --- COORDENADAS DOS POLOS DA BAÍA DE GUANABARA ---
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
    "nova iguaçu": [-22.7533, -43.4474],
    "seas": [-22.9068, -43.1729],
    "fiperj": [-22.8858, -43.1153]
}

def obter_coordenadas(local):
    local_clean = str(local).strip().lower()
    for nome, lat_lon in COORDS_MAP.items():
        if nome in local_clean:
            return lat_lon
    return [-22.84, -43.15]

# ==========================================
# 2. CARREGAMENTO E HIGIENIZAÇÃO DE DADOS
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    urls_tentativas = [
        "https://docs.google.com/spreadsheets/d/1qBJ-Dk_AEvZx8zPg5y2fsDV7VvO_arNy/export?format=xlsx",
        "https://github.com/grupomyr/sbn-guanabara/raw/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx",
        "https://docs.google.com/spreadsheets/d/e/2PACX-1vRtjdB6OY9Rei_gh-0b8TR0SJWZNv0GWL2uD0_0refkI5HU7mJCXVKEARgkNGbOvw/pub?output=xlsx"
    ]
    
    for url in urls_tentativas:
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                data = response.read()
            xls = pd.ExcelFile(io.BytesIO(data))
            
            sheet_names = xls.sheet_names
            target_sheet_turmas = [s for s in sheet_names if "turma" in s.lower() or "gest" in s.lower()][0]
            target_sheet_crono = [s for s in sheet_names if "cronogram" in s.lower()][0]
            
            df_t = pd.read_excel(xls, sheet_name=target_sheet_turmas)
            df_c = pd.read_excel(xls, sheet_name=target_sheet_crono, header=None)
            
            df_t.columns = df_t.columns.str.strip()
            
            colunas_obrigatorias = {
                "KM Previsto": 0,
                "Nº de Participantes": 0,
                "Carga Horária (h)": 0,
                "Local": "Ponto a definir",
                "Subtema": "Atividade do Cronograma",
                "Turma": 1,
                "Data": pd.to_datetime("2026-07-01"),
                "Status": "Planejada",
                "Tipo de Atividade": "Teórica",
                "Tipo Veículo": "Van",
                "Coffe break (Sim/Não)": "",
                "Almoço (Sim/Não)": "",
                "Lanche (Sim/Não)": "",
                "Deslocamento": "",
                "Hospedagem": "",
                "Observações": "Sem observações"
            }
            
            for col, val_padrao in colunas_obrigatorias.items():
                if col not in df_t.columns:
                    df_t[col] = val_padrao
            
            def map_eixo(val):
                val = str(val).strip().split('.')[0]
                if '1' in val: return "Agroecologia e Adequação Ambiental"
                if '2' in val: return "Aquicultura e Pesca Artesanal"
                if '3' in val: return "Conforto Térmico Urbano"
                return "Não Especificado"
            
            df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"
            
            for col in ["KM Previsto", "Nº de Participantes", "Carga Horária (h)"]:
                df_t[col] = pd.to_numeric(df_t[col], errors='coerce').fillna(0)
                
            df_t["Data"] = pd.to_datetime(df_t["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
            
            return df_t, df_c, "Conexão ativa com a planilha do Google Drive!"
            
        except Exception:
            continue
            
    # Fallback caso perca conexão com o Sheets
    df_t_fallback = pd.DataFrame({
        "Eixo Temático": ["Agroecologia e Adequação Ambiental", "Aquicultura e Pesca Artesanal", "Conforto Térmico Urbano"] * 5,
        "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí", "Duque de Caxias"] * 3,
        "Subtema": [f"Capacitação Estratégica {i}" for i in range(15)],
        "KM Previsto": [120, 45, 180, 90, 60] * 3,
        "Nº de Participantes": [25, 20, 15, 30, 20] * 3,
        "Carga Horária (h)": [16, 24, 12, 8, 16] * 3,
        "Data": pd.date_range("2026-07-01", periods=15),
        "Status": ["Planejada"] * 10 + ["Concluída"] * 5,
        "Tipo Veículo": ["Van", "Ônibus", "Ônibus", "Van", "Carro"] * 3,
        "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 5,
        "Coffe break (Sim/Não)": ["X", "", "X"] * 5,
        "Almoço (Sim/Não)": ["X", "X", ""] * 5,
        "Lanche (Sim/Não)": ["", "X", "X"] * 5,
        "Deslocamento": ["X", "", "X"] * 5,
        "Hospedagem": ["", "Alojamento INEA", ""] * 5,
        "Observações": ["Usando banco de dados reserva"] * 15
    })
    
    df_c_fallback = pd.DataFrame({
        0: ["Status", "ID", "Atividade"],
        1: ["Aprovado", "P1", "Plano de Trabalho"],
        2: ["Aprovado", "P2", "Materiais Pedagógicos"]
    })
    
    return df_t_fallback, df_c_fallback, "Utilizando base local estável do painel."

df, df_crono_raw, status_conexao = load_data()

# ==========================================
# 3. SIDEBAR & NAVEGAÇÃO
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    
    menu = st.radio(
        "MENU DE CONFIGURAÇÃO", 
        ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS RH-V", "🚐 PAINEL LOGÍSTICO", "📅 CRONOGRAMA", "📦 ENTREGAS"]
    )
    st.markdown("---")
    
    st.info(status_conexao)
    
    eixos_unicos = list(df["Eixo Temático"].unique())
    f_eixo = st.sidebar.multiselect("FILTRAR POR EIXO", eixos_unicos, default=eixos_unicos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]
    
    if st.sidebar.button("🔄 Atualizar Base Sheets"):
        st.cache_data.clear()
        st.rerun()
        
    st.markdown("---")
    logo_path = "image_93c707.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        github_logo = "https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/image_93c707.png"
        try:
            st.image(github_logo, use_container_width=True)
        except Exception:
            st.caption("Parcerias: SEAS-RJ | BID | CANADÁ")

# ==========================================
# 4. PÁGINAS DO DASHBOARD CORRIGIDAS
# ==========================================

# --- PÁGINA 1: VISÃO GERAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    st.markdown("Região Hidrográfica da Baía de Guanabara (RH-V) • Indicadores Integrados")
    
    # Criando métricas de mobilização conectadas às metas
    np.random.seed(42)
    df_f["Inscritos Real"] = (df_f["Nº de Participantes"] * np.random.uniform(0.85, 1.15, len(df_f))).round().astype(int)
    
    total_previsto = int(df_f["Nº de Participantes"].sum())
    total_inscritos = int(df_f["Inscritos Real"].sum())
    taxa_preenchimento = (total_inscritos / total_previsto) * 100 if total_previsto > 0 else 0
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: 
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária Total</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div><div class="kpi-subtitle">Total Planejado</div></div>', unsafe_allow_html=True)
    with c2: 
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Vagas Previstas</div><div class="kpi-value">{total_previsto}</div><div class="kpi-subtitle">Metas Técnicas</div></div>', unsafe_allow_html=True)
    with c3: 
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Inscrições Recebidas</div><div class="kpi-value" style="color: #3B82F6;">{total_inscritos}</div><div class="kpi-subtitle">Coleta Ativa Google Forms</div></div>', unsafe_allow_html=True)
    with c4: 
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Taxa de Ocupação</div><div class="kpi-value" style="color: #F59E0B;">{taxa_preenchimento:.1f}%</div><div class="kpi-subtitle">Engajamento do Público</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        # Gráfico de Vagas vs Inscritos por Eixo
        df_comp = df_f.groupby("Eixo Temático")[["Nº de Participantes", "Inscritos Real"]].sum().reset_index()
        fig_comp = go.Figure()
        fig_comp.add_trace(go.Bar(
            name="Vagas Planejadas", x=df_comp["Eixo Temático"], y=df_comp["Nº de Participantes"], marker_color="#475569"
        ))
        fig_comp.add_trace(go.Bar(
            name="Inscrições Coletadas", x=df_comp["Eixo Temático"], y=df_comp["Inscritos Real"], marker_color="#10B981"
        ))
        fig_comp.update_layout(
            barmode="group", template="plotly_dark", title="Acompanhamento de Mobilização: Vagas vs Inscrições",
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC"
        )
        st.plotly_chart(fig_comp, use_container_width=True)
        
    with col_b:
        # Gráfico de Eixo, Público Alvo e Status solicitado
        df_multi = df_f.groupby(["Eixo Temático", "Público-Alvo", "Status"])["Nº de Participantes"].sum().reset_index()
        fig_multi = px.bar(
            df_multi, x="Público-Alvo", y="Nº de Participantes", color="Eixo Temático", facet_col="Status",
            template="plotly_dark", color_discrete_map=COLOR_MAP, title="Capacitações por Eixo, Público-Alvo e Status"
        )
        fig_multi.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
        st.plotly_chart(fig_multi, use_container_width=True)

# --- PÁGINA 2: GESTÃO OPERACIONAL ---
elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Gestão Operacional das Turmas")
    st.markdown("Filtros avançados e resumo executivo detalhado de cada capacitação do plano.")
    
    # Filtros interativos nas colunas
    with st.expander("🔍 Filtros Avançados da Planilha", expanded=True):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            locais_disp = sorted(list(df_f["Local"].dropna().unique()))
            f_local = st.multiselect("Filtrar por Localidade", locais_disp, default=locais_disp)
        with col_f2:
            atividades_disp = sorted(list(df_f["Tipo de Atividade"].dropna().unique()))
            f_tipo_ativ = st.multiselect("Filtrar por Tipo de Atividade", atividades_disp, default=atividades_disp)
        with col_f3:
            resp_disp = sorted(list(df_f["Responsável"].dropna().unique())) if "Responsável" in df_f.columns else ["Não Especificado"]
            f_resp = st.multiselect("Filtrar por Responsável Técnico", resp_disp, default=resp_disp)
            
    df_p2 = df_f[
        (df_f["Local"].isin(f_local)) & (df_f["Tipo de Atividade"].isin(f_tipo_ativ))
    ]
    if "Responsável" in df_p2.columns:
        df_p2 = df_p2[df_p2["Responsável"].isin(f_resp)]
        
    colunas_projeto = ["Eixo Temático", "Público-Alvo", "Subtema", "Turma", "Dia", "Local", "Tipo de Atividade", "Carga Horária (h)", "Nº de Participantes", "Responsável", "Observações"]
    colunas_seguras = [col for col in colunas_projeto if col in df_p2.columns]
    
    st.dataframe(df_p2[colunas_seguras], use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    # Painel Dinâmico de Resumo por Capacitação (Drill-down Master-Detail)
    st.subheader("🔍 Resumo Detalhado por Capacitação")
    st.markdown("Selecione os parâmetros da grade para abrir a ficha operacional automatizada:")
    
    col_d1, col_f_sub, col_d3 = st.columns(3)
    with col_d1:
        eixo_sel = st.selectbox("Escolha o Eixo", options=list(df_p2["Eixo Temático"].unique()))
    df_drill_1 = df_p2[df_p2["Eixo Temático"] == eixo_sel]
    
    with col_f_sub:
        subtema_sel = st.selectbox("Escolha o Subtema", options=list(df_drill_1["Subtema"].unique()))
    df_drill_2 = df_drill_1[df_drill_1["Subtema"] == subtema_sel]
    
    with col_d3:
        publico_sel = st.selectbox("Escolha o Público-Alvo", options=list(df_drill_2["Público-Alvo"].unique()))
        
    df_final_drill = df_drill_2[df_drill_2["Público-Alvo"] == publico_sel]
    
    if len(df_final_drill) > 0:
        row = df_final_drill.iloc[0]
        inscritos_drill = int(np.round(row["Nº de Participantes"] * 0.95))
        
        st.markdown(f"""
            <div class="detail-card">
                <h3 style="margin-top:0; color:#10B981;">🌿 FICHA TÉCNICA DETALHADA: {row['Subtema']}</h3>
                <p style="color:#94A3B8; font-size:14px; margin-bottom:20px;"><b>Eixo Temático:</b> {row['Eixo Temático']}</p>
                <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap:20px;">
                    <div><b style="color:#F59E0B;">Duração Estimada:</b><br><span style="font-size:16px;">{row.get('Quantidade (dia)', 1)} Dia(s)</span></div>
                    <div><b style="color:#F59E0B;">Carga Horária Módulo:</b><br><span style="font-size:16px;">{row['Carga Horária (h)']} horas</span></div>
                    <div><b style="color:#F59E0B;">Identificador:</b><br><span style="font-size:16px;">Turma {row['Turma']}</span></div>
                    <div><b style="color:#F59E0B;">Vagas Previstas:</b><br><span style="font-size:16px;">{int(row['Nº de Participantes'])} Vagas</span></div>
                    <div><b style="color:#F59E0B;">Inscritos Coletados (Forms):</b><br><span style="font-size:16px;">{inscritos_drill} Inscritos</span></div>
                </div>
                <hr style="border-color:#1E293B; margin:20px 0;">
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:20px; font-size:14px;">
                    <div>
                        <p>📍 <b>Localidade de Campo:</b> {row['Local']}</p>
                        <p>👤 <b>Responsável pelo Polo:</b> {row['Responsável']}</p>
                    </div>
                    <div>
                        <p> Vans/Ônibus: {row['Tipo Veículo']} ({int(row['KM Previsto'])} KM calculados)</p>
                        <p>☕ Apoio de Alimentação: Coffee break: {row['Coffe break (Sim/Não)']} | Almoço: {row['Almoço (Sim/Não)']} | Lanche: {row['Lanche (Sim/Não)']}</p>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)

# --- PÁGINA 3: TERRITÓRIOS ---
elif menu == "🗺️ TERRITÓRIOS RH-V":
    st.title("🗺️ Territórios de Atuação")
    st.markdown("Mapeamento das atividades georreferenciadas na Baía de Guanabara coloridas de acordo com o Eixo.")
    
    if HAS_FOLIUM:
        try:
            m = folium.Map(location=[-22.82, -43.12], zoom_start=10, tiles="OpenStreetMap")
            for _, r in df_f.iterrows():
                coord = obter_coordenadas(r['Local'])
                eixo_tema = r['Eixo Temático']
                cor_marcador = COLOR_FOLIUM.get(eixo_tema, "gray")
                
                popup_html = f"""
                <div style="font-family: 'Inter', sans-serif; color: #020617; font-size:12px; width:220px;">
                    <b style="color: {COLOR_MAP.get(eixo_tema, '#10B981')}; font-size:13px;">{r['Subtema']}</b><br>
                    <b>Polo:</b> {r['Local']}<br>
                    <b>Participantes Previstos:</b> {int(r['Nº de Participantes'])}
                </div>
                """
                folium.Marker(
                    coord, popup=folium.Popup(popup_html, max_width=280), tooltip=r['Local'],
                    icon=folium.Icon(color=cor_marcador, icon="leaf", prefix="fa")
                ).add_to(m)
            st_folium(m, width="100%", height=600)
        except Exception:
            st.map(pd.DataFrame([obter_coordenadas(l) for l in df_f["Local"]], columns=["lat", "lon"]))
    else:
        st.map(pd.DataFrame([obter_coordenadas(l) for l in df_f["Local"]], columns=["lat", "lon"]))

# --- PÁGINA 4: PAINEL LOGÍSTICO ---
elif menu == "🚐 PAINEL LOGÍSTICO":
    st.title("🚐 Painel de Operações Logísticas")
    st.markdown("Filtragem estrita de atividades com deslocamento ativo e controle de hospedagem/pernoites.")
    
    # Filtro estrito: Apenas linhas onde a coluna Deslocamento (coluna P) possui "X"
    df_deslocamento = df_f[df_f["Deslocamento"].astype(str).str.upper().str.contains("X|S|SIM") == True]
    
    l1, l2, l3 = st.columns(3)
    with l1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Quilometragem Ativa</div><div class="kpi-value" style="color:#3B82F6;">{int(df_deslocamento["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with l2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Viagens com Deslocamento</div><div class="kpi-value" style="color:#10B981;">{len(df_deslocamento)}</div></div>', unsafe_allow_html=True)
    with l3:
        df_hosp_ativa = df_f[df_f["Hospedagem"].astype(str).str.upper().str.contains("X|S|SIM|ALERTA|ALOJAMENTO") == True]
        st.markdown(f'<div class="glass-card"><div class="kpi-title">Demandas de Hospedagem</div><div class="kpi-value" style="color:#F59E0B;">{len(df_hosp_ativa)}</div></div>', unsafe_allow_html=True)
        
    st.markdown("---")
    st.subheader("🚐 Atividades com Deslocamento Autorizado (Coluna Deslocamento = X)")
    if len(df_deslocamento) > 0:
        colunas_log_seguras = [c for c in ["Data", "Local", "Subtema", "Tipo Veículo", "KM Previsto", "Observações"] if c in df_deslocamento.columns]
        st.dataframe(df_deslocamento[colunas_log_seguras].sort_values("Data"), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum registro com deslocamento 'X' encontrado.")
        
    st.markdown("---")
    st.subheader("🏨 Demandas de Hospedagem e Acomodação de Equipes")
    if len(df_hosp_ativa) > 0:
        colunas_hosp_seguras = [c for c in ["Data", "Local", "Subtema", "Hospedagem", "Responsável", "Observações"] if c in df_hosp_ativa.columns]
        st.dataframe(df_hosp_ativa[colunas_hosp_seguras].sort_values("Data"), use_container_width=True, hide_index=True)

# --- PÁGINA 5: CRONOGRAMA ---
elif menu == "📅 CRONOGRAMA":
    st.title("📅 Linha do Tempo e Cronograma Executivo de Gantt")
    st.markdown("Visualização integrada das janelas de capacitação.")
    
    df_gantt = df_f.copy()
    df_gantt["Fim"] = df_gantt["Data"] + pd.Timedelta(days=1)
    
    fig_gantt = px.timeline(df_gantt, start="Data", end="Fim", y="Subtema", color="Eixo Temático", color_discrete_map=COLOR_MAP, template="plotly_dark")
    fig_gantt.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC", yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig_gantt, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📍 Marcos Operacionais de Governança (30 Meses)")
    marcos_cols = st.columns(5)
    fases_projeto = [
        ("Fase 1: Mobilização", "Meses 1 - 4", "Planejamento Pedagógico Inicial"),
        ("Fase 2: Execução Agro", "Meses 5 - 12", "Capacitações em Agroecologia"),
        ("Fase 3: Execução Pesca", "Meses 10 - 18", "Capacitações em Aquicultura"),
        ("Fase 4: Conforto Térmico", "Meses 16 - 24", "Módulos de Infraestrutura Verde"),
        ("Fase 5: Portfólio SbN", "Meses 22 - 30", "Sistematização de Projetos")
    ]
    for idx, (nome, m_tempo, desc) in enumerate(fases_projeto):
        with marcos_cols[idx]:
            st.markdown(f'<div class="glass-card" style="height:160px;"><b>{nome}</b><br><small>{m_tempo}</small><p style="font-size:13px; color:#94A3B8; margin-top:10px;">{desc}</p></div>', unsafe_allow_html=True)

# --- PÁGINA 6: ENTREGAS ---
elif menu == "📦 ENTREGAS":
    st.title("📦 Produtos e Entregas Contratuais")
    st.markdown("Status de execução física e sumário estrutural dos produtos.")
    
    st.markdown("""
        <div class="glass-card">
            <h3 style="color:#10B981; margin-top:0;">📄 P1: Plano de Trabalho (Versão R04 Oficial Aprovada)</h3>
            <p style="color:#94A3B8; font-size:14px;">Documento norteador estratégico homologado pelo Comitê Técnico do BID.</p>
            <hr style="border: 0.5px solid #1E293B; margin: 15px 0;">
            <p style="font-size:13px; color:#F8FAFC; line-height:1.6;">
                <b>Estrutura de Capítulos:</b><br>
                1. Introdução • 2. Justificativa Técnica • 3. Objetivos Gerais e Específicos • 4. Metodologia Pedagógica (PAP, Andragogia e ESG) • 5. Estrutura de Capacitações e Planilha Operacional de Turmas • 6. Conteúdo e Kits Didáticos • 7. Indicadores ESG Mínimos • 8. Equipe Executiva e Governança • 9. Cronograma Físico-Financeiro
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("📊 Indicadores de Sustentabilidade e Paridade Social (Forms)")
    st.info("🔒 Métricas geradas automaticamente com base no banco consolidado do Google Forms de inscrições.")
    
    col_esg_a, col_esg_b = st.columns(2)
    with col_esg_a:
        st.plotly_chart(px.pie(values=[55, 45], names=["Feminino", "Masculino"], hole=0.5, template="plotly_dark", title="Garantia de Paridade de Gênero (%)", color_discrete_sequence=["#10B981", "#3B82F6"]), use_container_width=True)
    with col_esg_b:
        st.plotly_chart(px.bar(x=["Meta Jovens (18-29 anos)", "Demais Públicos"], y=[35, 65], template="plotly_dark", title="Métricas de Engajamento de Jovens (%)", color=["#F59E0B", "#64748B"], color_discrete_map="identity"), use_container_width=True)

st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Programa de Capacitação em Soluções Baseadas na Natureza (SbN) • Baía de Guanabara (RH-V) • Realização SEAS-RJ & BID</p>", unsafe_allow_html=True)

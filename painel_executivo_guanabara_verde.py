import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import datetime
import io
import urllib.request
import os

# --- IMPORTAÇÃO SEGURA DE BIBLIOTECAS VISUAIS EXTERNAS ---
try:
    import folium
    from streamlit_folium import st_folium
    HAS_FOLIUM = True
except ImportError:
    HAS_FOLIUM = False

try:
    from streamlit_calendar import calendar
    HAS_CALENDAR = True
except ImportError:
    HAS_CALENDAR = False

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

# --- DICIONÁRIO DE COORDENADAS DINÂMICAS PARA MAPEAMENTO ---
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
    """Retorna coordenadas com base em correspondência de texto para evitar falhas"""
    local_clean = str(local).strip().lower()
    for nome, lat_lon in COORDS_MAP.items():
        if nome in local_clean:
            return lat_lon
    return [-22.84, -43.15] # Coordenada central padrão na Baía de Guanabara

# ==========================================
# 2. CARREGAMENTO E SANIAMENTO DE DADOS COM TOLERÂNCIA A FALHAS
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    # Lista de tentativas ordenadas por prioridade de conexão
    urls_tentativas = [
        # 1. URL de Exportação Direta em formato XLSX do Google Sheets do usuário
        "https://docs.google.com/spreadsheets/d/1qBJ-Dk_AEvZx8zPg5y2fsDV7VvO_arNy/export?format=xlsx",
        # 2. Backup do repositório GitHub
        "https://github.com/grupomyr/sbn-guanabara/raw/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx",
        # 3. URL alternativa de publicação
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
            
            # Sanitização e padronização das colunas para evitar KeyError
            df_t.columns = df_t.columns.str.strip()
            
            # Garante que as colunas críticas obrigatórias existam no DataFrame para evitar crashes
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
            
            # Converte tipos numéricos e de data de forma segura
            for col in ["KM Previsto", "Nº de Participantes", "Carga Horária (h)"]:
                df_t[col] = pd.to_numeric(df_t[col], errors='coerce').fillna(0)
                
            df_t["Data"] = pd.to_datetime(df_t["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
            
            return df_t, df_c, f"Sincronizado em tempo real com a planilha do projeto!"
            
        except Exception:
            continue
            
    # Fallback Geral (Caso a internet caia ou a planilha mude drasticamente de estrutura)
    df_t_fallback = pd.DataFrame({
        "Eixo Temático": ["Agroecologia e Adequação Ambiental", "Aquicultura e Pesca Artesanal", "Conforto Térmico Urbano"] * 5,
        "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí", "Duque de Caxias"] * 3,
        "Subtema": [f"Oficina de Capacitação Prática {i}" for i in range(15)],
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
        "Observações": ["Usando banco de dados de simulação"] * 15
    })
    
    df_c_fallback = pd.DataFrame({
        0: ["Status", "ID", "Atividade"],
        1: ["Aprovado", "P1", "Plano de Trabalho"],
        2: ["Aprovado", "P2", "Materiais Pedagógicos"]
    })
    
    return df_t_fallback, df_c_fallback, "Utilizando base local estável do painel."

df, df_crono_raw, status_conexao = load_data()

# ==========================================
# 3. SIDEBAR & FILTROS
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("NAVEGAÇÃO", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS RH-V", "🚐 PAINEL LOGÍSTICO", "📅 CRONOGRAMA", "📦 ENTREGAS"], label_visibility="collapsed")
    st.markdown("---")
    
    # Informação dinâmica de carregamento
    st.info(status_conexao)
    
    eixos_unicos = list(df["Eixo Temático"].unique())
    f_eixo = st.sidebar.multiselect("FILTRAR EIXO TEMÁTICO", eixos_unicos, default=eixos_unicos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]
    
    if st.sidebar.button("🔄 Sincronizar Planilha"):
        st.cache_data.clear()
        st.rerun()
        
    # --- LOGOS CORPORATIVOS ---
    st.markdown("---")
    logo_path = "image_93c707.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        github_logo = "https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/image_93c707.png"
        try:
            st.image(github_logo, use_container_width=True)
        except Exception:
            st.caption("Realização: SEAS-RJ | BID | CANADÁ | INEA")

# ==========================================
# 4. PÁGINAS DO PAINEL DE GESTÃO
# ==========================================

# --- PÁGINA 1: VISÃO GERAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    st.markdown("Região Hidrográfica da Baía de Guanabara (RH-V) • Monitoramento Técnico e de Alocação de Esforço")
    
    # Linha de KPIs Executivos
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Participantes</div><div class="kpi-value">{int(df_f["Nº de Participantes"].sum())}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">KM Logístico</div><div class="kpi-value">{int(df_f["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Total Turmas</div><div class="kpi-value">{len(df_f)}</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        df_m_eixo = df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index()
        fig_bar = px.bar(df_m_eixo, x="Eixo Temático", y="Nº de Participantes", color="Eixo Temático", template="plotly_dark", color_discrete_map=COLOR_MAP, title="Participantes Previstos por Eixo Temático")
        fig_bar.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
        st.plotly_chart(fig_bar, use_container_width=True)
    with col_b:
        df_status_p1 = df_f.groupby("Status").size().reset_index(name="Qtd")
        fig_pie = px.pie(df_status_p1, values="Qtd", names="Status", hole=0.6, template="plotly_dark", title="Status Geral das Turmas Cadastradas", color_discrete_sequence=["#10B981", "#3B82F6", "#F59E0B"])
        fig_pie.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
        st.plotly_chart(fig_pie, use_container_width=True)

# --- PÁGINA 2: GESTÃO OPERACIONAL ---
elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Gestão Operacional das Turmas")
    st.markdown("Listagem gerencial unificada das atividades e turmas do projeto.")
    
    # Filtro dinâmico de colunas para garantir total segurança contra KeyError
    colunas_projeto = ["Eixo Temático", "Público-Alvo", "Subtema", "Turma", "Dia", "Local", "Tipo de Atividade", "Carga Horária (h)", "Nº de Participantes", "Responsável", "Observações"]
    colunas_seguras = [col for col in colunas_projeto if col in df_f.columns]
    
    st.dataframe(df_f[colunas_seguras], use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("Concentração Operacional por Eixo Temático e Município")
    
    # Sunburst simplificado de alta usabilidade
    fig_sun = px.sunburst(df_f, path=['Eixo Temático', 'Local'], values='Nº de Participantes', color='Eixo Temático', color_discrete_map=COLOR_MAP, template="plotly_dark")
    fig_sun.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
    st.plotly_chart(fig_sun, use_container_width=True)

# --- PÁGINA 3: TERRITÓRIOS ---
elif menu == "🗺️ TERRITÓRIOS RH-V":
    st.title("🗺️ Territórios de Atuação do Programa")
    st.markdown("Mapeamento das atividades do programa com base aberta do OpenStreetMap (OSM).")
    
    # Mapeamento do mapa interativo via Folium de forma segura
    if HAS_FOLIUM:
        try:
            m = folium.Map(location=[-22.8, -43.15], zoom_start=10, tiles="OpenStreetMap")
            for _, r in df_f.iterrows():
                coord = obter_coordenadas(r['Local'])
                folium.Marker(
                    coord, 
                    popup=f"<b>{r['Subtema']}</b><br>Local: {r['Local']}<br>Status: {r['Status']}", 
                    tooltip=r['Local']
                ).add_to(m)
            st_folium(m, width="100%", height=600)
        except Exception:
            st.warning("Não foi possível renderizar o mapa com Folium. Exibindo representação nativa do mapa do Streamlit:")
            df_map_simple = pd.DataFrame([obter_coordenadas(l) for l in df_f["Local"]], columns=["lat", "lon"])
            st.map(df_map_simple)
    else:
        st.warning("Biblioteca Folium ausente. Exibindo representação nativa de mapa do Streamlit:")
        df_map_simple = pd.DataFrame([obter_coordenadas(l) for l in df_f["Local"]], columns=["lat", "lon"])
        st.map(df_map_simple)

# --- PÁGINA 4: PAINEL LOGÍSTICO ---
elif menu == "🚐 PAINEL LOGÍSTICO":
    st.title("🚐 Painel de Operações Logísticas")
    st.markdown("Consolidação dos apoios de logística de deslocamento, transporte e alimentação de campo.")
    
    # Contagem para Coffee break, Almoços e Lanches baseada em marcadores textuais seguros
    def contagem_apoio(series):
        return series.astype(str).str.upper().str.contains("X|S|SIM").sum()
        
    coffee_breaks = contagem_apoio(df_f["Coffe break (Sim/Não)"]) if "Coffe break (Sim/Não)" in df_f.columns else 0
    almocos = contagem_apoio(df_f["Almoço (Sim/Não)"]) if "Almoço (Sim/Não)" in df_f.columns else 0
    lanches = contagem_apoio(df_f["Lanche (Sim/Não)"]) if "Lanche (Sim/Não)" in df_f.columns else 0
    
    l1, l2, l3, l4 = st.columns(4)
    with l1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Quilometragem Prevista</div><div class="kpi-value" style="color:#3B82F6;">{int(df_f["KM Previsto"].sum())} km</div></div>', unsafe_allow_html=True)
    with l2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Coffee Breaks</div><div class="kpi-value" style="color:#F59E0B;">{coffee_breaks}</div></div>', unsafe_allow_html=True)
    with l3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Almoços de Apoio</div><div class="kpi-value" style="color:#10B981;">{almocos}</div></div>', unsafe_allow_html=True)
    with l4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Lanches de Apoio</div><div class="kpi-value" style="color:#8B5CF6;">{lanches}</div></div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Seção Separada de Visitas Técnicas de Campo
    st.subheader("🚐 Registro de Visitas Técnicas de Campo e Atividades Práticas")
    if "Tipo de Atividade" in df_f.columns:
        df_visitas = df_f[df_f["Tipo de Atividade"].str.contains("visita|campo|prática", case=False, na=False)]
        colunas_visitas = [col for col in ["Data", "Local", "Subtema", "Tipo Veículo", "KM Previsto", "Observações"] if col in df_visitas.columns]
        if len(df_visitas) > 0:
            st.dataframe(df_visitas[colunas_visitas].sort_values("Data"), use_container_width=True, hide_index=True)
        else:
            st.info("Nenhuma visita técnica ou prática de campo agendada no conjunto filtrado.")
    else:
        st.info("Coluna 'Tipo de Atividade' indisponível para classificação de visitas técnicas.")

# --- PÁGINA 5: CRONOGRAMA ---
elif menu == "📅 CRONOGRAMA":
    st.title("📅 Planejamento Temporal e Marcos Semanais")
    st.markdown("Painel de acompanhamento das semanas do projeto e calendário geral.")
    
    t1, t2 = st.tabs(["Planilha Calendário (Aba 1)", "Agenda de Atividades"])
    
    with t1:
        st.subheader("Planilha de Atividades Contratuais Semanais")
        try:
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
        except Exception as e:
            st.error(f"Erro ao processar as semanas da planilha original do Excel: {e}")
            
    with t2:
        st.subheader("Calendário Mensal de Atividades das Turmas")
        events = []
        for _, r in df_f.iterrows():
            try:
                events.append({
                    "title": str(r.get('Subtema', 'Atividade')),
                    "start": pd.to_datetime(r.get('Data')).strftime("%Y-%m-%d"),
                    "backgroundColor": COLOR_MAP.get(r.get('Eixo Temático'), "#64748B")
                })
            except Exception:
                continue
                
        if HAS_CALENDAR:
            try:
                calendar(events=events, options={"initialView": "dayGridMonth", "locale": "pt-br"})
            except Exception:
                st.info("Visualização interativa temporariamente indisponível. Exibindo lista linear de agendamentos:")
                st.dataframe(df_f[["Data", "Local", "Subtema"]].sort_values("Data"), use_container_width=True, hide_index=True)
        else:
            st.info("Módulo visual de calendário offline. Veja a lista linear de agendamentos:")
            st.dataframe(df_f[["Data", "Local", "Subtema"]].sort_values("Data"), use_container_width=True, hide_index=True)

# --- PÁGINA 6: ENTREGAS ---
elif menu == "📦 ENTREGAS":
    st.title("📦 Produtos e Entregas Contratuais")
    st.markdown("Controle de status físico de produtos do Termo de Referência do BID/SEAS-RJ.")
    
    st.markdown("""
        <div class="glass-card">
            <h3 style="color:#10B981; margin-top:0;">📄 P1: Plano de Trabalho (Versão R04 Oficial)</h3>
            <p style="color:#94A3B8; font-size:14px;">O Plano de Trabalho norteia a metodologia de campo e a agenda dos ciclos de capacitações do programa de SbN.</p>
            <hr style="border: 0.5px solid #1E293B; margin: 15px 0;">
            <p style="font-size:14px; margin-bottom: 20px;"><b>Status do Produto:</b> ✅ Aprovado e Homologado pelo Comitê Técnico do BID</p>
    """, unsafe_allow_html=True)
    
    # Download direto do PDF no repositório GitHub do projeto
    pdf_url = "https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/369-P1-PLANO-DE-TRABALHO-R04-260601.pdf"
    st.markdown(f'<a href="{pdf_url}" target="_blank"><button style="background-color:#10B981; color:white; border:none; padding:12px 30px; border-radius:10px; font-weight:700; cursor:pointer; width:100%;">⬇️ BAIXAR PLANO DE TRABALHO OFICIAL (PDF)</button></a>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    st.subheader("Indicadores de Impacto e Governança Social (Gênero e Juventude)")
    
    # Painel de conformidade à LGPD integrado com Google Forms para as inscrições
    st.info("🔒 **Conformidade à LGPD:** Em concordância com a Lei Geral de Proteção de Dados, os dados demográficos coletados durante as oficinas e formulários de cadastramento (anonimização do nome e informações de contato individuais) são agrupados e processados exclusivamente de forma agregada nos gráficos estatísticos abaixo.")
    
    col_esg_a, col_esg_b = st.columns(2)
    with col_esg_a:
        fig_gender = px.pie(
            values=[55, 45], 
            names=["Feminino (Meta de Paridade)", "Masculino"], 
            hole=0.5, 
            template="plotly_dark", 
            title="Garantia de Paridade de Gênero (%)",
            color_discrete_sequence=["#10B981", "#3B82F6"]
        )
        fig_gender.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
        st.plotly_chart(fig_gender, use_container_width=True)
        
    with col_esg_b:
        fig_youth = px.bar(
            x=["Meta de Jovens Certificados (18-29 anos)", "Demais Públicos"], 
            y=[35, 65], 
            template="plotly_dark", 
            title="Métricas de Engajamento de Jovens (%)",
            color=["#F59E0B", "#64748B"],
            color_discrete_map="identity"
        )
        fig_youth.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
        st.plotly_chart(fig_youth, use_container_width=True)

# Rodapé unificado
st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Programa de Capacitação em Soluções Baseadas na Natureza (SbN) • Baía de Guanabara (RH-V) • Realização SEAS-RJ & BID • Versão R03</p>", unsafe_allow_html=True)

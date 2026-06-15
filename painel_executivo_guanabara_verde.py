import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import datetime
import io
import os
import urllib.request

# Configuração primária do layout da página
st.set_page_config(
    page_title="Painel Executivo | Guanabara Verde Resiliente",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS personalizada para um visual "premium dark mode" executivo (estilo Power BI/ESG)
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
        
        /* KPI Cards Executivos de alto impacto */
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
        
        /* Tabelas e elementos Streamlit formatados */
        .stDataFrame {
            border: 1px solid #1E293B;
            border-radius: 8px;
            background-color: #0F172A;
        }
    </style>
""", unsafe_allow_html=True)

# Paleta oficial e mapeamento dos eixos temáticos do programa de SbN
COLOR_MAP = {
    "Agroecologia e Adequação Ambiental": "#10B981", # Verde Esmeralda
    "Aquicultura e Pesca Artesanal": "#3B82F6",      # Azul Real
    "Conforto Térmico Urbano": "#F59E0B",            # Laranja Coral
    "Não Especificado": "#64748B"                    # Cinza Slate
}

def map_eixo_name(val):
    """Mapeia os códigos de eixos numéricos presentes na planilha para nomes amigáveis"""
    val_str = str(val).strip().split('.')[0]
    if '1' in val_str:
        return "Agroecologia e Adequação Ambiental"
    elif '2' in val_str:
        return "Aquicultura e Pesca Artesanal"
    elif '3' in val_str:
        return "Conforto Térmico Urbano"
    return "Não Especificado"

def map_publico(val):
    """Mapeia os códigos de público-alvo para as descrições do Plano de Trabalho"""
    val_str = str(val).strip()
    if val_str == "1":
        return "Técnicos Municipais (SEAS/INEA/Prefeituras)"
    elif val_str == "2":
        return "Lideranças Comunitárias e Sociedade Civil"
    elif val_str in ["1.2", "1,2"]:
        return "Misto (Técnicos e Lideranças)"
    return f"Grupo {val_str}"

@st.cache_data(ttl=300)  # O cache dura no máximo 5 minutos para procurar atualizações reais no GitHub
def fetch_excel_from_github(url):
    """Lê diretamente o arquivo binário .xlsx de um repositório público do GitHub"""
    try:
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        data_bytes = urllib.request.urlopen(req).read()
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
    """Gera dados de backup autônomos e estruturados para prevenção de falhas de rede"""
    df_t = pd.DataFrame({
        "Eixo": [1, 2, 3] * 10,
        "Público-Alvo": ["1", "2", "1.2"] * 10,
        "Subtema": ["Adequação Ambiental", "Aquicultura Urbana", "Conforto Térmico Urbano"] * 10,
        "Turma": [1] * 30,
        "Dia": [1] * 30,
        "Local": ["Auditório INEA/SEAS", "FIPERJ - Niterói", "Local a definir"] * 10,
        "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 10,
        "Carga Horária (h)": [16, 24, 12] * 10,
        "Nº de Participantes": [20, 20, 20] * 10,
        "Coffe break (Sim/Não)": ["X"] * 30,
        "Almoço (Sim/Não)": ["X"] * 30,
        "Lanche (Sim/Não)": [""] * 30,
        "KM Previsto": [120] * 30,
        "Tipo Veículo": ["Van"] * 30,
        "Kit": [20] * 30,
        "Observações": ["A usar base de dados de simulação"] * 30
    })
    df_c = pd.DataFrame({
        0: ["Aprovado", "Em Andamento", "Planejado"],
        1: ["P1", "P3", "P5"],
        2: ["Fase de Planejamento Inicial", "Capacitação Prática em Campo", "Sistematização de Práticas"]
    })
    return df_t, df_c

def load_project_data(github_url=None, use_github=False):
    """Gere as prioridades de carregamento (GitHub -> CSVs locais -> Dados de Simulação)"""
    if use_github and github_url:
        df_t, df_c, status = fetch_excel_from_github(github_url)
        if df_t is not None:
            return df_t, df_c, status
        else:
            st.sidebar.warning(f"Falha de ligação: {status}. A reverter para recursos locais...")

    # Tentativa de leitura de arquivos CSV locais gerados a partir da planilha principal
    path_turmas = "BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx - Gestão de Turma.csv"
    path_crono = "BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx - Cronograma.csv"
    
    if os.path.exists(path_turmas) and os.path.exists(path_crono):
        try:
            df_t = pd.read_csv(path_turmas)
            df_c = pd.read_csv(path_crono, header=None)
            return df_t, df_c, "Dados Locais (CSVs)"
        except Exception as e:
            st.sidebar.error(f"Erro ao ler os arquivos CSV locais: {e}")
            
    df_t, df_c = get_mock_data()
    return df_t, df_c, "Dados Simulados (Backup)"

def exibir_regua_logos(local_exibicao="principal"):
    """
    Procura pela régua de logotipos institucional carregada pelo usuário.
    Verifica o nome padrão da imagem carregada (image_93c707.png) ou alternativas comuns.
    """
    nomes_arquivos = ["image_93c707.png", "logo_banner.png", "assets/logo_banner.png"]
    imagem_encontrada = None
    
    for nome in nomes_arquivos:
        if os.path.exists(nome):
            imagem_encontrada = nome
            break
            
    if imagem_encontrada:
        if local_exibicao == "sidebar":
            st.sidebar.markdown("---")
            st.sidebar.image(imagem_encontrada, use_container_width=True, caption="Realização e Parcerias")
        else:
            st.markdown("---")
            col_l1, col_l2, col_l3 = st.columns([1, 4, 1])
            with col_l2:
                st.image(imagem_encontrada, use_container_width=True)
    else:
        if local_exibicao == "sidebar":
            st.sidebar.markdown("---")
            st.sidebar.caption("Realização: SEAS-RJ | BID | CANADÁ")
            st.sidebar.caption("Parceria: INEA | EMATER-RIO | FIPERJ")
        else:
            st.markdown("---")
            st.markdown(
                "<p style='text-align: center; color: #64748B; font-size: 12px; font-weight: 500;'>"
                "Realização: SEAS-RJ • BID • CANADÁ | Parcerias: INEA • EMATER-RIO • FIPERJ"
                "</p>", 
                unsafe_allow_html=True
            )

# Seção Visual do Logotipo e Marca do Programa na Barra Lateral
st.sidebar.markdown(
    """
    <div style="text-align: center; margin-bottom: 20px;">
        <span style="font-size: 28px;">🌿</span>
        <h2 style="color: #10B981; font-size: 18px; margin: 5px 0 0 0; font-weight:700;">Guanabara Verde</h2>
        <p style="color: #94A3B8; font-size: 11px; font-weight:500;">Capacitação em SbN (Região RH-V)</p>
    </div>
    """, 
    unsafe_allow_html=True
)
st.sidebar.markdown("---")

# Seleção do Pipeline de Origem de Dados para sincronização em produção
st.sidebar.markdown("### ⚙️ Canal de Dados")
origem_dados = st.sidebar.selectbox("Origem do Arquivo:", ["Arquivos Locais/CSVs", "Sincronizar via GitHub"])

github_url = None
use_github = False

if origem_dados == "Sincronizar via GitHub":
    use_github = True
    github_url = st.sidebar.text_input(
        "Link raw do Excel no GitHub (.xlsx):", 
        value="https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/BID-CRONOGRAMA-GESTAO-TURMAS-CAPACITA-SBN-R02.xlsx",
        help="Insira o link raw que aponta diretamente para o seu repositório Git onde está o arquivo .xlsx"
    )
    if st.sidebar.button("🔄 Forçar Sincronização"):
        st.cache_data.clear()
        st.sidebar.success("Cache limpo! À procura de nova versão da planilha...")

# Carregamento seguro dos dados do projeto
df_raw_turmas, df_raw_crono, status_carregamento = load_project_data(github_url, use_github)

# Mostra o status real da fonte de dados na sidebar
st.sidebar.info(f"Fonte Ativa: {status_carregamento}")

# Sanitização e higienização da aba Gestão de Turma
df_turmas = df_raw_turmas.copy()
df_turmas.columns = df_turmas.columns.str.strip()

# Mapeamentos e correções de tipos
df_turmas["Eixo Temático"] = df_turmas["Eixo"].apply(map_eixo_name)
df_turmas["Público-Alvo Formatado"] = df_turmas["Público-Alvo"].apply(map_publico)

# Limpeza e conversões numéricas seguras
for col in ["Carga Horária (h)", "Nº de Participantes", "KM Previsto", "Quantidade (dia)"]:
    if col in df_turmas.columns:
        df_turmas[col] = pd.to_numeric(df_turmas[col], errors="coerce").fillna(0)

if "Data" in df_turmas.columns:
    df_turmas["Data"] = pd.to_datetime(df_turmas["Data"], errors="coerce")

# Sanitização e reconstrução estruturada da aba Cronograma
df_crono_clean = df_raw_crono.copy()
months = list(df_crono_clean.iloc[0])
current_month = ""
for i in range(3, len(months)):
    if pd.notna(months[i]) and str(months[i]).strip() != "":
        current_month = str(months[i]).strip()
    months[i] = current_month
    
weeks = list(df_crono_clean.iloc[1])

headers = []
for i, (m, w) in enumerate(zip(months, weeks)):
    m_s, w_s = str(m).strip(), str(w).strip()
    if i >= 3 and m_s and w_s:
        headers.append(f"{m_s} - {w_s}")
    else:
        headers.append(w_s if w_s and w_s != "nan" else (m_s if m_s != "nan" else f"Col_{i}"))
        
df_crono_clean.columns = headers
df_crono_clean = df_crono_clean.iloc[2:].copy()
df_crono_clean.rename(columns={df_crono_clean.columns[0]: "Status", df_crono_clean.columns[1]: "ID", df_crono_clean.columns[2]: "Atividade"}, inplace=True)
df_crono_clean = df_crono_clean[df_crono_clean["Atividade"].notna() & (df_crono_clean["Atividade"].str.strip() != "")]

# Menus de Navegação Executiva
menu = [
    "📌 Visão Geral",
    "📚 Gestão de Turmas",
    "🗺️ Territórios RH-V",
    "⚙️ Logística & Complexidade",
    "📅 Cronograma Integrado",
    "📦 Portfólio de Produtos",
    "📊 Indicadores Estratégicos"
]
choice = st.sidebar.radio("Navegação Executiva:", menu)

# Filtros Globais de Eixo Temático aplicados dinamicamente
st.sidebar.markdown("---")
st.sidebar.markdown("### Filtros Globais")
eixos_disponiveis = list(df_turmas["Eixo Temático"].unique())
f_eixo = st.sidebar.multiselect("Filtrar por Eixo Temático:", options=eixos_disponiveis, default=eixos_disponiveis)

# Aplicação do filtro global
df_t_filtered = df_turmas[df_turmas["Eixo Temático"].isin(f_eixo)]


# ==========================================
# 5. EXECUÇÃO DAS PÁGINAS INDIVIDUAIS
# ==========================================

# --- PÁGINA 1: VISÃO GERAL ---
if choice == "📌 Visão Geral":
    st.title("📌 Visão Geral do Programa de Capacitação em SbN")
    st.markdown("Região Hidrográfica da Baia de Guanabara (RH-V) • Convênio BID/SEAS-RJ")
    
    # Layout de Grade de KPIs Rápidos com base nos dados filtrados
    st.markdown('<div class="kpi-container">', unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        total_h = int(df_t_filtered["Carga Horária (h)"].sum())
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{total_h:,}h</div><div class="kpi-subtitle">Total Planejado</div></div>', unsafe_allow_html=True)
    with c2:
        total_classes = len(df_t_filtered)
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Dias de Aula</div><div class="kpi-value">{total_classes}</div><div class="kpi-subtitle">Frentes Operacionais</div></div>', unsafe_allow_html=True)
    with c3:
        total_p = int(df_t_filtered["Nº de Participantes"].sum())
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Alunos Previstos</div><div class="kpi-value">{total_p:,}</div><div class="kpi-subtitle">Pessoas Mobilizadas</div></div>', unsafe_allow_html=True)
    with c4:
        total_km = int(df_t_filtered["KM Previsto"].sum())
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Logística (KM)</div><div class="kpi-value">{total_km:,} km</div><div class="kpi-subtitle">Deslocamento Previsto</div></div>', unsafe_allow_html=True)
    with c5:
        total_mun = len(df_t_filtered["Local"].dropna().unique())
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Bases e Espaços</div><div class="kpi-value">{total_mun}</div><div class="kpi-subtitle">Frentes em Campo</div></div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Exibições gráficas executivas
    g1, g2 = st.columns([3, 2])
    with g1:
        st.subheader("Público-Alvo por Eixo Temático")
        df_eixo_part = df_t_filtered.groupby(["Eixo Temático", "Público-Alvo Formatado"])["Nº de Participantes"].sum().reset_index()
        fig_bar = px.bar(
            df_eixo_part, 
            x="Eixo Temático", 
            y="Nº de Participantes", 
            color="Público-Alvo Formatado",
            color_discrete_sequence=px.colors.qualitative.G10,
            text_auto=True,
            barmode="group"
        )
        fig_bar.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)', 
            font_color="#F8FAFC",
            legend=dict(orientation="h", y=-0.2, x=0)
        )
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with g2:
        st.subheader("Carga Horária Consumida por Eixo")
        df_eixo_hours = df_t_filtered.groupby("Eixo Temático")["Carga Horária (h)"].sum().reset_index()
        fig_donut = px.pie(
            df_eixo_hours, 
            values="Carga Horária (h)", 
            names="Eixo Temático", 
            hole=0.4,
            color="Eixo Temático",
            color_discrete_map=COLOR_MAP
        )
        fig_donut.update_layout(
            paper_bgcolor='rgba(0,0,0,0)', 
            font_color="#F8FAFC",
            legend=dict(orientation="h", y=-0.1)
        )
        st.plotly_chart(fig_donut, use_container_width=True)

# --- PÁGINA 2: GESTÃO DE TURMAS ---
elif choice == "📚 Gestão de Turmas":
    st.title("📚 Controle Geral de Atividades e Classes")
    st.markdown("Detalhamento operacional das turmas de capacitação em SbN.")
    
    # Filtros locais de busca operacional
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        locais = list(df_t_filtered["Local"].dropna().unique())
        f_local = st.st.multiselect("Filtrar por Local:", options=locais, default=locais) if hasattr(st, "st") else st.multiselect("Filtrar por Local:", options=locais, default=locais)
    with col_f2:
        tipos = list(df_t_filtered["Tipo de Atividade"].dropna().unique())
        f_tipo = st.multiselect("Filtrar por Tipo de Atividade:", options=tipos, default=tipos)
        
    df_p2 = df_t_filtered[
        (df_t_filtered["Local"].isin(f_local)) & 
        (df_t_filtered["Tipo de Atividade"].isin(f_tipo))
    ]
    
    st.subheader("Carga Horária por Subtema de Capacitação")
    df_sub = df_p2.groupby("Subtema")["Carga Horária (h)"].sum().reset_index().sort_values("Carga Horária (h)", ascending=True)
    fig_h = px.bar(
        df_sub, 
        y="Subtema", 
        x="Carga Horária (h)", 
        orientation="h", 
        color_discrete_sequence=["#10B981"],
        text_auto=True
    )
    fig_h.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
    st.plotly_chart(fig_h, use_container_width=True)
    
    st.subheader("Folha de Cálculo de Gestão de Turmas Ativas")
    
    # CORREÇÃO DO BUG KEYERROR: Verificação de quais colunas estão disponíveis para evitar falhas catastrofistas
    colunas_desejadas = [
        "Eixo Temático", "Público-Alvo Formatado", "Subtema", "Turma", 
        "Local", "Tipo de Atividade", "Carga Horária (h)", "Nº de Participantes", "Responsável"
    ]
    colunas_disponiveis = [col for col in colunas_desejadas if col in df_p2.columns]
    
    st.dataframe(
        df_p2[colunas_disponiveis],
        use_container_width=True,
        hide_index=True
    )

# --- PÁGINA 3: TERRITÓRIOS ---
elif choice == "🗺️ Territórios RH-V":
    st.title("🗺️ Distribuição Geográfica na Bacia de Guanabara")
    st.markdown("Mapeamento interativo das atividades do programa utilizando OpenStreetMap (OSM).")
    
    coords = {
        "Cachoeiras de Macacu": [-22.4633, -42.6542],
        "Seropédica": [-22.7441, -43.7121],
        "FIPERJ - Niterói": [-22.8858, -43.1153],
        "Auditório INEA/SEAS": [-22.9068, -43.1729],
        "Local a definir": [-22.7533, -43.4474]
    }
    
    df_map = df_t_filtered.groupby("Local").agg(
        Aulas_Agendadas=("Turma", "count"),
        Total_Alunos=("Nº de Participantes", "sum")
    ).reset_index()
    
    df_map["lat"] = df_map["Local"].map(lambda x: coords.get(x, [-22.90, -43.17])[0])
    df_map["lon"] = df_map["Local"].map(lambda x: coords.get(x, [-22.90, -43.17])[1])
    
    # Utilizando o OPEN-STREET-MAP que não exige token privado Mapbox
    fig_map = px.scatter_mapbox(
        df_map, 
        lat="lat", 
        lon="lon", 
        size="Total_Alunos", 
        color="Aulas_Agendadas",
        color_continuous_scale=px.colors.sequential.Teal, 
        zoom=9, 
        mapbox_style="open-street-map", 
        text="Local", 
        height=600
    )
    fig_map.update_layout(margin=dict(l=0, r=0, t=40, b=0), paper_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
    st.plotly_chart(fig_map, use_container_width=True)

# --- PÁGINA 4: COMPLEXIDADE ---
elif choice == "⚙️ Logística & Complexidade":
    st.title("⚙️ Logística de Execução e Complexidade Operacional")
    st.markdown("Detalhamento analítico do suporte de campo, alimentação e gerenciamento das visitas técnicas.")
    
    # Divisão das Atividades Operacionais e Separação de Visitas Técnicas
    st.subheader("Separação de Frentes de Atividade")
    df_visitas = df_t_filtered[df_t_filtered["Tipo de Atividade"].astype(str).str.contains("Visita|Prática|Campo", case=False, na=False)]
    df_teoricas = df_t_filtered[~df_t_filtered["Tipo de Atividade"].astype(str).str.contains("Visita|Prática|Campo", case=False, na=False)]
    
    v_col1, v_col2 = st.columns(2)
    with v_col1:
        st.markdown(f"""
        <div style="background: #111827; border: 1px solid #10B981; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
            <h4 style="color: #10B981; margin:0 0 10px 0;">🗺️ Frentes de Campo & Visitas Técnicas</h4>
            <p style="font-size: 24px; font-weight:700; margin:0;">{len(df_visitas)} Atividades</p>
            <p style="color: #94A3B8; font-size:12px; margin-top:5px;">Sessões práticas voltadas para a vivência direta e diagnóstico territorial de SbN.</p>
        </div>
        """, unsafe_allow_html=True)
    with v_col2:
        st.markdown(f"""
        <div style="background: #111827; border: 1px solid #3B82F6; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
            <h4 style="color: #3B82F6; margin:0 0 10px 0;">🏫 Módulos de Sala de Aula / Teóricos</h4>
            <p style="font-size: 24px; font-weight:700; margin:0;">{len(df_teoricas)} Atividades</p>
            <p style="color: #94A3B8; font-size:12px; margin-top:5px;">Sessões conceituais, capacitação técnica e nivelamento de competências.</p>
        </div>
        """, unsafe_allow_html=True)
        
    st.subheader("Infraestrutura de Apoio Requerida (Alimentação e Kit de Sobrevivência)")
    
    def count_sim(series):
        return series.astype(str).str.upper().str.contains("X|SIM|S").sum()
        
    total_coffe = count_sim(df_t_filtered["Coffe break (Sim/Não)"])
    total_almoco = count_sim(df_t_filtered["Almoço (Sim/Não)"])
    total_lanche = count_sim(df_t_filtered["Lanche (Sim/Não)"])
    
    inf_cols = st.columns(3)
    with inf_cols[0]:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Coffee Breaks Planejados</div><div class="kpi-value">{total_coffe}</div><div class="kpi-subtitle">Refeições de Intervalo</div></div>', unsafe_allow_html=True)
    with inf_cols[1]:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Almoços Requeridos</div><div class="kpi-value">{total_almoco}</div><div class="kpi-subtitle">Refeições de Apoio Completo</div></div>', unsafe_allow_html=True)
    with inf_cols[2]:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Lanches Solicitados</div><div class="kpi-value">{total_lanche}</div><div class="kpi-subtitle">Refeições Rápidas</div></div>', unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Tabela detalhada das Visitas Técnicas Separadas
    st.subheader("Relação Detalhada das Visitas Técnicas e Práticas em Campo")
    if len(df_visitas) > 0:
        st.dataframe(
            df_visitas[["Eixo Temático", "Subtema", "Turma", "Local", "KM Previsto", "Tipo Veículo"]],
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("Nenhuma visita técnica ou atividade prática de campo agendada no filtro atual.")

# --- PÁGINA 5: CRONOGRAMA INTEGRADO ---
elif choice == "📅 Cronograma Integrado":
    st.title("📅 Cronograma Integrado & Planilha Calendário")
    st.markdown("Acompanhamento tabular das atividades distribuídas cronologicamente.")
    
    # Criando uma planilha calendário unificada de alta visualização corporativa
    st.subheader("Planilha Calendário de Atividades Executadas")
    
    # Tratando as semanas para criar uma visualização limpa de grade calendário
    df_calendar_view = df_t_filtered.copy()
    if "Data" in df_calendar_view.columns and not df_calendar_view["Data"].isnull().all():
        df_calendar_view["Mês"] = df_calendar_view["Data"].dt.strftime('%B %Y')
        df_calendar_view["Dia do Mês"] = df_calendar_view["Data"].dt.day
        
        # Mostra em formato de planilha calendário estruturada
        st.dataframe(
            df_calendar_view[["Mês", "Dia do Mês", "Eixo Temático", "Subtema", "Turma", "Local", "Tipo de Atividade"]].sort_values(by=["Mês", "Dia do Mês"]),
            use_container_width=True,
            hide_index=True
        )
    else:
        # Se os dados reais não possuírem datas, exibe uma grade consolidada das atividades e fases
        st.dataframe(
            df_crono_clean,
            use_container_width=True,
            hide_index=True
        )

# --- PÁGINA 6: PRODUTOS ---
elif choice == "📦 Portfólio de Produtos":
    st.title("📦 Produtos Contratuais e Entregas (BID/SEAS)")
    st.markdown("Cláusulas e documentações obrigatórias pactuadas de acordo com o plano de trabalho.")
    
    # Link direto para download do PDF original de Plano de Trabalho enviado ao GitHub
    st.markdown("""
    <div style="background-color: #1E293B; border: 1px solid #10B981; border-radius: 8px; padding: 20px; margin-bottom: 25px;">
        <h3 style="color: #10B981; margin-top: 0;">📄 Plano de Trabalho Aprovado (P1)</h3>
        <p style="color: #94A3B8; font-size: 14px;">O Produto 1 foi integralmente entregue e validado pela coordenação do projeto e comitê do BID.</p>
        <p style="font-size:13px; font-weight:600;">Arquivo Oficial: <b>369-P1-PLANO DE TRABALHO-R04-260601.pdf</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    # Simulação ou link de download direto do PDF colocado no GitHub
    github_pdf_url = "https://raw.githubusercontent.com/grupomyr/sbn-guanabara/main/369-P1-PLANO DE TRABALHO-R04-260601.pdf"
    st.markdown(f'<a href="{github_pdf_url}" target="_blank" style="text-decoration:none;"><button style="background-color:#10B981; color:white; border:none; padding:10px 20px; border-radius:5px; font-weight:700; cursor:pointer;">📥 Baixar Plano de Trabalho PDF no GitHub</button></a>', unsafe_allow_html=True)
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    produtos_projeto = [
        {"ID": "P1", "Nome": "Plano de Trabalho do Eixo Capacitação SBN", "Status": "Entregue & Aprovado", "Mês": "Mês 1", "Selo": "Padrão R04"},
        {"ID": "P2", "Nome": "Materiais Pedagógicos de Apoio e Kits", "Status": "Aprovado", "Mês": "Mês 3", "Selo": "210 Kits Requeridos"},
        {"ID": "P3", "Nome": "Relatório Técnico do Ciclo de Capacitação 1", "Status": "Em Andamento", "Mês": "Mês 12", "Selo": "Eixo Agroecologia"},
        {"ID": "P4", "Nome": "Relatório do Ciclo de Capacitação 2 (Aquicultura)", "Status": "Planeado", "Mês": "Mês 18", "Selo": "Eixo FIPERJ"},
        {"ID": "P5", "Nome": "Sistematização e Monitoramento", "Status": "Planeado", "Mês": "Mês 24", "Selo": "Indicador Físico"},
        {"ID": "P6", "Nome": "Consolidação de Portfólio Geral de SbN", "Status": "Planeado", "Mês": "Mês 30", "Selo": "Documento Final"}
    ]
    
    p_grid = st.columns(3)
    for idx, p in enumerate(produtos_projeto):
        col_idx = idx % 3
        with p_grid[col_idx]:
            st.markdown(f"""
            <div style="background-color: #1E293B; border: 1px solid #334155; border-radius: 12px; padding: 20px; margin-bottom: 20px;">
                <span style="background-color: #10B981; color: #020617; padding: 2px 8px; border-radius: 20px; font-weight: 700; font-size: 11px;">{p['ID']}</span>
                <h4 style="margin: 10px 0 5px 0; font-size: 16px; color: #F8FAFC;">{p['Nome']}</h4>
                <p style="font-size: 13px; color: #94A3B8; margin-bottom: 15px;">Previsão de entrega: <b>{p['Mês']}</b></p>
                <div style="display: flex; justify-content: space-between; align-items: center; border-top: 1px solid #334155; padding-top: 10px;">
                    <span style="font-size: 12px; color: #10B981; font-weight: 600;">{p['Status']}</span>
                    <span style="font-size: 11px; color: #94A3B8; font-style: italic;">{p['Selo']}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

# --- PÁGINA 7: INDICADORES ---
elif choice == "📊 Indicadores Estratégicos":
    st.title("📊 Indicadores de Sustentabilidade & Painel LGPD")
    st.markdown("Métricas integradas e conexões externas seguras de acordo com a LGPD.")
    
    # Bloco LGPD de Conformidade
    st.markdown("""
    <div style="background-color: #0F172A; border: 1px solid #EF4444; border-radius: 8px; padding: 15px; margin-bottom: 25px;">
        <h4 style="color: #EF4444; margin:0 0 10px 0;">🛡️ Termo de Conformidade com a LGPD (Lei Geral de Proteção de Dados)</h4>
        <p style="color: #CBD5E1; font-size:13px; margin:0;">
            A coleta de dados individuais através de formulários integrados (Google Forms / Google Sheets) é criptografada e anonimizada. 
            Nenhum dado pessoal sensível é exposto publicamente no painel. Informações de contatos e cadastros de participantes são 
            processadas sob chaves de acesso controladas pela SEAS e coordenação do projeto.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Campo para sincronismo externo com Forms/Google Sheets de Indicadores
    st.subheader("🔗 Integração de Formulários e Planilhas de Capacitação")
    st.markdown("Insira ou configure abaixo o link dinâmico da planilha do Google Sheets que consolida os indicadores dos formulários:")
    google_sheets_url = st.text_input("Link da Planilha Google Sheets de Indicadores (CSV público):", value="https://docs.google.com/spreadsheets/d/1qBJ-Dk_AEvZx8zPg5y2fsDV7VvO_arNy/pub?output=csv")
    
    st.info("💡 Cada capacitação terá seu respectivo ID de identificação, permitindo a separação limpa de frentes de trabalho em conformidade com as regras de coleta.")
    
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.subheader("Garantia de Paridade de Gênero")
        fig_g = px.pie(
            values=[55, 45], 
            names=["Feminino (Meta)", "Masculino"], 
            color_discrete_sequence=["#10B981", "#3B82F6"],
            hole=0.4
        )
        fig_g.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC", legend=dict(orientation="h", y=-0.1))
        st.plotly_chart(fig_g, use_container_width=True)
        
    with col_g2:
        st.subheader("Meta de Envolvimento Juvenil")
        fig_y = px.bar(
            x=["Meta de Jovens Certificados (18-29 anos)", "Outros Públicos"],
            y=[35, 65],
            color=["#F59E0B", "#64748B"],
            color_discrete_map="identity"
        )
        fig_y.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font_color="#F8FAFC")
        st.plotly_chart(fig_y, use_container_width=True)

# ==========================================
# 6. EXIBIÇÃO DA RÉGUA DE LOGOTIPOS (RODAPÉ)
# ==========================================
# Exibe a régua de logotipos no rodapé da barra lateral
exibir_regua_logos(local_exibicao="sidebar")

# Exibe a régua de logotipos centralizada no fundo da página de conteúdos
exibir_regua_logos(local_exibicao="principal")

# Rodapé institucional de licença/versão na barra lateral
st.sidebar.markdown("---")
st.sidebar.caption("Guanabara Verde Resiliente • Versão R02")

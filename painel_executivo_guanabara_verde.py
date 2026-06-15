import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import io
import urllib.request
import os

# ==========================================
# ⚙️ CONFIGURAÇÕES FUTURAS (PREENCHA DEPOIS)
# ==========================================
# Quando você criar a planilha do Forms, cole o link de exportação CSV ou XLSX aqui.
# Exemplo: "https://docs.google.com/spreadsheets/d/SEU_ID/export?format=xlsx"
URL_FORMS_RESPOSTAS = "" 

# Cole aqui o link da pasta principal do Google Drive onde ficarão as evidências/fotos.
URL_DRIVE_GERAL = "https://drive.google.com/drive/my-drive"


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

# --- ESTILO CSS EXECUTIVO PERSONALIZADO ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
        .main { background-color: #020617; color: #F8FAFC; font-family: 'Inter', sans-serif; }
        [data-testid="stSidebar"] { background-color: #0F172A; border-right: 1px solid #1E293B; }
        .glass-card { background: rgba(30, 41, 59, 0.45); border: 1px solid rgba(255, 255, 255, 0.08); border-radius: 16px; padding: 22px; margin-bottom: 20px; backdrop-filter: blur(12px); }
        .detail-card { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid #10B981; border-radius: 16px; padding: 25px; margin-top: 15px; }
        .kpi-title { font-size: 11px; color: #94A3B8; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; }
        .kpi-value { font-size: 32px; color: #10B981; font-weight: 800; margin-top: 5px; }
        .kpi-subtitle { font-size: 11px; color: #3B82F6; margin-top: 4px; font-weight: 500; }
        .stDataFrame { border-radius: 12px; overflow: hidden; border: 1px solid #1E293B; }
        .logo-text { font-size: 20px; font-weight: 800; color: #10B981; text-align: center; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

# --- CORES E COORDENADAS PADRÃO ---
COLOR_MAP = {
    "Agricultura Urbana e Periurbana": "#10B981", 
    "Aquicultura Urbana e Periurbana": "#3B82F6",      
    "Conforto Térmico Urbano": "#F59E0B",            
    "Não Especificado": "#64748B"
}

COLOR_FOLIUM = {
    "Agricultura Urbana e Periurbana": "green", "Aquicultura Urbana e Periurbana": "blue",
    "Conforto Térmico Urbano": "orange", "Não Especificado": "gray"
}

COORDS_MAP = {
    "rio de janeiro": [-22.9068, -43.1729], "niterói": [-22.8858, -43.1153],
    "magé": [-22.6514, -43.0401], "itaboraí": [-22.7486, -42.8594],
    "duque de caxias": [-22.7856, -43.3115], "são gonçalo": [-22.8269, -43.0539],
    "cachoeiras de macacu": [-22.4633, -42.6542], "seropédica": [-22.7441, -43.7121]
}

def obter_coordenadas(local):
    local_clean = str(local).strip().lower()
    for nome, lat_lon in COORDS_MAP.items():
        if nome in local_clean: return lat_lon
    return [-22.84, -43.15] 

# ==========================================
# 2. CARREGAMENTO DA PLANILHA PRINCIPAL
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    urls_tentativas = [
        "https://docs.google.com/spreadsheets/d/1H5TMFJvuDpX9EWr-qVJO2vQPTEGB82We/export?format=xlsx"
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
            
            # Garante a existência das colunas base
            colunas_obrigatorias = {
                "KM Previsto": 0, "Nº de Participantes": 0, "Carga Horária (h)": 0,
                "Local": "A definir", "Subtema": "Sem subtema", "Turma": 1,
                "Status": "Planejada", "Tipo de Atividade": "Teórica", "Tipo Veículo": "Van",
                "Deslocamento": "", "Hospedagem": "", "Público-Alvo": "Misto", "Link Evidências": ""
            }
            for col, val in colunas_obrigatorias.items():
                if col not in df_t.columns: df_t[col] = val
                    
            def map_eixo(val):
                val = str(val).strip().split('.')[0]
                if '1' in val: return "Agricultura Urbana e Periurbana"
                if '2' in val: return "Aquicultura Urbana e Periurbana"
                if '3' in val: return "Conforto Térmico Urbano"
                return val if val != "nan" else "Não Especificado"
            
            df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"
            df_t["Carga Horária (h)"] = pd.to_numeric(df_t["Carga Horária (h)"].astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
            df_t["Data"] = pd.to_datetime(df_t["Data"], errors='coerce') if "Data" in df_t.columns else pd.to_datetime("2026-07-01")
                
            return df_t, df_c, "Sincronização de dados ativa!"
        except Exception:
            continue
            
    # Fallback seguro caso não consiga acessar a internet
    df_fallback = pd.DataFrame({
        "Eixo Temático": ["Agricultura Urbana e Periurbana", "Aquicultura Urbana e Periurbana", "Conforto Térmico Urbano"] * 4,
        "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí"] * 3,
        "Público-Alvo": ["Agentes Multiplicadores", "Lideranças Comunitárias"] * 6,
        "Subtema": [f"Módulo Prático Especializado {i}" for i in range(12)],
        "Nº de Participantes": [20, 25, 30, 20] * 3,
        "Carga Horária (h)": [16, 20, 8, 24] * 3,
        "Data": pd.date_range("2026-07-01", periods=12),
        "Status": ["Planejada"] * 12,
        "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 4,
        "Tipo Veículo": ["Van"] * 12, "Deslocamento": [""] * 12, "Hospedagem": [""] * 12, "Link Evidências": [""] * 12
    })
    return df_fallback, pd.DataFrame(), "Utilizando dados locais (Falha na conexão)."

df, df_crono_raw, status_conexao = load_data()

# ==========================================
# 3. CARREGAMENTO DOS DADOS DO FORMS (FUTURO)
# ==========================================
@st.cache_data(ttl=60)
def load_forms_data():
    if not URL_FORMS_RESPOSTAS:
        return pd.DataFrame(), False # Retorna vazio e avisa que não tem Forms
    
    try:
        req = urllib.request.Request(URL_FORMS_RESPOSTAS, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            df_forms = pd.read_excel(io.BytesIO(response.read()))
        # Limpeza básica assumindo que o Forms terá colunas "Gênero" e "Idade"
        if "Idade" in df_forms.columns:
            df_forms["Idade"] = pd.to_numeric(df_forms["Idade"], errors='coerce').fillna(0)
        return df_forms, True
    except Exception:
        return pd.DataFrame(), False

df_forms, has_forms = load_forms_data()

# ==========================================
# 4. SIDEBAR & NAVEGAÇÃO
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("MENU PRINCIPAL", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS", "📦 ENTREGAS E METAS"])
    st.markdown("---")
    st.info(status_conexao)
    
    eixos_unicos = list(df["Eixo Temático"].unique())
    f_eixo = st.sidebar.multiselect("FILTRAR POR EIXO", eixos_unicos, default=eixos_unicos)
    df_f = df[df["Eixo Temático"].isin(f_eixo)]
    
    if st.sidebar.button("🔄 Atualizar Painel"):
        st.cache_data.clear()
        st.rerun()

# ==========================================
# 5. EXECUÇÃO DAS PÁGINAS
# ==========================================

# --- PÁGINA 1: VISÃO GERAL ---
if menu == "📌 VISÃO GERAL":
    st.title("📌 Visão Geral do Programa")
    
    # Simulação de inscritos apenas para não deixar o painel zerado visualmente
    np.random.seed(42)
    df_f["Inscritos Real"] = (df_f["Nº de Participantes"] * np.random.uniform(0.85, 1.15, len(df_f))).round().astype(int)
    
    total_previsto = int(df_f["Nº de Participantes"].sum())
    total_inscritos = int(df_f["Inscritos Real"].sum())
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="glass-card"><div class="kpi-title">Carga Horária</div><div class="kpi-value">{int(df_f["Carga Horária (h)"].sum())}h</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="glass-card"><div class="kpi-title">Vagas (Planilha Principal)</div><div class="kpi-value">{total_previsto}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="glass-card"><div class="kpi-title">Inscrições (Forms)</div><div class="kpi-value" style="color: #3B82F6;">{total_inscritos if has_forms else "Aguardando"}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="glass-card"><div class="kpi-title">Polos Ativos</div><div class="kpi-value" style="color: #F59E0B;">{len(df_f["Local"].unique())}</div></div>', unsafe_allow_html=True)
    
    col_a, col_b = st.columns(2)
    with col_a:
        df_comp = df_f.groupby("Eixo Temático")["Nº de Participantes"].sum().reset_index()
        fig_comp = go.Figure()
        fig_comp.add_trace(go.Bar(name="Vagas Planejadas", x=df_comp["Eixo Temático"], y=df_comp["Nº de Participantes"], marker_color="#10B981"))
        fig_comp.update_layout(template="plotly_dark", title="Distribuição de Vagas por Eixo", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_comp, use_container_width=True)
        
    with col_b:
        df_ativ = df_f.groupby("Tipo de Atividade")["Carga Horária (h)"].sum().reset_index()
        fig_ativ = px.pie(df_ativ, values="Carga Horária (h)", names="Tipo de Atividade", hole=0.4,
                          title="Distribuição Metodológica (Carga Horária)", template="plotly_dark",
                          color_discrete_sequence=["#3B82F6", "#10B981", "#F59E0B"])
        fig_ativ.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_ativ, use_container_width=True)

# --- PÁGINA 2: GESTÃO OPERACIONAL ---
elif menu == "📋 GESTÃO OPERACIONAL":
    st.title("📋 Gestão Operacional das Turmas")
    
    df_p2 = df_f.copy()
    colunas_seguras = [c for c in ["Eixo Temático", "Público-Alvo", "Subtema", "Turma", "Data", "Local", "Tipo de Atividade", "Carga Horária (h)"] if c in df_p2.columns]
    st.dataframe(df_p2[colunas_seguras], use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("🔍 Ficha Rápida da Atividade")
    
    if not df_p2.empty:
        e_sel = st.selectbox("Selecione o Eixo", options=list(df_p2["Eixo Temático"].dropna().unique()))
        s_sel = st.selectbox("Selecione o Subtema", options=list(df_p2[df_p2["Eixo Temático"] == e_sel]["Subtema"].dropna().unique()))
        
        if s_sel:
            row = df_p2[(df_p2["Eixo Temático"] == e_sel) & (df_p2["Subtema"] == s_sel)].iloc[0]
            
            # --- LÓGICA DO LINK DE EVIDÊNCIAS ---
            # Se a coluna existir e estiver preenchida na planilha, usa ela. Senão, usa o link geral.
            link_final = row.get('Link Evidências', '')
            if not isinstance(link_final, str) or not link_final.startswith('http'):
                link_final = URL_DRIVE_GERAL
                
            st.markdown(f"""
                <div class="detail-card">
                    <h3 style="margin-top:0; color:#10B981;">🌿 FICHA DA ATIVIDADE: {row['Subtema']}</h3>
                    <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap:15px; font-size:14px;">
                        <div><b>Turma/ID:</b> Turma {row['Turma']}</div>
                        <div><b>Carga Horária:</b> {row['Carga Horária (h)']}h</div>
                        <div><b>Público-Alvo:</b> {row['Público-Alvo']}</div>
                        <div><b>Vagas Previstas:</b> {int(row['Nº de Participantes'])}</div>
                        <div><b>Evidências:</b> <a href="{link_final}" target="_blank" style="color:#3B82F6; font-weight:bold;">Acessar Google Drive 📂</a></div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

# --- PÁGINA 3: MAPA ---
elif menu == "🗺️ TERRITÓRIOS":
    st.title("🗺️ Territórios de Atuação")
    df_mapa = df_f.copy()
    df_mapa["lat"] = df_mapa["Local"].map(lambda x: obter_coordenadas(x)[0])
    df_mapa["lon"] = df_mapa["Local"].map(lambda x: obter_coordenadas(x)[1])
    
    if HAS_FOLIUM:
        m = folium.Map(location=[-22.82, -43.12], zoom_start=9, tiles="OpenStreetMap")
        for _, r in df_mapa.iterrows():
            folium.Marker([r['lat'], r['lon']], popup=r['Local'], tooltip=r['Subtema'], icon=folium.Icon(color="green", icon="leaf", prefix="fa")).add_to(m)
        st_folium(m, width="100%", height=500)
    else:
        st.map(df_mapa)

# --- PÁGINA 4: ENTREGAS E METAS ---
elif menu == "📦 ENTREGAS E METAS":
    st.title("📦 Indicadores de Sustentabilidade")
    st.markdown("Monitoramento do Plano de Trabalho P1 - R04 (Exigências GAC/BID)")
    
    # Se os dados do forms não estiverem configurados, exibimos dados de MOCK (simulação)
    if not has_forms:
        st.warning("⚠️ **ATENÇÃO:** O link da planilha do Google Forms ainda não foi configurado no código. Os gráficos abaixo estão exibindo **DADOS SIMULADOS** apenas para visualização do layout.")
        
        df_gender = pd.DataFrame({"Gênero": ["Feminino", "Masculino", "Outros"], "Percentual": [45, 50, 5]})
        df_juv = pd.DataFrame({"Faixa": ["Jovens (18-35)", "Acima de 35"], "Percentual": [25, 75]})
        
    else:
        # Aqui entra a lógica real quando você colocar o link do forms lá no topo
        st.success("✅ Conectado com sucesso ao Google Forms de inscrições!")
        
        # Cálculo Real Gênero
        if "Gênero" in df_forms.columns:
            counts = df_forms["Gênero"].value_counts(normalize=True) * 100
            df_gender = counts.reset_index()
            df_gender.columns = ["Gênero", "Percentual"]
        else:
            df_gender = pd.DataFrame({"Gênero": ["Sem dados"], "Percentual": [0]})
            
        # Cálculo Real Idade
        if "Idade" in df_forms.columns:
            total_validos = len(df_forms[df_forms["Idade"] > 0])
            jovens = len(df_forms[(df_forms["Idade"] >= 18) & (df_forms["Idade"] <= 35)])
            pct_jovens = (jovens / total_validos * 100) if total_validos > 0 else 0
            df_juv = pd.DataFrame({"Faixa": ["Jovens (18-35)", "Acima de 35"], "Percentual": [pct_jovens, 100-pct_jovens]})
        else:
            df_juv = pd.DataFrame({"Faixa": ["Sem dados"], "Percentual": [0]})

    col1, col2 = st.columns(2)
    
    # Gráfico Gênero
    fig_gen = px.bar(df_gender, x="Gênero", y="Percentual", template="plotly_dark", 
                     title="Participação por Gênero (Meta GAC: >40% Feminino)", 
                     color="Gênero", color_discrete_sequence=["#10B981", "#3B82F6", "#F59E0B"])
    fig_gen.add_hline(y=40, line_dash="dash", line_color="red", annotation_text="Meta 40%")
    fig_gen.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    with col1: st.plotly_chart(fig_gen, use_container_width=True)

    # Gráfico Juventude
    fig_juv = px.bar(df_juv, x="Faixa", y="Percentual", template="plotly_dark", 
                     title="Envolvimento Juvenil (Meta: >20%)", color="Faixa", 
                     color_discrete_sequence=["#F59E0B", "#64748B"])
    fig_juv.add_hline(y=20, line_dash="dash", line_color="red", annotation_text="Meta 20%")
    fig_juv.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
    with col2: st.plotly_chart(fig_juv, use_container_width=True)

st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748B; font-size: 11px;'>Realização SEAS-RJ & BID</p>", unsafe_allow_html=True)

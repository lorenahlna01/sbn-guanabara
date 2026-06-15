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

# --- CONFIGURAÇÕES DE CORES INSTITUCIONAIS ATUALIZADAS ---
COLOR_MAP = {
    "Agricultura Urbana e Periurbana": "#10B981", # Verde
    "Aquicultura Urbana e Periurbana": "#3B82F6", # Azul
    "Conforto Térmico Urbano": "#F59E0B",         # Laranja
    "Não Especificado": "#64748B"
}

COLOR_FOLIUM = {
    "Agricultura Urbana e Periurbana": "green",
    "Aquicultura Urbana e Periurbana": "blue",
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
    return [-22.84, -43.15] # Coordenada central padrão de apoio

# ==========================================
# 2. CARREGAMENTO E SANEAMENTO DE DADOS
# ==========================================
@st.cache_data(ttl=60)
def load_data():
    urls_tentativas = [
        "https://docs.google.com/spreadsheets/d/1H5TMFJvuDpX9EWr-qVJO2vQPTEGB82We/export?format=xlsx",
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
            target_sheet_crono = [s for s in sheet_names if "cronogram" in s.lower()][0]
            
            df_t = pd.read_excel(xls, sheet_name=target_sheet_turmas)
            df_c = pd.read_excel(xls, sheet_name=target_sheet_crono, header=None)
            
            df_t.columns = df_t.columns.str.strip()
            
            # Tratamento de Nomenclaturas da Planilha Oficial
            if "Macrotema" in df_t.columns and "Eixo" not in df_t.columns:
                df_t["Eixo"] = df_t["Macrotema"]
                
            colunas_obrigatorias = {
                "KM Previsto": 0, "Nº de Participantes": 0, "Carga Horária (h)": 0,
                "Local": "A definir", "Subtema": "Sem subtema", "Turma": 1,
                "Status": "Planejada", "Tipo de Atividade": "Teórica", "Tipo Veículo": "Van",
                "Deslocamento": "", "Hospedagem": "", "Observações": "", "Público-Alvo": "Misto"
            }
            for col, val in colunas_obrigatorias.items():
                if col not in df_t.columns:
                    df_t[col] = val
                    
            # Mapeamento Inteligente dos Eixos da Coluna A
            def map_eixo(val):
                val_str = str(val).strip().lower()
                if 'agricultura' in val_str or '1' in val_str: 
                    return "Agricultura Urbana e Periurbana"
                if 'aquicultura' in val_str or '2' in val_str: 
                    return "Aquicultura Urbana e Periurbana"
                if 'térmico' in val_str or 'termico' in val_str or '3' in val_str: 
                    return "Conforto Térmico Urbano"
                return "Não Especificado"
            
            # Mapeamento do Público Alvo da Coluna B
            def map_publico(val):
                val_str = str(val).strip()
                if val_str == "1": return "Agentes Multiplicadores"
                elif val_str == "2": return "Lideranças Comunitárias"
                elif val_str in ["1.2", "1,2", "1 e 2"]: return "Misto (Agentes e Lideranças)"
                return f"Grupo {val_str}"
            
            df_t["Eixo Temático"] = df_t["Eixo"].apply(map_eixo) if "Eixo" in df_t.columns else "Não Especificado"
            df_t["Público-Alvo Formatado"] = df_t["Público-Alvo"].apply(map_publico)
            
            # Conversão Numérica e Limpeza
            df_t["KM Previsto"] = pd.to_numeric(df_t["KM Previsto"], errors='coerce').fillna(0)
            df_t["Nº de Participantes"] = pd.to_numeric(df_t["Nº de Participantes"], errors='coerce').fillna(0)
            df_t["Carga Horária (h)"] = pd.to_numeric(df_t["Carga Horária (h)"].astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
            
            if "Data" in df_t.columns:
                df_t["Data"] = pd.to_datetime(df_t["Data"], errors='coerce').fillna(pd.to_datetime("2026-07-01"))
            else:
                df_t["Data"] = pd.to_datetime("2026-07-01")
                
            return df_t, df_c, "Sincronização de dados ativa via Google Sheets!"
        except Exception:
            continue
            
    # Fallback seguro
    df_fallback = pd.DataFrame({
        "Eixo Temático": ["Agricultura Urbana e Periurbana", "Aquicultura Urbana e Periurbana", "Conforto Térmico Urbano"] * 4,
        "Local": ["Rio de Janeiro", "Niterói", "Magé", "Itaboraí"] * 2 + ["Duque de Caxias"] * 4,
        "Público-Alvo Formatado": ["Agentes Multiplicadores", "Lideranças Comunitárias", "Misto (Agentes e Lideranças)", "Agentes Multiplicadores"] * 3,
        "Subtema": [f"Módulo Prático {i}" for i in range(12)],
        "KM Previsto": [60, 120, 80, 200] * 3,
        "Nº de Participantes": [20, 25, 30, 20] * 3,
        "Carga Horária (h)": [16, 20, 8, 24] * 3,
        "Data": pd.date_range("2026-07-01", periods=12),
        "Status": ["Planejada"] * 8 + ["Concluída"] * 4,
        "Tipo de Atividade": ["Teórica", "Prática", "Visita Técnica"] * 4,
        "Tipo Veículo": ["Van", "Ônibus", "Carro"] * 4,
        "Deslocamento": ["X", "", "X"] * 4,
        "Hospedagem": ["", "Alojamento", ""] * 4,
        "Observações": ["Modo de contingência ativado."] * 12
    })
    
    df_c_fallback = pd.DataFrame({
        0: ["Status", "ID", "Atividades / Produtos (TdR)", "Mês 1", "Mês 2", "Mês 3"],
        1: [np.nan, np.nan, np.nan, "S1", "S2", "S3"],
        2: ["Aprovado", "P1", "Plano de Trabalho", "■", "", ""],
        3: ["Em Andamento", "P2", "Materiais Pedagógicos", "", "■", ""]
    })
    
    return df_fallback, df_c_fallback, "Utilizando base local estável de reserva."

df, df_crono_raw, status_conexao = load_data()

# ==========================================
# 3. SIDEBAR & NAVEGAÇÃO
# ==========================================
with st.sidebar:
    st.markdown('<div class="logo-text">🌿 GUANABARA VERDE</div>', unsafe_allow_html=True)
    menu = st.radio("MENU PRINCIPAL", ["📌 VISÃO GERAL", "📋 GESTÃO OPERACIONAL", "🗺️ TERRITÓRIOS RH-V", "🚐 PAINEL LOGÍSTICO", "📅 CRONOGRAMA", "📦 ENTREGAS"])

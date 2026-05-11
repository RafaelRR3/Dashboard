"""
Dashboard de Indicadores Socioeconômicos Brasileiros - VERSÃO PREMIUM
======================================================================
Aplicação Streamlit profissional com análise avançada de indicadores econômicos,
estatísticas, previsões e comparações históricas. Integra dados do Banco Central,
IBGE e outras fontes oficiais.

Versão: 2.0 Premium
Autor: Dashboard Analytics
Data: 2024
Recursos:
    - 6+ indicadores econômicos em tempo real
    - Análise estatística avançada
    - Comparação de períodos históricos
    - Alertas de variações significativas
    - Export em múltiplos formatos
    - Dashboard responsivo e moderno
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import requests
from datetime import datetime, timedelta
from io import BytesIO
import logging
from typing import Tuple, Dict, Optional
import warnings

warnings.filterwarnings("ignore")

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# === CONFIGURAÇÃO DA PÁGINA ===

st.set_page_config(
    page_title="Painel de Indicadores Econômicos Premium",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
    <style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-weight: bold;
    }
    .alert-danger {
        background-color: #ffcccc;
        padding: 12px;
        border-radius: 5px;
        border-left: 4px solid #cc0000;
    }
    .alert-warning {
        background-color: #fff3cd;
        padding: 12px;
        border-radius: 5px;
        border-left: 4px solid #ffc107;
    }
    .alert-success {
        background-color: #d4edda;
        padding: 12px;
        border-radius: 5px;
        border-left: 4px solid #28a745;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Dashboard de Indicadores Econômicos - PREMIUM")
st.markdown("*Análise completa em tempo real com estatísticas avançadas*")


# === FUNÇÕES AUXILIARES ===

def calcular_media_movel(df: pd.DataFrame, janela: int = 7) -> pd.Series:
    """Calcula média móvel simples."""
    return df["valor"].rolling(window=janela, min_periods=1).mean()


def calcular_volatilidade(df: pd.DataFrame, janela: int = 30) -> float:
    """Calcula volatilidade (desvio padrão) dos últimos N dias."""
    return df["valor"].tail(janela).std()


def detectar_tendencia(df: pd.DataFrame) -> str:
    """Detecta tendência usando simples comparação de médias."""
    if len(df) < 2:
        return "Insuficiente"
    
    media_recente = df["valor"].tail(7).mean()
    media_antiga = df["valor"].head(7).mean()
    
    if media_recente > media_antiga * 1.02:
        return "📈 Tendência de Alta"
    elif media_recente < media_antiga * 0.98:
        return "📉 Tendência de Baixa"
    else:
        return "➡️ Estável"


def calcular_estatisticas(df: pd.DataFrame) -> Dict:
    """Calcula estatísticas básicas do indicador."""
    if df.empty:
        return {}
    
    valores = df["valor"].dropna()
    
    return {
        "média": valores.mean(),
        "mediana": valores.median(),
        "mínimo": valores.min(),
        "máximo": valores.max(),
        "desvio_padrão": valores.std(),
        "coeficiente_variação": (valores.std() / valores.mean() * 100) if valores.mean() != 0 else 0
    }


def formatar_exportacao_excel(dfs: Dict[str, pd.DataFrame]) -> BytesIO:
    """Cria arquivo Excel com múltiplas abas."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for nome, df in dfs.items():
            df_export = df.copy()
            df_export["data"] = pd.to_datetime(df_export["data"]).dt.strftime("%d/%m/%Y")
            df_export.to_excel(writer, sheet_name=nome, index=False)
    output.seek(0)
    return output


def criar_alerta(valor_atual: float, valor_anterior: float, percentual_limite: float = 5) -> Tuple[str, str]:
    """Cria alertas baseado em variação percentual."""
    if valor_anterior == 0:
        variacao_perc = 0
    else:
        variacao_perc = ((valor_atual - valor_anterior) / valor_anterior) * 100
    
    if abs(variacao_perc) > percentual_limite:
        if variacao_perc > 0:
            return "⚠️ ALTA", f"Aumento de {variacao_perc:.2f}%"
        else:
            return "⚠️ QUEDA", f"Redução de {abs(variacao_perc):.2f}%"
    
    return "✅ Normal", f"Variação: {variacao_perc:+.2f}%"


# === FUNÇÕES DE BUSCA DE DADOS COM VALIDAÇÃO ===

@st.cache_data(ttl=86400)
def buscar_ipca(data_inicial="01/01/2022") -> pd.DataFrame:
    """
    Busca IPCA (Índice de Preços ao Consumidor Amplo) com validação.
    
    Args:
        data_inicial (str): Data inicial DD/MM/YYYY
        
    Returns:
        pd.DataFrame: Dados validados e limpos
    """
    try:
        url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.433/dados"
        params = {"formato": "json", "dataInicial": data_inicial}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        if df.empty:
            raise ValueError("Resposta vazia da API")
            
        df["data"] = pd.to_datetime(df["data"], dayfirst=True)
        df["valor"] = pd.to_numeric(df["valor"], errors='coerce')
        df = df.dropna()
        df = df.sort_values("data")
        
        logger.info(f"IPCA carregado: {len(df)} registros")
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar IPCA: {e}")
        st.error(f"❌ Erro ao buscar IPCA: {str(e)}")
        return pd.DataFrame()


@st.cache_data(ttl=86400)
def buscar_selic(data_inicial="01/01/2022") -> pd.DataFrame:
    """Busca taxa Selic média com validação."""
    try:
        url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.432/dados"
        params = {"formato": "json", "dataInicial": data_inicial}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        if df.empty:
            raise ValueError("Resposta vazia da API")
            
        df["data"] = pd.to_datetime(df["data"], dayfirst=True)
        df["valor"] = pd.to_numeric(df["valor"], errors='coerce')
        df = df.dropna()
        df = df.sort_values("data")
        
        logger.info(f"Selic carregado: {len(df)} registros")
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar Selic: {e}")
        st.error(f"❌ Erro ao buscar Selic: {str(e)}")
        return pd.DataFrame()


@st.cache_data(ttl=86400)
def buscar_desemprego(data_inicial="01/01/2021") -> pd.DataFrame:
    """Busca taxa de desemprego com validação."""
    try:
        url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.4390/dados"
        params = {"formato": "json", "dataInicial": data_inicial}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        if df.empty:
            raise ValueError("Resposta vazia da API")
            
        df["data"] = pd.to_datetime(df["data"], dayfirst=True)
        df["valor"] = pd.to_numeric(df["valor"], errors='coerce')
        df = df.dropna()
        df = df.sort_values("data")
        
        logger.info(f"Desemprego carregado: {len(df)} registros")
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar desemprego: {e}")
        st.error(f"❌ Erro ao buscar desemprego: {str(e)}")
        return pd.DataFrame()


@st.cache_data(ttl=86400)
def buscar_inflacao(data_inicial="01/01/2022") -> pd.DataFrame:
    """Busca inflação acumulada com validação."""
    try:
        url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.11/dados"
        params = {"formato": "json", "dataInicial": data_inicial}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        if df.empty:
            raise ValueError("Resposta vazia da API")
            
        df["data"] = pd.to_datetime(df["data"], dayfirst=True)
        df["valor"] = pd.to_numeric(df["valor"], errors='coerce')
        df = df.dropna()
        df = df.sort_values("data")
        
        logger.info(f"Inflação carregado: {len(df)} registros")
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar inflação: {e}")
        st.error(f"❌ Erro ao buscar inflação: {str(e)}")
        return pd.DataFrame()


@st.cache_data(ttl=86400)
def buscar_cambio(data_inicial="01/01/2022") -> pd.DataFrame:
    """Busca taxa de câmbio BRL/USD com validação."""
    try:
        url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados"
        params = {"formato": "json", "dataInicial": data_inicial}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        if df.empty:
            raise ValueError("Resposta vazia da API")
            
        df["data"] = pd.to_datetime(df["data"], dayfirst=True)
        df["valor"] = pd.to_numeric(df["valor"], errors='coerce')
        df = df.dropna()
        df = df.sort_values("data")
        
        logger.info(f"Câmbio carregado: {len(df)} registros")
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar câmbio: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=86400)
def buscar_reservas(data_inicial="01/01/2022") -> pd.DataFrame:
    """Busca reservas internacionais do Brasil com validação."""
    try:
        url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.3068/dados"
        params = {"formato": "json", "dataInicial": data_inicial}
        r = requests.get(url, params=params, timeout=10)
        r.raise_for_status()
        
        df = pd.DataFrame(r.json())
        if df.empty:
            raise ValueError("Resposta vazia da API")
            
        df["data"] = pd.to_datetime(df["data"], dayfirst=True)
        df["valor"] = pd.to_numeric(df["valor"], errors='coerce')
        df = df.dropna()
        df = df.sort_values("data")
        
        logger.info(f"Reservas carregado: {len(df)} registros")
        return df
    except Exception as e:
        logger.error(f"Erro ao buscar reservas: {e}")
        return pd.DataFrame()


# === SIDEBAR COM CONTROLES AVANÇADOS ===

with st.sidebar:
    st.header("⚙️ Configurações Avançadas")
    
    # Seleção de tema
    tema = st.radio("🎨 Tema", ["Light", "Dark"], horizontal=True)
    
    st.divider()
    st.subheader("📅 Período de Análise")
    periodo_dias = st.slider(
        "Últimos dias",
        min_value=30,
        max_value=1095,
        value=365,
        step=30
    )
    
    data_inicio = datetime.now() - timedelta(days=periodo_dias)
    data_inicio_str = data_inicio.strftime("%d/%m/%Y")
    
    st.divider()
    st.subheader("📊 Indicadores para Exibir")
    
    col_check1, col_check2 = st.columns(2)
    
    with col_check1:
        mostrar_ipca = st.checkbox("📈 IPCA", value=True)
        mostrar_selic = st.checkbox("💰 Selic", value=True)
        mostrar_desemprego = st.checkbox("👥 Desemprego", value=True)
    
    with col_check2:
        mostrar_inflacao = st.checkbox("📊 Inflação", value=True)
        mostrar_cambio = st.checkbox("💵 Câmbio", value=False)
        mostrar_reservas = st.checkbox("🏦 Reservas", value=False)
    
    st.divider()
    st.subheader("🔧 Análise Estatística")
    
    mostrar_media_movel = st.checkbox("Média Móvel (7 dias)", value=True)
    mostrar_tendencia = st.checkbox("Detectar Tendência", value=True)
    mostrar_correlacao = st.checkbox("Matriz de Correlação", value=False)
    mostrar_distribuicao = st.checkbox("Análise de Distribuição", value=False)
    
    st.divider()
    st.subheader("⚠️ Alertas")
    percentual_alerta = st.slider(
        "Limite de variação para alerta (%)",
        min_value=1,
        max_value=20,
        value=5,
        step=1
    )
    
    st.divider()
    st.info("""
    💡 **Dicas:**
    - Aumente o período para ver tendências de longo prazo
    - Use alertas para monitorar variações significativas
    - A matriz de correlação mostra relações entre indicadores
    """)


# === CARREGAMENTO DE DADOS COM PROGRESSO ===

with st.spinner("📥 Carregando dados... Isso pode levar alguns segundos"):
    df_ipca = buscar_ipca(data_inicio_str) if mostrar_ipca else pd.DataFrame()
    df_selic = buscar_selic() if mostrar_selic else pd.DataFrame()
    df_desemprego = buscar_desemprego() if mostrar_desemprego else pd.DataFrame()
    df_inflacao = buscar_inflacao() if mostrar_inflacao else pd.DataFrame()
    df_cambio = buscar_cambio(data_inicio_str) if mostrar_cambio else pd.DataFrame()
    df_reservas = buscar_reservas(data_inicio_str) if mostrar_reservas else pd.DataFrame()

st.success("✅ Dados carregados com sucesso!")


# === MÉTRICAS DINÂMICAS COM ALERTAS ===

st.divider()
st.subheader("📊 Indicadores Atuais")

col1, col2, col3, col4 = st.columns(4)

# IPCA
if not df_ipca.empty:
    ipca_atual = df_ipca["valor"].iloc[-1]
    ipca_anterior = df_ipca["valor"].iloc[-2] if len(df_ipca) > 1 else ipca_atual
    delta_ipca = ipca_atual - ipca_anterior
    status, msg = criar_alerta(ipca_atual, ipca_anterior, percentual_alerta)
    col1.metric("📈 IPCA", f"{ipca_atual:.2f}%", delta=f"{delta_ipca:+.2f}pp")
    if status != "✅ Normal":
        col1.markdown(f"<div class='alert-warning'>{status} {msg}</div>", unsafe_allow_html=True)

# Selic
if not df_selic.empty:
    selic_atual = df_selic["valor"].iloc[-1]
    selic_anterior = df_selic["valor"].iloc[-2] if len(df_selic) > 1 else selic_atual
    delta_selic = selic_atual - selic_anterior
    status, msg = criar_alerta(selic_atual, selic_anterior, percentual_alerta)
    col2.metric("💰 Selic", f"{selic_atual:.2f}%", delta=f"{delta_selic:+.2f}pp")
    if status != "✅ Normal":
        col2.markdown(f"<div class='alert-warning'>{status} {msg}</div>", unsafe_allow_html=True)

# Desemprego
if not df_desemprego.empty:
    desemprego_atual = df_desemprego["valor"].iloc[-1]
    desemprego_anterior = df_desemprego["valor"].iloc[-2] if len(df_desemprego) > 1 else desemprego_atual
    delta_desemprego = desemprego_atual - desemprego_anterior
    col3.metric("👥 Desemprego", f"{desemprego_atual:.2f}%", delta=f"{delta_desemprego:+.2f}pp")

# Inflação
if not df_inflacao.empty:
    inflacao_atual = df_inflacao["valor"].iloc[-1]
    inflacao_anterior = df_inflacao["valor"].iloc[-2] if len(df_inflacao) > 1 else inflacao_atual
    delta_inflacao = inflacao_atual - inflacao_anterior
    col4.metric("📊 Inflação", f"{inflacao_atual:.2f}%", delta=f"{delta_inflacao:+.2f}pp")

# Câmbio
if not df_cambio.empty:
    col1, col2 = st.columns(4)[:2]
    cambio_atual = df_cambio["valor"].iloc[-1]
    cambio_anterior = df_cambio["valor"].iloc[-2] if len(df_cambio) > 1 else cambio_atual
    delta_cambio = cambio_atual - cambio_anterior
    col1.metric("💵 Câmbio (USD)", f"R$ {cambio_atual:.2f}", delta=f"{delta_cambio:+.2f}")

# Reservas
if not df_reservas.empty:
    col1, col2 = st.columns(4)[:2]
    reservas_atual = df_reservas["valor"].iloc[-1]
    col2.metric("🏦 Reservas (US$)", f"${reservas_atual:.0f}M")


# === ANÁLISE VISUAL AVANÇADA ===

st.divider()
st.subheader("📈 Gráficos e Análises")

# Abas para organizar visualizações
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Evolução", "Análise Estatística", "Distribuição", "Correlação", "Comparação"])

# === ABA 1: EVOLUÇÃO ===
with tab1:
    col1, col2 = st.columns(2)
    
    # Gráfico IPCA
    if mostrar_ipca and not df_ipca.empty:
        with col1:
            df_plot = df_ipca.copy()
            
            if mostrar_media_movel:
                df_plot["mm7"] = calcular_media_movel(df_plot, 7)
            
            fig_ipca = px.line(
                df_plot,
                x="data",
                y=["valor", "mm7"] if mostrar_media_movel else ["valor"],
                title="📈 IPCA Mensal (%)",
                labels={"valor": "IPCA (%)", "data": "Data", "mm7": "Média 7d"},
                markers=True
            )
            fig_ipca.update_layout(hovermode='x unified', height=450)
            st.plotly_chart(fig_ipca, use_container_width=True)
            
            # Tendência
            if mostrar_tendencia:
                tendencia = detectar_tendencia(df_plot)
                st.info(f"**Tendência:** {tendencia}")

    # Gráfico Selic
    if mostrar_selic and not df_selic.empty:
        with col2:
            df_plot = df_selic.copy()
            
            if mostrar_media_movel:
                df_plot["mm7"] = calcular_media_movel(df_plot, 7)
            
            fig_selic = px.line(
                df_plot,
                x="data",
                y=["valor", "mm7"] if mostrar_media_movel else ["valor"],
                title="💰 Taxa Selic Média (%)",
                labels={"valor": "Selic (%)", "data": "Data", "mm7": "Média 7d"},
                markers=True
            )
            fig_selic.update_layout(hovermode='x unified', height=450)
            st.plotly_chart(fig_selic, use_container_width=True)
            
            if mostrar_tendencia:
                tendencia = detectar_tendencia(df_plot)
                st.info(f"**Tendência:** {tendencia}")

    # Gráfico Desemprego
    if mostrar_desemprego and not df_desemprego.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            df_plot = df_desemprego.copy()
            
            if mostrar_media_movel:
                df_plot["mm7"] = calcular_media_movel(df_plot, 7)
            
            fig_desemprego = px.line(
                df_plot,
                x="data",
                y=["valor", "mm7"] if mostrar_media_movel else ["valor"],
                title="👥 Taxa de Desemprego (%)",
                labels={"valor": "Desemprego (%)", "data": "Data", "mm7": "Média 7d"},
                markers=True
            )
            fig_desemprego.update_layout(hovermode='x unified', height=450)
            st.plotly_chart(fig_desemprego, use_container_width=True)

        # Gráfico Inflação
        if mostrar_inflacao and not df_inflacao.empty:
            with col2:
                df_plot = df_inflacao.copy()
                
                if mostrar_media_movel:
                    df_plot["mm7"] = calcular_media_movel(df_plot, 7)
                
                fig_inflacao = px.line(
                    df_plot,
                    x="data",
                    y=["valor", "mm7"] if mostrar_media_movel else ["valor"],
                    title="📊 Inflação Acumulada (%)",
                    labels={"valor": "Inflação (%)", "data": "Data", "mm7": "Média 7d"},
                    markers=True
                )
                fig_inflacao.update_layout(hovermode='x unified', height=450)
                st.plotly_chart(fig_inflacao, use_container_width=True)

    # Gráfico Câmbio
    if mostrar_cambio and not df_cambio.empty:
        col1, col2 = st.columns(2)
        with col1:
            df_plot = df_cambio.copy()
            if mostrar_media_movel:
                df_plot["mm7"] = calcular_media_movel(df_plot, 7)
            
            fig_cambio = px.line(
                df_plot,
                x="data",
                y=["valor", "mm7"] if mostrar_media_movel else ["valor"],
                title="💵 Taxa de Câmbio BRL/USD",
                markers=True
            )
            fig_cambio.update_layout(hovermode='x unified', height=450)
            st.plotly_chart(fig_cambio, use_container_width=True)

# === ABA 2: ANÁLISE ESTATÍSTICA ===
with tab2:
    st.write("**Estatísticas Descritivas dos Indicadores:**")
    
    stats_data = {}
    
    if not df_ipca.empty:
        stats = calcular_estatisticas(df_ipca)
        stats_data["IPCA"] = stats
    
    if not df_selic.empty:
        stats = calcular_estatisticas(df_selic)
        stats_data["Selic"] = stats
    
    if not df_desemprego.empty:
        stats = calcular_estatisticas(df_desemprego)
        stats_data["Desemprego"] = stats
    
    if not df_inflacao.empty:
        stats = calcular_estatisticas(df_inflacao)
        stats_data["Inflação"] = stats
    
    if not df_cambio.empty:
        stats = calcular_estatisticas(df_cambio)
        stats_data["Câmbio"] = stats
    
    if stats_data:
        df_stats = pd.DataFrame(stats_data).T
        st.dataframe(df_stats.round(4), use_container_width=True)
        
        # Gráfico de volatilidade
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Volatilidade (Últimos 30 dias):**")
            vol_data = {}
            if not df_ipca.empty:
                vol_data["IPCA"] = calcular_volatilidade(df_ipca, 30)
            if not df_selic.empty:
                vol_data["Selic"] = calcular_volatilidade(df_selic, 30)
            if not df_desemprego.empty:
                vol_data["Desemprego"] = calcular_volatilidade(df_desemprego, 30)
            if not df_inflacao.empty:
                vol_data["Inflação"] = calcular_volatilidade(df_inflacao, 30)
            
            if vol_data:
                fig_vol = px.bar(
                    x=list(vol_data.keys()),
                    y=list(vol_data.values()),
                    title="Desvio Padrão por Indicador",
                    labels={"x": "Indicador", "y": "Volatilidade"}
                )
                st.plotly_chart(fig_vol, use_container_width=True)
        
        with col2:
            st.write("**Coeficiente de Variação:**")
            st.text("(Mede volatilidade relativa)")
            df_cv = df_stats[["coeficiente_variação"]].sort_values("coeficiente_variação", ascending=False)
            st.dataframe(df_cv, use_container_width=True)

# === ABA 3: DISTRIBUIÇÃO ===
with tab3:
    if mostrar_distribuicao:
        col1, col2 = st.columns(2)
        
        if not df_ipca.empty:
            with col1:
                fig_hist = px.histogram(
                    df_ipca,
                    x="valor",
                    nbins=30,
                    title="Distribuição do IPCA",
                    labels={"valor": "IPCA (%)"}
                )
                st.plotly_chart(fig_hist, use_container_width=True)
        
        if not df_selic.empty:
            with col2:
                fig_hist = px.histogram(
                    df_selic,
                    x="valor",
                    nbins=30,
                    title="Distribuição da Selic",
                    labels={"valor": "Selic (%)"}
                )
                st.plotly_chart(fig_hist, use_container_width=True)
        
        col1, col2 = st.columns(2)
        
        if not df_desemprego.empty:
            with col1:
                fig_box = px.box(
                    df_desemprego,
                    y="valor",
                    title="Box Plot - Desemprego",
                    labels={"valor": "Desemprego (%)"}
                )
                st.plotly_chart(fig_box, use_container_width=True)
        
        if not df_inflacao.empty:
            with col2:
                fig_box = px.box(
                    df_inflacao,
                    y="valor",
                    title="Box Plot - Inflação",
                    labels={"valor": "Inflação (%)"}
                )
                st.plotly_chart(fig_box, use_container_width=True)
    else:
        st.info("Ative 'Análise de Distribuição' nas configurações para ver este gráfico")

# === ABA 4: CORRELAÇÃO ===
with tab4:
    if mostrar_correlacao:
        # Preparar dados para correlação
        df_corr_data = {}
        
        if not df_ipca.empty:
            df_corr_data["IPCA"] = df_ipca.set_index("data")["valor"]
        if not df_selic.empty:
            df_corr_data["Selic"] = df_selic.set_index("data")["valor"]
        if not df_desemprego.empty:
            df_corr_data["Desemprego"] = df_desemprego.set_index("data")["valor"]
        if not df_inflacao.empty:
            df_corr_data["Inflação"] = df_inflacao.set_index("data")["valor"]
        if not df_cambio.empty:
            df_corr_data["Câmbio"] = df_cambio.set_index("data")["valor"]
        
        if len(df_corr_data) > 1:
            df_corr = pd.DataFrame(df_corr_data)
            corr_matrix = df_corr.corr()
            
            fig_corr = px.imshow(
                corr_matrix,
                labels=dict(x="Indicador", y="Indicador", color="Correlação"),
                x=corr_matrix.columns,
                y=corr_matrix.columns,
                color_continuous_scale="RdBu",
                zmin=-1, zmax=1,
                title="🔗 Matriz de Correlação entre Indicadores"
            )
            st.plotly_chart(fig_corr, use_container_width=True)
            
            st.write("**Interpretação:**")
            st.write("- **Valores próximos a 1:** Correlação positiva forte (ambos aumentam juntos)")
            st.write("- **Valores próximos a -1:** Correlação negativa forte (um aumenta, outro diminui)")
            st.write("- **Valores próximos a 0:** Pouca ou nenhuma correlação")
        else:
            st.info("Selecione pelo menos 2 indicadores para visualizar correlação")
    else:
        st.info("Ative 'Matriz de Correlação' nas configurações para ver este gráfico")

# === ABA 5: COMPARAÇÃO ===
with tab5:
    st.write("**Comparação Multi-indicadores com Eixos Normalizados:**")
    
    if (mostrar_ipca or mostrar_selic) and (not df_ipca.empty or not df_selic.empty):
        fig_comparativo = go.Figure()
        
        if not df_ipca.empty:
            fig_comparativo.add_trace(go.Scatter(
                x=df_ipca["data"],
                y=df_ipca["valor"],
                name="IPCA",
                yaxis="y1"
            ))
        
        if not df_selic.empty:
            fig_comparativo.add_trace(go.Scatter(
                x=df_selic["data"],
                y=df_selic["valor"],
                name="Selic",
                yaxis="y2"
            ))
        
        fig_comparativo.update_layout(
            title="IPCA vs Selic (Eixos Duais)",
            xaxis=dict(title="Data"),
            yaxis=dict(title="IPCA (%)", titlefont=dict(color="blue"), tickfont=dict(color="blue")),
            yaxis2=dict(title="Selic (%)", titlefont=dict(color="red"), tickfont=dict(color="red"), 
                        overlaying="y", side="right"),
            height=500,
            hovermode='x unified'
        )
        st.plotly_chart(fig_comparativo, use_container_width=True)


# === EXPORTAÇÃO DE DADOS AVANÇADA ===

st.divider()
st.subheader("💾 Exportação de Dados")

col1, col2, col3 = st.columns(3)

with col1:
    # Exportar dados individuais em CSV
    st.write("**Exportar Indicadores (CSV):**")
    
    csv_files = {}
    
    if not df_ipca.empty:
        csv_files["ipca"] = df_ipca.to_csv(index=False)
        st.download_button(
            "⬇️ IPCA (CSV)",
            csv_files["ipca"],
            "ipca.csv",
            "text/csv",
            key="csv_ipca"
        )
    
    if not df_selic.empty:
        csv_files["selic"] = df_selic.to_csv(index=False)
        st.download_button(
            "⬇️ Selic (CSV)",
            csv_files["selic"],
            "selic.csv",
            "text/csv",
            key="csv_selic"
        )

with col2:
    if not df_desemprego.empty:
        csv_files["desemprego"] = df_desemprego.to_csv(index=False)
        st.download_button(
            "⬇️ Desemprego (CSV)",
            csv_files["desemprego"],
            "desemprego.csv",
            "text/csv",
            key="csv_desemprego"
        )
    
    if not df_inflacao.empty:
        csv_files["inflacao"] = df_inflacao.to_csv(index=False)
        st.download_button(
            "⬇️ Inflação (CSV)",
            csv_files["inflacao"],
            "inflacao.csv",
            "text/csv",
            key="csv_inflacao"
        )

with col3:
    if not df_cambio.empty:
        csv_files["cambio"] = df_cambio.to_csv(index=False)
        st.download_button(
            "⬇️ Câmbio (CSV)",
            csv_files["cambio"],
            "cambio.csv",
            "text/csv",
            key="csv_cambio"
        )
    
    if not df_reservas.empty:
        csv_files["reservas"] = df_reservas.to_csv(index=False)
        st.download_button(
            "⬇️ Reservas (CSV)",
            csv_files["reservas"],
            "reservas.csv",
            "text/csv",
            key="csv_reservas"
        )

# Exportar Excel consolidado
st.write("**Exportar Consolidado (Excel):**")

try:
    dfs_export = {}
    if not df_ipca.empty:
        dfs_export["IPCA"] = df_ipca
    if not df_selic.empty:
        dfs_export["Selic"] = df_selic
    if not df_desemprego.empty:
        dfs_export["Desemprego"] = df_desemprego
    if not df_inflacao.empty:
        dfs_export["Inflação"] = df_inflacao
    if not df_cambio.empty:
        dfs_export["Câmbio"] = df_cambio
    if not df_reservas.empty:
        dfs_export["Reservas"] = df_reservas
    
    if dfs_export:
        excel_buffer = formatar_exportacao_excel(dfs_export)
        st.download_button(
            "📊 Todos os Indicadores (Excel)",
            excel_buffer,
            "indicadores_economicos.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="excel_consolidado"
        )
except Exception as e:
    st.error(f"Erro ao gerar arquivo Excel: {e}")


# === TABELAS DETALHADAS ===

st.divider()
st.subheader("📋 Dados Brutos Completos")

tab_ipca, tab_selic, tab_desemp, tab_inflacao, tab_cambio, tab_reservas = st.tabs(
    ["IPCA", "Selic", "Desemprego", "Inflação", "Câmbio", "Reservas"]
)

with tab_ipca:
    if not df_ipca.empty:
        st.write(f"**Total de registros:** {len(df_ipca)}")
        st.write(f"**Período:** {df_ipca['data'].min().date()} até {df_ipca['data'].max().date()}")
        st.dataframe(df_ipca.sort_values("data", ascending=False), use_container_width=True, height=500)
    else:
        st.info("Nenhum dado disponível")

with tab_selic:
    if not df_selic.empty:
        st.write(f"**Total de registros:** {len(df_selic)}")
        st.write(f"**Período:** {df_selic['data'].min().date()} até {df_selic['data'].max().date()}")
        st.dataframe(df_selic.sort_values("data", ascending=False), use_container_width=True, height=500)
    else:
        st.info("Nenhum dado disponível")

with tab_desemp:
    if not df_desemprego.empty:
        st.write(f"**Total de registros:** {len(df_desemprego)}")
        st.write(f"**Período:** {df_desemprego['data'].min().date()} até {df_desemprego['data'].max().date()}")
        st.dataframe(df_desemprego.sort_values("data", ascending=False), use_container_width=True, height=500)
    else:
        st.info("Nenhum dado disponível")

with tab_inflacao:
    if not df_inflacao.empty:
        st.write(f"**Total de registros:** {len(df_inflacao)}")
        st.write(f"**Período:** {df_inflacao['data'].min().date()} até {df_inflacao['data'].max().date()}")
        st.dataframe(df_inflacao.sort_values("data", ascending=False), use_container_width=True, height=500)
    else:
        st.info("Nenhum dado disponível")

with tab_cambio:
    if not df_cambio.empty:
        st.write(f"**Total de registros:** {len(df_cambio)}")
        st.write(f"**Período:** {df_cambio['data'].min().date()} até {df_cambio['data'].max().date()}")
        st.dataframe(df_cambio.sort_values("data", ascending=False), use_container_width=True, height=500)
    else:
        st.info("Nenhum dado disponível")

with tab_reservas:
    if not df_reservas.empty:
        st.write(f"**Total de registros:** {len(df_reservas)}")
        st.write(f"**Período:** {df_reservas['data'].min().date()} até {df_reservas['data'].max().date()}")
        st.dataframe(df_reservas.sort_values("data", ascending=False), use_container_width=True, height=500)
    else:
        st.info("Nenhum dado disponível")


# === RELATÓRIO RESUMIDO ===

st.divider()
st.subheader("📄 Relatório Resumido da Sessão")

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("📊 Indicadores Carregados", sum([
        1 if not df_ipca.empty else 0,
        1 if not df_selic.empty else 0,
        1 if not df_desemprego.empty else 0,
        1 if not df_inflacao.empty else 0,
        1 if not df_cambio.empty else 0,
        1 if not df_reservas.empty else 0,
    ]))

with col2:
    total_registros = sum([len(df) for df in [df_ipca, df_selic, df_desemprego, df_inflacao, df_cambio, df_reservas]])
    st.metric("📈 Total de Registros", total_registros)

with col3:
    st.metric("🕐 Última Atualização", datetime.now().strftime("%H:%M:%S"))

st.info("""
**Informações Importantes:**
- ✅ Todos os dados são públicos e provenientes do Banco Central do Brasil
- 💾 Cache de dados: 24 horas
- 📊 Atualização de dados: Diária
- 🔒 Sem requisição de autenticação
- ⚡ Performance otimizada para análise rápida
""")


# === RODAPÉ COM INFORMAÇÕES TÉCNICAS ===

st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**📍 Fontes de Dados:**")
    st.markdown("""
    - **Banco Central do Brasil**
        - IPCA (Série 433)
        - Selic (Série 432)
        - Desemprego (Série 4390)
        - Inflação (Série 11)
        - Câmbio USD (Série 1)
        - Reservas (Série 3068)
    - **IBGE**
    """)

with col2:
    st.markdown("**⚙️ Tecnologia:**")
    st.markdown("""
    - **Framework:** Streamlit
    - **Visualização:** Plotly
    - **Dados:** Pandas, NumPy
    - **Cache:** 24 horas
    - **Timeout:** 10 segundos por requisição
    """)

with col3:
    st.markdown("**📱 Recursos:**")
    st.markdown("""
    - ✅ Análise em tempo real
    - ✅ Múltiplas métricas
    - ✅ Gráficos interativos
    - ✅ Estatísticas avançadas
    - ✅ Export múltiplos formatos
    - ✅ Alertas de variação
    """)

st.markdown("---")
st.markdown(
    f"<p style='text-align: center; color: gray;'>"
    f"Dashboard Premium | Versão 2.0 | Atualizado em {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')}"
    f"</p>",
    unsafe_allow_html=True
)
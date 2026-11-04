import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard de Contratos e Posições", layout="wide")

st.title("📊 Dashboard Consolidado - BTG & XP")
st.markdown("Visualização integrada de Contratos, Posições e Renda Fixa.")

# Carregar base consolidada ou ler diretamente os excels
@st.cache_data
def carregar_dados():
    # Leitura simplificada para demonstração
    df_pos_btg = pd.read_excel("data/Posição.xlsx", sheet_name="Export")
    xls_xp = pd.ExcelFile("data/relatorio_posicoes_consolidada_04112026141010.xlsx")
    df_xp_fin = pd.read_excel(xls_xp, sheet_name="Financeiro")
    return df_pos_btg, df_xp_fin

df_btg, df_xp = carregar_dados()

# Métricas principais
col1, col2 = st.columns(2)
with col1:
    st.metric("Total Contas BTG", df_btg['Conta'].nunique() if 'Conta' in df_btg.columns else 0)
with col2:
    patrimonio_xp = df_xp['PatrimonioTotal'].sum() if 'PatrimonioTotal' in df_xp.columns else 0
    st.metric("Patrimônio Total XP (Soma)", f"R$ {patrimonio_xp:,.2f}")

st.subheader("Visão Geral das Posições - BTG (Mercados)")
if 'Mercado' in df_btg.columns:
    fig = px.pie(df_btg, names='Mercado', title="Distribuição por Mercado - BTG")
    st.plotly_chart(fig, use_container_width=True)

st.subheader("Dados Detalhados - XP (Financeiro)")
st.dataframe(df_xp.head(20))

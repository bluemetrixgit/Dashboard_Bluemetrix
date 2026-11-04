import glob
import os
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard Bluemetrix - Controle de Contratos",
    page_icon="📈",
    layout="wide",
)

st.title("📊 Dashboard Bluemetrix - Controle de Contratos")


# Função para carregar e processar os dados com segurança contra colunas duplicadas
@st.cache_data
def carregar_dados():
  colunas_desejadas = [
      "Corretora",
      "Cliente",
      "Conta",
      "Status",
      "Situação",
      "Carteira",
      "Observações",
  ]

  path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
  if not os.path.exists(path_contratos):
    path_contratos = "Controle de Contratos - Atualizado 2026.xlsx"

  if not os.path.exists(path_contratos):
    st.error(
        f"Arquivo de contratos não encontrado no caminho: {path_contratos}"
    )
    return pd.DataFrame()

  try:
    df_btg = pd.read_excel(path_contratos, sheet_name="BTG", header=1)
    df_xp = pd.read_excel(path_contratos, sheet_name="XP", header=1)

    df_btg["Corretora"] = "BTG"
    df_xp["Corretora"] = "XP"

    # Remove colunas duplicadas do DataFrame para evitar erro de reindexação
    df_btg = df_btg.loc[:, ~df_btg.columns.duplicated()]
    df_xp = df_xp.loc[:, ~df_xp.columns.duplicated()]

    # Seleciona apenas as colunas desejadas que realmente existem na planilha
    cols_btg = [c for c in colunas_desejadas if c in df_btg.columns]
    cols_xp = [c for c in colunas_desejadas if c in df_xp.columns]

    df_btg_filtrado = df_btg[cols_btg].copy()
    df_xp_filtrado = df_xp[cols_xp].copy()

    contratos_geral = pd.concat(
        [df_btg_filtrado, df_xp_filtrado], ignore_index=True
    )
    return contratos_geral

  except Exception as e:
    st.error(f"Erro ao processar o arquivo Excel de Contratos: {e}")
    return pd.DataFrame()


# Carregando os dados
contratos_df = carregar_dados()

if not contratos_df.empty:
  # ==========================================
  # CRIAÇÃO DAS ABAS: BTG e XP
  # ==========================================
  aba_btg, aba_xp = st.tabs(["BTG", "XP"])

  # --- ABA BTG ---
  with aba_btg:
    st.subheader("🔵 Dashboard - Corretora BTG")

    df_btg_view = contratos_df[contratos_df["Corretora"] == "BTG"].copy()

    if "Carteira" in df_btg_view.columns:
      carteiras_btg = ["Todas"] + sorted(
          [str(c) for c in df_btg_view["Carteira"].dropna().unique()]
      )
      cart_btg_sel = st.selectbox(
          "Filtrar Carteira (BTG)", carteiras_btg, key="sel_cart_btg"
      )
      if cart_btg_sel != "Todas":
        df_btg_view = df_btg_view[
            df_btg_view["Carteira"].astype(str) == cart_btg_sel
        ]

    st.metric("Total de Contratos (BTG)", len(df_btg_view))
    st.markdown("### Contratos Cadastrados")
    st.dataframe(df_btg_view, use_container_width=True)

  # --- ABA XP ---
  with aba_xp:
    st.subheader("🟡 Dashboard - Corretora XP")

    df_xp_view = contratos_df[contratos_df["Corretora"] == "XP"].copy()

    if "Carteira" in df_xp_view.columns:
      carteiras_xp = ["Todas"] + sorted(
          [str(c) for c in df_xp_view["Carteira"].dropna().unique()]
      )
      cart_xp_sel = st.selectbox(
          "Filtrar Carteira (XP)", carteiras_xp, key="sel_cart_xp"
      )
      if cart_xp_sel != "Todas":
        df_xp_view = df_xp_view[
            df_xp_view["Carteira"].astype(str) == cart_xp_sel
        ]

    st.metric("Total de Contratos (XP)", len(df_xp_view))
    st.markdown("### Contratos Cadastrados")
    st.dataframe(df_xp_view, use_container_width=True)

else:
  st.warning(
      "Nenhum dado foi carregado. Verifique se o arquivo 'Controle de"
      " Contratos - Atualizado 2026.xlsx' está dentro da pasta 'data/' no"
      " GitHub."
  )






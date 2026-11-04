import glob
import os
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard Bluemetrix - Consolidado",
    page_icon="📈",
    layout="wide",
)

st.title("📊 Dashboard Bluemetrix - Controle de Contratos")


# Função para carregar e processar os dados com segurança
@st.cache_data
def carregar_dados():
  # Colunas exatas desejadas do Controle de Contratos
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

  try:
    # Lendo ignorando a primeira linha (header=1)
    df_btg = pd.read_excel(path_contratos, sheet_name="BTG", header=1)
    df_xp = pd.read_excel(path_contratos, sheet_name="XP", header=1)

    df_btg["Corretora"] = "BTG"
    df_xp["Corretora"] = "XP"

    # Filtrando apenas as colunas que existem na planilha para evitar erros
    cols_btg_existentes = [
        c for c in colunas_desejadas if c in df_btg.columns
    ] + [
        c
        for c in df_btg.columns
        if "/" in str(c) or pd.api.types.is_numeric_dtype(df_btg[c])
    ]
    cols_xp_existentes = [c for c in colunas_desejadas if c in df_xp.columns] + [
        c
        for c in df_xp.columns
        if "/" in str(c) or pd.api.types.is_numeric_dtype(df_xp[c])
    ]

    df_btg = df_btg[cols_btg_existentes]
    df_xp = df_xp[cols_xp_existentes]

    contratos_geral = pd.concat([df_btg, df_xp], ignore_index=True)
  except Exception as e:
    st.error(f"Erro ao ler o Controle de Contratos: {e}")
    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

  # Normalizando as contas para string (removendo .0)
  if "Conta" in contratos_geral.columns:
    contratos_geral["Conta_Str"] = (
        contratos_geral["Conta"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )

  # Leitura e Filtro de Match da Posição BTG
  path_pos_btg = (
      "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
  )
  btg_pos = pd.DataFrame()
  try:
    if os.path.exists(path_pos_btg):
      df_pos_btg_raw = pd.read_excel(path_pos_btg, sheet_name="Export")
      if "Conta" in df_pos_btg_raw.columns and "Conta_Str" in contratos_geral.columns:
        df_pos_btg_raw["Conta_Str"] = (
            df_pos_btg_raw["Conta"]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )
        contas_validas_btg = set(
            contratos_geral[contratos_geral["Corretora"] == "BTG"][
                "Conta_Str"
            ].dropna()
        )
        btg_pos = df_pos_btg_raw[
            df_pos_btg_raw["Conta_Str"].isin(contas_validas_btg)
        ].copy()
  except Exception as e:
    st.warning(f"Aviso ao ler Posição BTG: {e}")

  # Leitura e Filtro de Match da XP Financeiro
  arquivos_xp = glob.glob("data/relatorio_posicoes_consolidada*.xlsx") + glob.glob(
      "relatorio_posicoes_consolidada*.xlsx"
  )
  xp_fin = pd.DataFrame()
  if arquivos_xp:
    arquivo_xp_recente = max(arquivos_xp, key=os.path.getmtime)
    try:
      xls_xp = pd.ExcelFile(arquivo_xp_recente)
      if "Financeiro" in xls_xp.sheet_names:
        df_xp_fin_raw = pd.read_excel(xls_xp, sheet_name="Financeiro")
        col_conta_xp = (
            "CodigoCliente"
            if "CodigoCliente" in df_xp_fin_raw.columns
            else "Conta"
        )
        if (
            col_conta_xp in df_xp_fin_raw.columns
            and "Conta_Str" in contratos_geral.columns
        ):
          df_xp_fin_raw["Conta_Str"] = (
              df_xp_fin_raw[col_conta_xp]
              .astype(str)
              .str.replace(r"\.0$", "", regex=True)
              .str.strip()
          )
          contas_validas_xp = set(
              contratos_geral[contratos_geral["Corretora"] == "XP"][
                  "Conta_Str"
              ].dropna()
          )
          xp_fin = df_xp_fin_raw[
              df_xp_fin_raw["Conta_Str"].isin(contas_validas_xp)
          ].copy()
    except Exception as e:
      st.warning(f"Aviso ao ler o relatório da XP: {e}")

  return contratos_geral, xp_fin, btg_pos


# Carregando os dados
contratos_df, xp_fin_df, btg_pos_df = carregar_dados()

if not contratos_df.empty:
  # ==========================================
  # CRIAÇÃO DAS ABAS: BTG e XP
  # ==========================================
  aba_btg, aba_xp = st.tabs(["BTG", "XP"])

  # --- ABA BTG ---
  with aba_btg:
    st.subheader("🔵 Dashboard - Corretora BTG")

    df_btg_view = contratos_df[contratos_df["Corretora"] == "BTG"].copy()

    # Filtro lateral específico ou geral para BTG
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
      "Nenhum dado encontrado. Verifique se os arquivos Excel estão na pasta"
      " correta."
  )






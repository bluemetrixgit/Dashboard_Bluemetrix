import glob
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard Bluemetrix - Posição e Contratos",
    page_icon="📈",
    layout="wide",
)

st.title("📊 Dashboard Bluemetrix - Posição Consolidada de Clientes")


# =========================
# FUNÇÕES AUXILIARES
# =========================


def limpar_conta(coluna):
  """Padroniza a conta removendo decimals .

  0, espaços e zeros à esquerda.
  """
  return (
      coluna.astype(str)
      .str.replace(r"\.0$", "", regex=True)
      .str.strip()
      .str.lstrip("0")
  )


def formatar_real(valor):
  """Formata valores numéricos para o padrão de moeda brasileiro R$ X.XXX,XX."""
  if pd.isna(valor) or valor == 0:
    return "R$ 0,00"
  return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# =========================
# CARREGAR E PROCESSAR DADOS (SEM CACHE)
# =========================


def carregar_dados():
  colunas_contrato = [
      "Corretora",
      "Cliente",
      "Conta",
      "Status",
      "Situação",
      "Carteira",
      "Observações",
  ]

  # 1. Leitura do Controle de Contratos
  path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
  if not os.path.exists(path_contratos):
    path_contratos = "Controle de Contratos - Atualizado 2026.xlsx"

  if not os.path.exists(path_contratos):
    st.error(f"Arquivo não encontrado: {path_contratos}")
    return pd.DataFrame(), pd.DataFrame()

  try:
    df_btg = pd.read_excel(
        path_contratos, sheet_name="BTG", header=1, engine="openpyxl"
    )
    df_xp = pd.read_excel(
        path_contratos, sheet_name="XP", header=1, engine="openpyxl"
    )

    df_btg["Corretora"] = "BTG"
    df_xp["Corretora"] = "XP"

    # Limpar duplicidades de colunas
    df_btg = df_btg.loc[:, ~df_btg.columns.duplicated()]
    df_xp = df_xp.loc[:, ~df_xp.columns.duplicated()]

    cols_btg = [c for c in colunas_contrato if c in df_btg.columns]
    cols_xp = [c for c in colunas_contrato if c in df_xp.columns]

    contratos_geral = pd.concat(
        [df_btg[cols_btg], df_xp[cols_xp]], ignore_index=True
    )

    if "Conta" in contratos_geral.columns:
      contratos_geral["Conta_Limpa"] = limpar_conta(contratos_geral["Conta"])
  except Exception as e:
    st.error(f"Erro ao ler Controle de Contratos: {e}")
    return pd.DataFrame(), pd.DataFrame()

  # 2. Leitura da Posição BTG
  path_pos_btg = (
      "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
  )
  btg_pos_pivot = pd.DataFrame()

  if os.path.exists(path_pos_btg):
    try:
      posicao = pd.read_excel(
          path_pos_btg, sheet_name="Export", engine="openpyxl"
      )
      posicao.columns = posicao.columns.str.strip()

      if "Conta" in posicao.columns:
        posicao["Conta_Limpa"] = limpar_conta(posicao["Conta"])

        # Regra de definição das colunas por tipo de Mercado
        def definir_coluna(row):
          mercado = str(row.get("Mercado", "")).strip().lower()
          if "renda fixa" in mercado:
            return str(row.get("Produto", "Outros RF"))
          elif "renda variável" in mercado or "renda variavel" in mercado:
            return str(row.get("Sub Mercado", "Outros RV"))
          return f"{row.get('Mercado', 'Outros')}"

        posicao["Coluna_Mercado"] = posicao.apply(definir_coluna, axis=1)

        # Tratar o Valor Bruto para garantir conversão numérica limpa
        col_vb = (
            "Valor Bruto"
            if "Valor Bruto" in posicao.columns
            else [c for c in posicao.columns if "valor" in c.lower()][0]
        )
        if posicao[col_vb].dtype == object:
          posicao["Valor_Bruto_Num"] = pd.to_numeric(
              posicao[col_vb]
              .astype(str)
              .str.replace("R$", "", regex=False)
              .str.replace(".", "", regex=False)
              .str.replace(",", ".", regex=False)
              .str.strip(),
              errors="coerce",
          ).fillna(0)
        else:
          posicao["Valor_Bruto_Num"] = pd.to_numeric(
              posicao[col_vb], errors="coerce"
          ).fillna(0)

        # Tabela dinâmica pivoteando as posições por Conta
        btg_pos_pivot = posicao.pivot_table(
            index="Conta_Limpa",
            columns="Coluna_Mercado",
            values="Valor_Bruto_Num",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

    except Exception as e:
      st.warning(f"Aviso ao ler Posição BTG: {e}")

  return contratos_geral, btg_pos_pivot


# =========================
# MONTAGEM DA INTERFACE
# =========================

contratos_df, btg_pos_pivot = carregar_dados()

if not contratos_df.empty:
  aba_btg, aba_xp = st.tabs(["BTG", "XP"])

  # -------------------------
  # ABA BTG
  # -------------------------
  with aba_btg:
    st.subheader("🔵 Corretora BTG - Contratos & Posições")

    df_btg_view = contratos_df[contratos_df["Corretora"] == "BTG"].copy()

    # Cruzamento (Merge) entre Contratos e Posições Pivoted pela Conta Limpa
    if not btg_pos_pivot.empty:
      df_btg_view = df_btg_view.merge(
          btg_pos_pivot, on="Conta_Limpa", how="left"
      )

      # Preencher NaN com 0 nas colunas dinâmicas criadas
      colunas_posicao = [
          c for c in btg_pos_pivot.columns if c != "Conta_Limpa"
      ]
      for col in colunas_posicao:
        df_btg_view[col] = df_btg_view[col].fillna(0)

    # Filtro de Carteira
    if "Carteira" in df_btg_view.columns:
      carteiras = ["Todas"] + sorted(
          [str(c) for c in df_btg_view["Carteira"].dropna().unique()]
      )
      sel_cart = st.selectbox(
          "Filtrar por Carteira", carteiras, key="cart_btg_sel"
      )
      if sel_cart != "Todas":
        df_btg_view = df_btg_view[
            df_btg_view["Carteira"].astype(str) == sel_cart
        ]

    st.metric("Total de Contratos", len(df_btg_view))

    # Formatação Visual da Tabela
    df_exibir_btg = df_btg_view.copy()

    # Formatar valores das colunas dinâmicas em Moeda R$
    if not btg_pos_pivot.empty:
      for col in colunas_posicao:
        if col in df_exibir_btg.columns:
          df_exibir_btg[col] = df_exibir_btg[col].apply(formatar_real)

    # Remover a coluna chave temporária da visualização
    if "Conta_Limpa" in df_exibir_btg.columns:
      df_exibir_btg = df_exibir_btg.drop(columns=["Conta_Limpa"])

    st.dataframe(df_exibir_btg, use_container_width=True)

  # -------------------------
  # ABA XP
  # -------------------------
  with aba_xp:
    st.subheader("🟡 Corretora XP - Contratos")

    df_xp_view = contratos_df[contratos_df["Corretora"] == "XP"].copy()

    if "Carteira" in df_xp_view.columns:
      carteiras_xp = ["Todas"] + sorted(
          [str(c) for c in df_xp_view["Carteira"].dropna().unique()]
      )
      sel_cart_xp = st.selectbox(
          "Filtrar por Carteira", carteiras_xp, key="cart_xp_sel"
      )
      if sel_cart_xp != "Todas":
        df_xp_view = df_xp_view[
            df_xp_view["Carteira"].astype(str) == sel_cart_xp
        ]

    if "Conta_Limpa" in df_xp_view.columns:
      df_xp_view = df_xp_view.drop(columns=["Conta_Limpa"])

    st.metric("Total de Contratos", len(df_xp_view))
    st.dataframe(df_xp_view, use_container_width=True)

else:
  st.warning(
      "Nenhum dado encontrado. Verifique se os arquivos Excel estão na pasta."
  )




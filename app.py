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


# Função com cache otimizado para carregar e cruzar os dados com segurança
@st.cache_data(ttl=60)
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

    df_btg = df_btg.loc[:, ~df_btg.columns.duplicated()]
    df_xp = df_xp.loc[:, ~df_xp.columns.duplicated()]

    cols_btg = [c for c in colunas_desejadas if c in df_btg.columns]
    cols_xp = [c for c in colunas_desejadas if c in df_xp.columns]

    df_btg_filtrado = df_btg[cols_btg].copy()
    df_xp_filtrado = df_xp[cols_xp].copy()

    contratos_geral = pd.concat(
        [df_btg_filtrado, df_xp_filtrado], ignore_index=True
    )

    if "Conta" in contratos_geral.columns:
      contratos_geral["Conta_Str"] = (
          contratos_geral["Conta"]
          .astype(str)
          .str.replace(r"\.0$", "", regex=True)
          .str.strip()
      )
  except Exception as e:
    st.error(f"Erro ao processar o arquivo Excel de Contratos: {e}")
    return pd.DataFrame(), pd.DataFrame()

  # Leitura e processamento da Posição BTG
  path_pos_btg = (
      "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
  )
  btg_pos_pivot = pd.DataFrame()

  if os.path.exists(path_pos_btg):
    try:
      df_pos_btg = pd.read_excel(
          path_pos_btg, sheet_name="Export", engine="openpyxl"
      )
      df_pos_btg.columns = [str(c).strip() for c in df_pos_btg.columns]

      if "Conta" in df_pos_btg.columns:
        df_pos_btg["Conta_Str"] = (
            df_pos_btg["Conta"]
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )

        if not df_pos_btg.empty and "Mercado" in df_pos_btg.columns:

          def define_coluna_posicao(row):
            mercado = str(row.get("Mercado", "")).strip().lower()
            if "renda fixa" in mercado:
              return str(row.get("Produto", "Outros RF"))
            elif "renda variável" in mercado or "renda variavel" in mercado:
              return str(row.get("Sub Mercado", "Outros RV"))
            return f"{row.get('Mercado', 'Outros')} - Outros"

          df_pos_btg["Coluna_Dinamica"] = df_pos_btg.apply(
              define_coluna_posicao, axis=1
          )

          col_vb = (
              "Valor Bruto"
              if "Valor Bruto" in df_pos_btg.columns
              else [
                  c for c in df_pos_btg.columns if "valor" in c.lower()
              ][0]
              if [c for c in df_pos_btg.columns if "valor" in c.lower()]
              else None
          )
          if col_vb:
            df_pos_btg["Valor Bruto"] = pd.to_numeric(
                df_pos_btg[col_vb], errors="coerce"
            ).fillna(0)
          else:
            df_pos_btg["Valor Bruto"] = 0

          # Pivot table agrupando por conta e colunas dinâmicas (soma do Valor Bruto)
          btg_pos_pivot = df_pos_btg.pivot_table(
              index="Conta_Str",
              columns="Coluna_Dinamica",
              values="Valor Bruto",
              aggfunc="sum",
              fill_value=0,
          ).reset_index()
    except Exception as e:
      st.warning(f"Aviso ao processar Posição BTG: {e}")

  return contratos_geral, btg_pos_pivot


contratos_df, btg_pos_pivot = carregar_dados()

if not contratos_df.empty:
  aba_btg, aba_xp = st.tabs(["BTG", "XP"])

  # --- ABA BTG ---
  with aba_btg:
    st.subheader("🔵 Dashboard - Corretora BTG (Com Posições Casadas)")

    df_btg_view = contratos_df[contratos_df["Corretora"] == "BTG"].copy()

    # Fazendo o merge (left join) para trazer as colunas dinâmicas com o Valor Bruto
    if not btg_pos_pivot.empty:
      df_btg_view = pd.merge(
          df_btg_view, btg_pos_pivot, on="Conta_Str", how="left"
      )
      colunas_pos_criadas = [
          c for c in btg_pos_pivot.columns if c != "Conta_Str"
      ]
      for col in colunas_pos_criadas:
        if col in df_btg_view.columns:
          df_btg_view[col] = df_btg_view[col].fillna(0)

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
    st.markdown("### Contratos e Posições Consolidadas")
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
  st.warning("Nenhum dado foi carregado. Verifique os arquivos na pasta data/.")




import os
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard Bluemetrix", page_icon="📈", layout="wide"
)

st.title("📊 Dashboard Bluemetrix - Consolidado")


# Função para limpar e converter os valores financeiros com segurança
def limpar_e_converter(df, colunas_alvo):
  df_copia = df.copy()
  for col in colunas_alvo:
    if col in df_copia.columns:
      # Se for do tipo object (texto), limpa símbolos e ajusta formato brasileiro para padrão float
      if df_copia[col].dtype == object:
        df_copia[col] = (
            df_copia[col]
            .astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
      df_copia[col] = pd.to_numeric(df_copia[col], errors="coerce").fillna(0)
  return df_copia


@st.cache_data
def carregar_dados():
  contratos_geral = pd.DataFrame()
  xp_fin = pd.DataFrame()
  btg_pos = pd.DataFrame()

  path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
  if not os.path.exists(path_contratos):
    path_contratos = "Controle de Contratos - Atualizado 2026.xlsx"

  try:
    df_btg_contratos = pd.read_excel(path_contratos, sheet_name="BTG", header=1)
    df_xp_contratos = pd.read_excel(path_contratos, sheet_name="XP", header=1)

    df_btg_contratos["Corretora"] = "BTG"
    df_xp_contratos["Corretora"] = "XP"

    # Identifica estritamente as colunas que são datas (contêm '/' no nome) a partir de 31/01/2023
    def obter_colunas_datas(df):
      cols = list(df.columns)
      if "31/01/2023" in cols:
        idx = cols.index("31/01/2023")
        # Pega todas as colunas a partir de 31/01/2023 que possuem '/' no nome (formato de data DD/MM/AAAA)
        return [c for c in cols[idx:] if "/" in str(c)]
      return []

    col_datas_btg = obter_colunas_datas(df_btg_contratos)
    col_datas_xp = obter_colunas_datas(df_xp_contratos)

    # Limpa e converte os valores monetários dessas colunas de data em ambos os dataframes
    df_btg_contratos = limpar_e_converter(df_btg_contratos, col_datas_btg)
    df_xp_contratos = limpar_e_converter(df_xp_contratos, col_datas_xp)

    contratos_geral = pd.concat(
        [df_btg_contratos, df_xp_contratos], ignore_index=True
    )

  except Exception as e:
    st.error(f"Erro ao ler o Controle de Contratos: {e}")

  # Normalizando as contas
  if "Conta" in contratos_geral.columns:
    contratos_geral["Conta_Str"] = (
        contratos_geral["Conta"]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.strip()
    )

  # Leitura de Posição BTG
  path_pos_btg = (
      "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
  )
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
      else:
        btg_pos = df_pos_btg_raw
  except Exception as e:
    st.warning(f"Aviso ao ler Posição BTG: {e}")

  # Leitura de XP Financeiro
  arquivos_xp = glob.glob("data/relatorio_posicoes_consolidada*.xlsx") + glob.glob(
      "relatorio_posicoes_consolidada*.xlsx"
  )
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
        else:
          xp_fin = df_xp_fin_raw
    except Exception as e:
      st.warning(f"Aviso ao ler o relatório da XP: {e}")

  return contratos_geral, xp_fin, btg_pos


import glob

# Carregando os dados corrigidos
contratos_df, xp_fin_df, btg_pos_df = carregar_dados()

if not contratos_df.empty:
  # ==========================================
  # FILTROS LATERAIS (CORRETORA E CARTEIRA)
  # ==========================================
  st.sidebar.header("Filtros Globais")

  corretora_selecionada = st.sidebar.selectbox(
      "Selecione a Corretora", ["Ambas", "BTG", "XP"]
  )

  if corretora_selecionada != "Ambas":
    contratos_filtrados = contratos_df[
        contratos_df["Corretora"] == corretora_selecionada
    ]
  else:
    contratos_filtrados = contratos_df

  if "Carteira" in contratos_filtrados.columns:
    carteiras_disponiveis = [
        str(c) for c in contratos_filtrados["Carteira"].dropna().unique()
    ]
    carteiras_disponiveis.sort()
    carteira_selecionada = st.sidebar.selectbox(
        "Selecione a Carteira", ["Todas"] + carteiras_disponiveis
    )

    if carteira_selecionada != "Todas":
      contratos_filtrados = contratos_filtrados[
          contratos_filtrados["Carteira"].astype(str) == carteira_selecionada
      ]

  # Identificando estritamente colunas de meses que contêm '/' a partir de '31/01/2023'
  todas_colunas = list(contratos_filtrados.columns)
  if "31/01/2023" in todas_colunas:
    idx_inicio = todas_colunas.index("31/01/2023")
    colunas_meses = [
        col
        for col in todas_colunas[idx_inicio:]
        if "/" in str(col)
        and pd.api.types.is_numeric_dtype(contratos_filtrados[col])
    ]
  else:
    colunas_meses = [
        col
        for col in contratos_filtrados.columns
        if "/" in str(col)
        and pd.api.types.is_numeric_dtype(contratos_filtrados[col])
    ]

  # ==========================================
  # CRIAÇÃO DAS ABAS DO DASHBOARD
  # ==========================================
  aba_geral, aba_evolucao = st.tabs(
      ["Visão Geral", "Evolução Patrimonial por Mês"]
  )

  with aba_geral:
    st.subheader(f"Visão Geral - Corretora: {corretora_selecionada}")

    col1, col2 = st.columns(2)
    col1.metric("Total de Contratos (Com Match)", len(contratos_filtrados))
    col2.metric("Corretora Selecionada", corretora_selecionada)

    st.markdown("### Amostra dos Dados de Contratos")

    if colunas_meses:
      df_exibicao = contratos_filtrados.head(10).copy()
      for col in colunas_meses:
        df_exibicao[col] = df_exibicao[col].apply(
            lambda x: f"R$ {x:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
            if pd.notnull(x)
            else ""
        )
      st.dataframe(df_exibicao, use_container_width=True)
    else:
      st.dataframe(contratos_filtrados.head(10), use_container_width=True)

  with aba_evolucao:
    st.subheader("📈 Evolução Patrimonial por Mês (A partir de 31/01/2023)")
    st.markdown(
        "Análise temporal com base nas colunas financeiras mensais a partir de"
        " Janeiro de 2023."
    )

    if colunas_meses:
      soma_mensal = contratos_filtrados[colunas_meses].sum()

      if not soma_mensal.empty and soma_mensal.sum() > 0:
        fig, ax = plt.subplots(figsize=(11, 5))
        ax.plot(
            soma_mensal.index,
            soma_mensal.values,
            marker="o",
            linestyle="-",
            color="#1f77b4",
            linewidth=2,
        )
        ax.set_title(
            f"Evolução Patrimonial Mensal ({corretora_selecionada}"
            f" {f'- Carteira: {carteira_selecionada}' if 'carteira_selejonada' in locals() and carteira_selecionada != 'Todas' else ''})",
            fontsize=13,
        )
        ax.set_xlabel("Meses", fontsize=11)
        ax.set_ylabel("Valor Consolidado (R$)", fontsize=11)

        # Ajustando o eixo Y para saltos de 100.000.000,00 (100 milhões)
        ax.yaxis.set_major_locator(ticker.MultipleLocator(100_000_000))
        ax.yaxis.set_major_formatter(
            ticker.FuncFormatter(lambda x, p: f"R$ {x*1e-6:.0f}M")
        )

        plt.xticks(rotation=45)
        ax.grid(True, linestyle="--", alpha=0.6)
        st.pyplot(fig)

        st.markdown("### Tabela de Valores Mensais")
        df_tabela_mensal = soma_mensal.reset_index(
            name="Patrimônio Consolidado"
        )
        df_tabela_mensal["Patrimônio Consolidado"] = df_tabela_mensal[
            "Patrimônio Consolidado"
        ].apply(
            lambda x: f"R$ {x:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

        st.dataframe(df_tabela_mensal, use_container_width=True)
      else:
        st.info(
            "Não foram encontrados dados numéricos nas colunas mensais"
            " filtradas."
        )
    else:
      st.warning(
          "Não foi possível identificar colunas mensais a partir de 31/01/2023."
      )
else:
  st.warning(
      "Nenhum dado de contratos foi carregado. Verifique se os arquivos Excel"
      " estão corretos na pasta."
  )






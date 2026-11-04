import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard Bluemetrix", page_icon="📈", layout="wide"
)

st.title("📊 Dashboard Bluemetrix - Consolidado")


# Função para carregar e processar os dados diretamente dos Excel (com cache)
@st.cache_data
def carregar_dados():
  try:
    path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
    df_btg_contratos = pd.read_excel(path_contratos, sheet_name="BTG", header=1)
    df_xp_contratos = pd.read_excel(path_contratos, sheet_name="XP", header=1)

    df_btg_contratos["Corretora"] = "BTG"
    df_xp_contratos["Corretora"] = "XP"

    # Unifica os contratos das duas corretoras
    contratos_geral = pd.concat(
        [df_btg_contratos, df_xp_contratos], ignore_index=True
    )

    path_pos_btg = "data/Posição.xlsx"
    path_pos_xp = "data/relatorio_posicoes_consolidada.xlsx"

    btg_pos = (
        pd.read_excel(path_pos_btg, sheet_name="Export")
        if os.path.exists(path_pos_btg)
        else pd.DataFrame()
    )

    xls_xp = pd.ExcelFile(path_pos_xp) if os.path.exists(path_pos_xp) else None
    xp_fin = (
        pd.read_excel(xls_xp, sheet_name="Financeiro")
        if xls_xp and "Financeiro" in xls_xp.sheet_names
        else pd.DataFrame()
    )

    return contratos_geral, xp_fin, btg_pos

  except Exception as e:
    st.error(f"Erro ao carregar os arquivos Excel: {e}")
    return None, None, None


# Carregando os dados
contratos_df, xp_fin_df, btg_pos_df = carregar_dados()

if contratos_df is not None:
  # ==========================================
  # FILTROS LATERAIS (CORRETORA E CARTEIRA)
  # ==========================================
  st.sidebar.header("Filtros Globais")

  # 1. Filtro de Corretora
  corretora_selecionada = st.sidebar.selectbox(
      "Selecione a Corretora", ["Ambas", "BTG", "XP"]
  )

  if corretora_selecionada != "Ambas":
    contratos_filtrados = contratos_df[
        contratos_df["Corretora"] == corretora_selecionada
    ]
  else:
    contratos_filtrados = contratos_df

  # 2. Filtro de Carteira
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

  # Identificando colunas de meses a partir de "31/01/2023"
  todas_colunas = list(contratos_filtrados.columns)
  if "31/01/2023" in todas_colunas:
    idx_inicio = todas_colunas.index("31/01/2023")
    # Pega da coluna "31/01/2023" em diante que sejam numéricas
    colunas_meses = [
        col
        for col in todas_colunas[idx_inicio:]
        if pd.api.types.is_numeric_dtype(contratos_filtrados[col])
    ]
  else:
    # Fallback caso a data exata não exista
    colunas_excluir = [
        "Conta",
        "Cliente",
        "Corretora",
        "Status",
        "Assessor",
        "Produto",
        "Escritório",
        "UF",
        "Operador",
        "Situação",
        "Exceção",
        "Início da Gestão",
        "Data distrato",
        "Carteira",
    ]
    colunas_meses = [
        col
        for col in contratos_filtrados.columns
        if col not in colunas_excluir
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
    col1.metric("Total de Contratos", len(contratos_filtrados))
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
            f" {f'- Carteira: {carteira_selecionada}' if 'carteira_selecionada' in locals() and carteira_selecionada != 'Todas' else ''})",
            fontsize=13,
        )
        ax.set_xlabel("Meses", fontsize=11)
        ax.set_ylabel("Valor Consolidado (R$)", fontsize=11)

        # Ajustando o eixo Y para saltos de 100.000.000,00 (100 milhões)
        ax.yaxis.set_major_locator(ticker.MultipleLocator(100_000_000))
        # Formatando os valores do eixo Y para milhões/bilhões de forma limpa ou padrão moeda
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







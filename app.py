import matplotlib.pyplot as plt
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
    # 1. Leitura do Controle de Contratos (ignorando a primeira linha com header=1)
    path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
    df_btg_contratos = pd.read_excel(path_contratos, sheet_name="BTG", header=1)
    df_xp_contratos = pd.read_excel(path_contratos, sheet_name="XP", header=1)

    df_btg_contratos["Corretora"] = "BTG"
    df_xp_contratos["Corretora"] = "XP"

    # Unifica os contratos das duas corretoras
    contratos_geral = pd.concat(
        [df_btg_contratos, df_xp_contratos], ignore_index=True
    )

    # 2. Leitura das Posições (Opcional)
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


import os

# Carregando os dados
contratos_df, xp_fin_df, btg_pos_df = carregar_dados()

if contratos_df is not None:
  # ==========================================
  # FILTRO LATERAL: CORRETORA
  # ==========================================
  st.sidebar.header("Filtros Globais")
  corretora_selecionada = st.sidebar.selectbox(
      "Selecione a Corretora", ["Ambas", "BTG", "XP"]
  )

  # Filtrando o dataframe de contratos com base na escolha
  if corretora_selecionada != "Ambas":
    contratos_filtrados = contratos_df[
        contratos_df["Corretora"] == corretora_selecionada
    ]
  else:
    contratos_filtrados = contratos_df

  # Identificando colunas descritivas para ignorar na formatação de moeda
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
  ]

  # Colunas que são numéricas e representam meses/datas
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

    # Formatando as colunas de meses para o formato de moeda brasileiro na exibição
    if colunas_meses:
      # Criamos uma cópia apenas para exibição formatada com R$
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
    st.subheader("📈 Evolução Patrimonial por Mês")
    st.markdown(
        "Análise temporal com base nas colunas financeiras mensais do"
        " Controle de Contratos."
    )

    if colunas_meses:
      # Somando os valores financeiros por mês para a(s) corretora(s) selecionada(s)
      soma_mensal = contratos_filtrados[colunas_meses].sum()

      if not soma_mensal.empty and soma_mensal.sum() > 0:
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(
            soma_mensal.index,
            soma_mensal.values,
            marker="o",
            linestyle="-",
            color="#1f77b4",
            linewidth=2,
        )
        ax.set_title(
            f"Evolução Patrimonial Mensal ({corretora_selecionada})",
            fontsize=14,
        )
        ax.set_xlabel("Meses", fontsize=12)
        ax.set_ylabel("Valor Consolidado (R$)", fontsize=12)
        plt.xticks(rotation=45)
        ax.grid(True, linestyle="--", alpha=0.6)

        st.pyplot(fig)

        st.markdown("### Tabela de Valores Mensais")

        # Formatando a tabela de resumo mensal para R$ mantendo os títulos das datas
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
          "Não foi possível identificar colunas numéricas mensais no Controle"
          " de Contratos."
      )







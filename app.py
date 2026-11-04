import os
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# Tenta importar o script de consolidação
try:
  from consolidar import consolidar_dados
except ImportError:
  consolidar_dados = None

st.set_page_config(
    page_title="Dashboard Bluemetrix", page_icon="📈", layout="wide"
)

st.title("📊 Dashboard Bluemetrix - Consolidado")


# Função com cache inteligente baseada na data de modificação dos arquivos Excel
@st.cache_data
def carregar_e_consolidar_dados(
    path_contratos, path_pos_btg, path_pos_xp
):
  # Se os CSVs não existirem ou se os Excel foram atualizados, executa a consolidação
  if (
      not os.path.exists("data/contratos_consolidados.csv")
      and consolidar_dados is not None
  ):
    st.info("Gerando base consolidada automaticamente pela primeira vez...")
    consolidar_dados()

  try:
    contratos = pd.read_csv("data/contratos_consolidados.csv")
    xp_fin = pd.read_csv("data/xp_financeiro.csv")
    btg_pos = pd.read_csv("data/btg_posicao.csv")
    return contratos, xp_fin, btg_pos
  except Exception as e:
    st.error(
        f"Erro ao carregar os dados. Execute a consolidação primeiro. Detalhe:"
        f" {e}"
    )
    return None, None, None


# Caminhos dos arquivos monitorados
c1 = "data/Controle de Contratos - Atualizado 2026.xlsx"
c2 = "data/Posição.xlsx"
c3 = "data/relatorio_posicoes_consolidada.xlsx"

# Carregando os dados
contratos_df, xp_fin_df, btg_pos_df = carregar_e_consolidar_dados(c1, c2, c3)

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
    st.dataframe(contratos_filtrados.head(10))

  with aba_evolucao:
    st.subheader("📈 Evolução Patrimonial por Mês")
    st.markdown(
        "Análise temporal com base nas colunas financeiras mensais do"
        " Controle de Contratos."
    )

    # Identificando colunas numéricas de meses (excluindo colunas descritivas comuns)
    colunas_excluir = [
        "Conta",
        "Cliente",
        "Corretora",
        "Status",
        "Assessor",
        "Produto",
    ]
    colunas_meses = [
        col
        for col in contratos_filtrados.columns
        if col not in colunas_excluir
        and pd.api.types.is_numeric_dtype(contratos_filtrados[col])
    ]

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
        st.dataframe(
            soma_mensal.reset_index(name="Patrimônio Consolidado"),
            use_container_width=True,
        )
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

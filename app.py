import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard Bluemetrix", page_icon="📈", layout="wide"
)

st.title("📊 Dashboard Bluemetrix - Consolidado")


# Função para carregar os dados salvos pelo script de consolidação
@st.cache_data
def carregar_dados_dashboard():
  try:
    contratos = pd.read_csv("data/contratos_consolidados.csv")
    xp_fin = pd.read_csv("data/xp_financeiro.csv")
    btg_pos = pd.read_csv("data/btg_posicao.csv")
    return contratos, xp_fin, btg_pos
  except Exception as e:
    st.error(
        f"Erro ao carregar os dados. Execute o script de consolidação primeiro."
        f" Detalhe: {e}"
    )
    return None, None, None


contratos_df, xp_fin_df, btg_pos_df = carregar_dados_dashboard()

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

    # Métricas rápidas de exemplo
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Contratos", len(contratos_filtrados))
    col2.metric("Corretoras Ativas", corretora_selecionada)

    st.markdown("### Amostra dos Dados de Contratos")
    st.dataframe(contratos_filtrados.head(10))

  with aba_evolucao:
    st.subheader("📈 Evolução Patrimonial por Mês")
    st.markdown(
        "Análise temporal com base nas colunas financeiras mensais do"
        " Controle de Contratos."
    )

    # DICA: Identifique quais colunas representam os meses na sua planilha de contratos
    # Exemplo genérico buscando colunas que contenham nomes de meses ou anos (ex: 'Jan', 'Fev', '2026', etc.)
    # Substitua a lista abaixo pelas colunas reais de meses da sua planilha.
    colunas_meses = [
        col
        for col in contratos_filtrados.columns
        if col
        not in [
            "Conta",
            "Cliente",
            "Corretora",
            "Status",
        ]  # Ajuste os filtros de colunas textuais
    ]

    # Se você tiver colunas específicas de meses (ex: 'Janeiro', 'Fevereiro', etc.), filtre-as diretamente:
    # colunas_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

    if colunas_meses:
      # Somando os valores financeiros por mês para a(s) corretora(s) selecionada(s)
      soma_mensal = (
          contratos_filtrados[colunas_meses]
          .select_dtypes(include=["number"])
          .sum()
      )

      if not soma_mensal.empty and soma_mensal.sum() > 0:
        # Exibindo Gráfico de Linha da Evolução Patrimonial
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
        st.dataframe(soma_mensal.reset_index(name="Patrimônio Consolidado"))
      else:
        st.info(
            "Não foram encontrados dados numéricos válidos nas colunas de"
            " meses selecionadas."
        )
    else:
      st.warning(
          "Não foi possível identificar automaticamente as colunas de meses"
          " no Controle de Contratos."
      )

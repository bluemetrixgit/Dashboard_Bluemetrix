import pandas as pd


def consolidar_dados():
  print("Iniciando a consolidação dos dados...")

  # 1. Leitura do Controle de Contratos (onde estão as colunas financeiras mensais)
  path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
  df_btg_contratos = pd.read_excel(path_contratos, sheet_name="BTG")
  df_xp_contratos = pd.read_excel(path_contratos, sheet_name="XP")

  # Adicionamos uma coluna de identificação da corretora em cada dataframe
  df_btg_contratos["Corretora"] = "BTG"
  df_xp_contratos["Corretora"] = "XP"

  # 2. Leitura das Posições BTG e XP
  path_pos_btg = "data/Posição.xlsx"
  path_pos_xp = "data/relatorio_posicoes_consolidada.xlsx"

  df_pos_btg = pd.read_excel(path_pos_btg, sheet_name="Export")
  df_pos_btg["Corretora"] = "BTG"

  xls_xp = pd.ExcelFile(path_pos_xp)
  df_xp_financeiro = pd.read_excel(xls_xp, sheet_name="Financeiro")
  df_xp_financeiro["Corretora"] = "XP"

  # Unificando os contratos de ambas as corretoras para a aba de Evolução Patrimonial
  # (Certifique-se de que os nomes das colunas de meses batem entre as abas BTG e XP)
  df_contratos_geral = pd.concat(
      [df_btg_contratos, df_xp_contratos], ignore_index=True
  )

  # Salvando bases para o Dashboard ler separadamente ou unificadas
  df_contratos_geral.to_csv("data/contratos_consolidados.csv", index=False)
  df_xp_financeiro.to_csv("data/xp_financeiro.csv", index=False)
  df_pos_btg.to_csv("data/btg_posicao.csv", index=False)

  print("Consolidação concluída com sucesso!")


if __name__ == "__main__":
  consolidar_dados()

import pandas as pd

def consolidar_dados():
    print("Iniciando a consolidação dos dados...")
    
    # 1. Leitura do Controle de Contratos
    path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
    df_btg_contratos = pd.read_excel(path_contratos, sheet_name="BTG")
    df_xp_contratos = pd.read_excel(path_contratos, sheet_name="XP")
    
    # 2. Leitura das Posições BTG e XP
    path_pos_btg = "data/Posição.xlsx"
    path_pos_xp = "data/relatorio_posicoes_consolidada_04112026141010.xlsx"
    
    df_pos_btg = pd.read_excel(path_pos_btg, sheet_name="Export")
    
    # O relatório da XP possui múltiplas abas importantes (Financeiro, Ações, Renda Fixa, Fundos, etc.)
    xls_xp = pd.ExcelFile(path_pos_xp)
    df_xp_financeiro = pd.read_excel(xls_xp, sheet_name="Financeiro")
    df_xp_rf = pd.read_excel(xls_xp, sheet_name="Renda Fixa") if "Renda Fixa" in xls_xp.sheet_names else pd.DataFrame()
    
    # Exemplo de tratamento e consolidação básica:
    print(f"Contratos BTG carregados: {df_btg_contratos.shape}")
    print(f"Contratos XP carregados: {df_xp_contratos.shape}")
    print(f"Posição BTG carregada: {df_pos_btg.shape}")
    print(f"Financeiro XP carregado: {df_xp_financeiro.shape}")
    
    # Aqui você pode realizar merges por código de cliente/conta, somar patrimônios, etc.
    # Exemplo de tabela consolidada salva para o Dashboard:
    consolidado = pd.merge(df_xp_financeiro, df_xp_contratos, left_on="CodigoCliente", right_on="Conta", how="outer")
    
    consolidado.to_csv("data/base_consolidada.csv", index=False)
    print("Consolidação concluída com sucesso! Arquivo salvo em data/base_consolidada.csv")

if __name__ == "__main__":
    consolidar_dados()

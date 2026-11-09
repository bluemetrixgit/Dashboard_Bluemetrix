import os
import io
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(
    page_title="Dashboard Bluemetrix - Posição e Contratos",
    page_icon="📈",
    layout="wide"
)

st.title("📊 Dashboard Bluemetrix - Posição Consolidada de Clientes")

# =========================
# FUNÇÕES AUXILIARES
# =========================

def limpar_conta(coluna):
    """Padroniza a conta removendo decimals .0, espaços e zeros à esquerda."""
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
    return (
        f"R$ {valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

def limpar_coluna_numerica(serie):
    """Limpa strings financeiras (R$, pontos, vírgulas) convertendo para float seguro."""
    if serie.dtype == object:
        serie = (
            serie.astype(str)
            .str.replace("R$", "", regex=False)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.strip()
        )
    return pd.to_numeric(serie, errors="coerce").fillna(0)

def gerar_excel_corretora(df_view_base, colunas_posicao, pos_pivot):
    """Gera o Excel formatado para as abas BTG e XP com PL Total e Percentuais."""
    df_excel = df_view_base.copy()
    
    if "Conta" in df_excel.columns and "Conta_Limpa" not in df_excel.columns:
        df_excel["Conta_Limpa"] = limpar_conta(df_excel["Conta"])

    if not pos_pivot.empty and "Conta_Limpa" in df_excel.columns:
        df_excel = df_excel.merge(pos_pivot, on="Conta_Limpa", how="left")

    for col in colunas_posicao:
        if col in df_excel.columns:
            df_excel[col] = pd.to_numeric(df_excel[col], errors="coerce").fillna(0.0)
        else:
            df_excel[col] = 0.0

    cols_existentes_pos = [c for c in colunas_posicao if c in df_excel.columns]
    df_excel["PL Total"] = df_excel[cols_existentes_pos].sum(axis=1)

    for col in cols_existentes_pos:
        col_nome_perc = f"{col} %"
        df_excel[col_nome_perc] = df_excel.apply(
            lambda r: (r[col] / r["PL Total"]) if r["PL Total"] > 0 else 0.0, axis=1
        )

    cols = list(df_excel.columns)
    if "Observações" in cols:
        idx_obs = cols.index("Observações")
        if "PL Total" in cols:
            cols.remove("PL Total")
            cols.insert(idx_obs + 1, "PL Total")
        df_excel = df_excel[cols]

    df_export = df_excel.drop(columns=["Conta_Limpa"], errors="ignore")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Consolidado")
    
    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active

    obs_col_idx = None
    for col_idx in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col_idx).value == "Observações":
            obs_col_idx = col_idx
            break

    font_header = Font(name="Calibri", size=11, bold=True)
    border_bottom = Border(bottom=Side(style="medium", color="000000"))
    
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.border = border_bottom
        if obs_col_idx and col_idx > obs_col_idx:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            header_val = str(ws.cell(row=1, column=col_idx).value)
            
            if obs_col_idx and col_idx > obs_col_idx:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if header_val == "PL Total" or header_val in colunas_posicao:
                cell.number_format = 'R$ #,##0.00'
            elif header_val.endswith("%"):
                cell.number_format = '0.00%'

    final_output = io.BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

def gerar_excel_base(df_base):
    """Gera o Excel formatado para a aba Base colocando as colunas de contrato nas primeiras posições."""
    df_excel = df_base.copy()
    
    # Ordem desejada para as primeiras colunas
    cols_contrato_desejadas = ["Corretora", "Cliente", "Conta", "Status", "Situação", "Carteira", "Observações"]
    
    # Identifica as colunas restantes (que vieram da posição)
    colunas_atuais = [c for c in df_excel.columns if c != "Conta_Limpa"]
    cols_extras = [c for c in colunas_atuais if c not in cols_contrato_desejadas]
    
    # Monta a nova ordem com as colunas de contrato primeiro
    cols_ordenadas = [c for c in cols_contrato_desejadas if c in colunas_atuais] + cols_extras
    df_export = df_excel[cols_ordenadas]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Base_Posicao")
    
    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws = wb.active

    font_header = Font(name="Calibri", size=11, bold=True)
    border_bottom = Border(bottom=Side(style="medium", color="000000"))
    
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = font_header
        cell.border = border_bottom
        cell.alignment = Alignment(horizontal="left", vertical="center")

    final_output = io.BytesIO()
    wb.save(final_output)
    return final_output.getvalue()

# =========================
# CARREGAR E PROCESSAR DADOS
# =========================
@st.cache_data
def carregar_dados():
    colunas_contrato = [
        "Corretora",
        "Cliente",
        "Conta",
        "Status",
        "Situação",
        "Carteira",
        "Observações"
    ]

    path_contratos = "data/Controle de Contratos - Atualizado 2026.xlsx"
    if not os.path.exists(path_contratos):
        path_contratos = "Controle de Contratos - Atualizado 2026.xlsx"

    if not os.path.exists(path_contratos):
        st.error(f"Arquivo não encontrado: {path_contratos}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    try:
        df_btg = pd.read_excel(path_contratos, sheet_name="BTG", header=1, engine="openpyxl")
        df_xp = pd.read_excel(path_contratos, sheet_name="XP", header=1, engine="openpyxl")

        df_btg["Corretora"] = "BTG"
        df_xp["Corretora"] = "XP"

        df_btg = df_btg.loc[:, ~df_btg.columns.duplicated()]
        df_xp = df_xp.loc[:, ~df_xp.columns.duplicated()]

        cols_btg = [c for c in colunas_contrato if c in df_btg.columns]
        cols_xp = [c for c in colunas_contrato if c in df_xp.columns]

        contratos_geral = pd.concat([df_btg[cols_btg], df_xp[cols_xp]], ignore_index=True)

        if "Status" in contratos_geral.columns:
            contratos_geral = contratos_geral.dropna(subset=["Status"])
            contratos_geral = contratos_geral[
                contratos_geral["Status"].astype(str).str.strip().str.lower() != "encerrado"
            ]
            contratos_geral = contratos_geral[
                contratos_geral["Status"].astype(str).str.strip() != ""
            ]

        if "Conta" in contratos_geral.columns:
            contratos_geral["Conta_Limpa"] = limpar_conta(contratos_geral["Conta"])
    except Exception as e:
        st.error(f"Erro ao ler Controle de Contratos: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # -------------------------
    # POSIÇÃO BTG
    # -------------------------
    path_pos_btg = "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
    btg_pos_pivot = pd.DataFrame()
    posicao_global_btg = pd.DataFrame()
    base_pos_btg = pd.DataFrame()

    if os.path.exists(path_pos_btg):
        try:
            posicao = pd.read_excel(path_pos_btg, sheet_name="Export", engine="openpyxl")
            posicao.columns = posicao.columns.str.strip()

            col_conta_btg = "CodigoCliente" if "CodigoCliente" in posicao.columns else "Conta"
            if col_conta_btg in posicao.columns:
                posicao["Conta_Limpa"] = limpar_conta(posicao[col_conta_btg])

                df_contratos_btg_ref = df_btg[cols_btg].copy()
                df_contratos_btg_ref["Conta_Limpa"] = limpar_conta(df_contratos_btg_ref["Conta"])
                base_pos_btg = posicao.merge(df_contratos_btg_ref.drop(columns=["Conta"], errors="ignore"), on="Conta_Limpa", how="left")

                def definir_coluna_btg(row):
                    submercado = str(row.get("Sub Mercado", "")).strip().upper()
                    produto = str(row.get("Produto", "")).strip().upper()
                    mercado = str(row.get("Mercado", "")).strip().lower()

                    if "TESOURO DIRETO" in submercado:
                        return "Tesouro Direto"
                    elif "TITULO PUBLICO" in submercado or "TÍTULO PÚBLICO" in submercado:
                        return "Título Público"
                    elif "PRIVADO" in submercado:
                        return "CrPr"
                    
                    alvos_lc = ["LCI", "LCA", "LCD"]
                    if any(item in submercado or item in produto for item in alvos_lc):
                        return "LC"
                        
                    alvos_lf = ["LFSC", "LFSN", "LF"]
                    if submercado in alvos_lf or produto in alvos_lf or any(f" {item} " in f" {submercado} " for item in alvos_lf):
                        return "LF"
                    
                    if "renda fixa" in mercado:
                        return str(row.get("Produto", "Outros RF"))
                    elif "renda variável" in mercado or "renda variavel" in mercado:
                        return str(row.get("Sub Mercado", "Outros RV"))
                    
                    return f"{row.get('Mercado', 'Outros')}"

                posicao["Coluna_Mercado"] = posicao.apply(definir_coluna_btg, axis=1)

                col_vb = "Valor Bruto" if "Valor Bruto" in posicao.columns else [c for c in posicao.columns if "valor" in c.lower()][0]
                posicao["Valor_Bruto_Num"] = limpar_coluna_numerica(posicao[col_vb])

                posicao_global_btg = posicao.copy()

                btg_pos_pivot = posicao.pivot_table(
                    index="Conta_Limpa",
                    columns="Coluna_Mercado",
                    values="Valor_Bruto_Num",
                    aggfunc="sum",
                    fill_value=0
                ).reset_index()

                colunas_para_remover = [c for c in btg_pos_pivot.columns if pd.isna(c) or str(c).lower() in ["nan", "none", "valor em trânsito", "valor em transito"]]
                btg_pos_pivot = btg_pos_pivot.drop(columns=colunas_para_remover, errors="ignore")

                colunas_para_agrupar = [
                    "Conta Corrente", "COE", "Criptoativos", 
                    "Derivativos", "Fundos Cetipados", "OPCAO", "Tesouro Direto"
                ]
                cols_existentes = [c for c in colunas_para_agrupar if c in btg_pos_pivot.columns]
                
                if cols_existentes:
                    btg_pos_pivot["Outros"] = btg_pos_pivot[cols_existentes].sum(axis=1)
                    btg_pos_pivot = btg_pos_pivot.drop(columns=cols_existentes)

                    conta_limpa_col = btg_pos_pivot.pop("Conta_Limpa")
                    outros_col = btg_pos_pivot.pop("Outros")
                    btg_pos_pivot["Conta_Limpa"] = conta_limpa_col
                    btg_pos_pivot["Outros"] = outros_col

        except Exception as e:
            st.warning(f"Aviso ao ler Posição BTG: {e}")

    # -------------------------
    # POSIÇÃO XP
    # -------------------------
    candidatos_xp = [
        "data/relatório_posicoes_consolidada.xlsx",
        "relatório_posicoes_consolidada.xlsx",
        "data/relatorio_posicoes_consolidada.xlsx",
        "relatorio_posicoes_consolidada.xlsx",
        "data/Posição XP.xlsx",
        "Posição XP.xlsx"
    ]

    path_pos_xp = None
    for p in candidatos_xp:
        if os.path.exists(p):
            path_pos_xp = p
            break

    xp_pos_pivot = pd.DataFrame()
    posicao_global_xp = pd.DataFrame()
    base_pos_xp = pd.DataFrame()

    abas_mapeamento_xp = {
        "Ações": "ValorAtual",
        "Aluguéis": "ValorAtual",
        "Coe": "ValorFinanceiroBruto",
        "Custódia Remunerada": "Custódia Remunerada",
        "Fundos": "ValorBruto",
        "Fundos Imobiliários": "ValorAtual",
        "Garantias": "ValorAtual",
        "Opções Flexíveis": "Posicao",
        "Previdência": "ValorReservaAcumulada",
        "Proventos": "PrecoAtual",
        "Proventos Fundo Imob": "ValorLiquidoAtual",
        "Renda Fixa": "ValorFinanceiroBruto",
        "Tesouro Direto": "ValorBruto"
    }

    if path_pos_xp:
        try:
            xls = pd.ExcelFile(path_pos_xp, engine="openpyxl")
            dfs_por_aba = []
            dfs_globais = []
            dfs_base_xp = []

            df_contratos_xp_ref = df_xp[cols_xp].copy()
            df_contratos_xp_ref["Conta_Limpa"] = limpar_conta(df_contratos_xp_ref["Conta"])

            for aba_nome, col_alvo in abas_mapeamento_xp.items():
                if aba_nome in xls.sheet_names:
                    df_aba = pd.read_excel(xls, sheet_name=aba_nome)
                    df_aba.columns = df_aba.columns.str.strip()

                    col_conta_xp = None
                    for candidato in ["CodigoCliente", "Código do Cliente", "Codigo Cliente", "Conta"]:
                        if candidato in df_aba.columns:
                            col_conta_xp = candidato
                            break

                    if col_conta_xp:
                        df_aba["Conta_Limpa"] = limpar_conta(df_aba[col_conta_xp])
                        df_aba["Aba_Posicao"] = aba_nome

                        df_aba_base = df_aba.merge(df_contratos_xp_ref.drop(columns=["Conta"], errors="ignore"), on="Conta_Limpa", how="left")
                        dfs_base_xp.append(df_aba_base)

                        if col_alvo in df_aba.columns:
                            df_aba["Valor_Calculado"] = limpar_coluna_numerica(df_aba[col_alvo])
                        else:
                            cands = [c for c in df_aba.columns if col_alvo.lower() in c.lower()]
                            if cands:
                                df_aba["Valor_Calculado"] = limpar_coluna_numerica(df_aba[cands[0]])
                            else:
                                df_aba["Valor_Calculado"] = 0

                        if aba_nome in ["Coe", "Renda Fixa", "Tesouro Direto"]:
                            col_venc_encontrada = None
                            for c in df_aba.columns:
                                if "datavencimento" in c.lower().replace(" ", "").replace("_", ""):
                                    col_venc_encontrada = c
                                    break
                            
                            if col_venc_encontrada:
                                df_aba["Data_Vencimento_Calculada"] = pd.to_datetime(df_aba[col_venc_encontrada], errors="coerce")
                            else:
                                df_aba["Data_Vencimento_Calculada"] = pd.NaT
                        else:
                            df_aba["Data_Vencimento_Calculada"] = pd.NaT

                        dfs_globais.append(df_aba)

                        if aba_nome == "Renda Fixa" and "Categoria" in df_aba.columns:
                            def classificar_rf(row):
                                cat_str = str(row.get("Categoria", "")).strip().upper()
                                if cat_str == "CDB":
                                    return "CDB"
                                elif cat_str in ["CDCA", "CRI", "CRA", "DEB", "CPRF", "LF"]:
                                    return "CrPr"
                                elif cat_str in ["LCA", "LCI", "LCD"]:
                                    return "LC"
                                elif cat_str in ["FND", "FUNDOS"]:
                                    return "Fundos"
                                elif cat_str in ["LFT", "NTN-B", "NTN-F"]:
                                    return "Títulos Públicos"
                                else:
                                    return "Outros RF"

                            df_aba["Categoria_RF"] = df_aba.apply(classificar_rf, axis=1)
                            df_res_rf = df_aba.groupby(["Conta_Limpa", "Categoria_RF"])["Valor_Calculado"].sum().unstack(fill_value=0).reset_index()
                            dfs_por_aba.append(df_res_rf)
                        elif aba_nome in ["Fundos", "Fundos_X", "Fundos_Y"]:
                            df_res_aba = df_aba.groupby("Conta_Limpa")["Valor_Calculado"].sum().reset_index()
                            df_res_aba.rename(columns={"Valor_Calculado": "Fundos"}, inplace=True)
                            dfs_por_aba.append(df_res_aba)
                        else:
                            df_res_aba = df_aba.groupby("Conta_Limpa")["Valor_Calculado"].sum().reset_index()
                            df_res_aba.rename(columns={"Valor_Calculado": aba_nome}, inplace=True)
                            dfs_por_aba.append(df_res_aba)

            if dfs_por_aba:
                from functools import reduce
                xp_pos_pivot = reduce(lambda left, right: pd.merge(left, right, on="Conta_Limpa", how="outer"), dfs_por_aba).fillna(0)
                
                cols_fundos_encontradas = [c for c in xp_pos_pivot.columns if str(c).lower() in ["fundos", "fundos_x", "fundos_y", "fnd"]]
                if cols_fundos_encontradas:
                    xp_pos_pivot["Fundos"] = xp_pos_pivot[cols_fundos_encontradas].sum(axis=1)
                    cols_para_remover = [c for c in cols_fundos_encontradas if c != "Fundos"]
                    xp_pos_pivot = xp_pos_pivot.drop(columns=cols_para_remover, errors="ignore")

            if dfs_globais:
                posicao_global_xp = pd.concat(dfs_globais, ignore_index=True)

            if dfs_base_xp:
                base_pos_xp = pd.concat(dfs_base_xp, ignore_index=True)

        except Exception as e:
            st.warning(f"Aviso ao ler relatório de posições da XP: {e}")

    return contratos_geral, btg_pos_pivot, posicao_global_btg, xp_pos_pivot, posicao_global_xp, base_pos_btg, base_pos_xp

# =========================
# MONTAGEM DA INTERFACE
# =========================

contratos_df, btg_pos_pivot, posicao_global_btg, xp_pos_pivot, posicao_global_xp, base_pos_btg, base_pos_xp = carregar_dados()

if not contratos_df.empty:
    st.sidebar.title("Navegação")
    corretora_selecionada = st.sidebar.radio("Selecione a Corretora", ["BTG", "XP", "Base"])
    st.sidebar.markdown("---")

    if corretora_selecionada == "BTG":
        # =========================================================================
        # SEÇÃO BTG
        # =========================================================================
        st.subheader("🔵 Corretora BTG - Contratos & Posições")

        df_base_corretora = contratos_df[contratos_df["Corretora"] == "BTG"].copy()

        st.sidebar.header("Filtros BTG")

        status_disponiveis = sorted([str(s) for s in df_base_corretora["Status"].dropna().unique()]) if "Status" in df_base_corretora.columns else []
        status_selecionados = st.sidebar.multiselect("Filtrar por Status", status_disponiveis, key="status_btg_sel")
        
        df_f1 = df_base_corretora.copy()
        if status_selecionados:
            df_f1 = df_f1[df_f1["Status"].astype(str).isin(status_selecionados)]

        carteiras_disponiveis = sorted([str(c) for c in df_f1["Carteira"].dropna().unique()]) if "Carteira" in df_f1.columns else []
        carteiras_selecionadas = st.sidebar.multiselect("Filtrar por Carteira", carteiras_disponiveis, key="cart_btg_sel")
        
        df_f2 = df_f1.copy()
        if carteiras_selecionadas:
            df_f2 = df_f2[df_f2["Carteira"].astype(str).isin(carteiras_selecionadas)]

        contas_disponiveis = sorted([str(c) for c in df_f2["Conta"].dropna().unique()]) if "Conta" in df_f2.columns else []
        contas_selecionadas = st.sidebar.multiselect("Filtrar por Conta", contas_disponiveis, key="conta_btg_sel")
        
        df_btg_view = df_f2.copy()
        if contas_selecionadas:
            df_btg_view = df_btg_view[df_btg_view["Conta"].astype(str).isin(contas_selecionadas)]

        st.sidebar.markdown("---")
        st.sidebar.header("Filtro Vencimentos")
        produtos_disponiveis = []
        if not posicao_global_btg.empty and "Produto" in posicao_global_btg.columns:
            produtos_disponiveis = sorted([str(p) for p in posicao_global_btg["Produto"].dropna().unique()])
        produtos_selecionados = st.sidebar.multiselect("Filtrar por Produto", produtos_disponiveis, key="venc_prod_sel_btg")

        st.metric("Total de Contratos (BTG)", len(df_btg_view))

        colunas_posicao = []
        if not btg_pos_pivot.empty:
            df_btg_view_pos = df_base_corretora.merge(btg_pos_pivot, on="Conta_Limpa", how="left")
            if status_selecionados:
                df_btg_view_pos = df_btg_view_pos[df_btg_view_pos["Status"].astype(str).isin(status_selecionados)]
            if carteiras_selecionadas:
                df_btg_view_pos = df_btg_view_pos[df_btg_view_pos["Carteira"].astype(str).isin(carteiras_selecionadas)]
            if contas_selecionadas:
                df_btg_view_pos = df_btg_view_pos[df_btg_view_pos["Conta"].astype(str).isin(contas_selecionadas)]

            colunas_posicao = [c for c in btg_pos_pivot.columns if c != "Conta_Limpa"]
            for col in colunas_posicao:
                df_btg_view_pos[col] = df_btg_view_pos[col].fillna(0)
        else:
            df_btg_view_pos = df_btg_view

        df_exibir_btg = df_btg_view.copy()
        if not btg_pos_pivot.empty:
            df_exibir_btg = df_exibir_btg.merge(btg_pos_pivot, on="Conta_Limpa", how="left")
            for col in colunas_posicao:
                if col in df_exibir_btg.columns:
                    df_exibir_btg[col] = df_exibir_btg[col].apply(formatar_real)

        if "Conta_Limpa" in df_exibir_btg.columns:
            df_exibir_btg = df_exibir_btg.drop(columns=["Conta_Limpa"])

        st.dataframe(df_exibir_btg, use_container_width=True)

        st.download_button(
            label="📥 Baixar Planilha BTG em Excel (com PL Total e %)",
            data=gerar_excel_corretora(df_btg_view, colunas_posicao, btg_pos_pivot),
            file_name="posicao_contratos_btg.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("---")

        if colunas_posicao:
            st.markdown("### 📊 Consolidado de Valor Bruto por Mercado/Produto (BTG)")
            soma_posicoes = df_btg_view_pos[colunas_posicao].sum() / 1e6

            if not soma_posicoes.empty and soma_posicoes.sum() > 0:
                fig, ax = plt.subplots(figsize=(12, 5))
                fig.patch.set_facecolor("#0e1117")
                ax.set_facecolor("#0e1117")

                ax.bar(
                    soma_posicoes.index,
                    soma_posicoes.values,
                    color="#1f77b4",
                    edgecolor="#ffffff",
                    linewidth=0.5,
                    width=0.6
                )

                ax.set_title("Valor Bruto por Produto / Sub Mercado (R$ Milhões)", color="white", fontsize=12, pad=15)
                ax.set_ylabel("Valor Bruto (R$ Milhões)", color="white", fontsize=10)
                plt.xticks(rotation=45, ha="right", color="white", fontsize=9)
                ax.tick_params(colors="white")
                ax.xaxis.label.set_color("white")
                ax.yaxis.label.set_color("white")

                for spine in ax.spines.values():
                    spine.set_edgecolor("white")

                ax.grid(True, linestyle="--", alpha=0.2, color="white")
                plt.tight_layout()
                st.pyplot(fig)

                st.markdown(
                    "> **ℹ️ Legenda da categoria 'Outros':** A coluna *Outros* consolida os valores de: "
                    "*Conta Corrente, COE, Criptoativos, Derivativos, Fundos Cetipados, OPCAO e Tesouro Direto*."
                )

                st.markdown("### 🥧 Distribuição Percentual por Categoria (BTG)")
                soma_pizza = soma_posicoes[soma_posicoes > 0]

                if not soma_pizza.empty:
                    fig_pie, ax_pie = plt.subplots(figsize=(7, 4))
                    fig_pie.patch.set_facecolor("#0e1117")
                    ax_pie.set_facecolor("#0e1117")

                    cores_neutras = [
                        "#1f77b4", "#4682b4", "#5f9ea0", "#708090", 
                        "#5c6b73", "#3a506b", "#4a5568", "#2d3748", "#4a607a"
                    ]

                    wedges, texts, autotexts = ax_pie.pie(
                        soma_pizza.values,
                        autopct='%1.1f%%',
                        startangle=140,
                        colors=cores_neutras[:len(soma_pizza)],
                        textprops=dict(color="white", fontsize=8),
                        pctdistance=0.75
                    )

                    for autotext in autotexts:
                        autotext.set_color('white')
                        autotext.set_fontsize(7.5)

                    legend_pie = ax_pie.legend(
                        wedges, 
                        soma_pizza.index,
                        title="Categorias",
                        loc="center left",
                        bbox_to_anchor=(1, 0, 0.5, 1),
                        facecolor="#161b22",
                        edgecolor="white",
                        labelcolor="white",
                        fontsize=8
                    )
                    legend_pie.get_title().set_color('white')

                    ax_pie.set_title("Participação Percentual por Categoria", color="white", fontsize=11, pad=12)
                    plt.tight_layout()
                    st.pyplot(fig_pie)

                st.markdown("### 📉 Gráfico de Dispersão - Descolamento das Aplicações (BTG)")
                
                df_disp_btg = btg_pos_pivot.merge(df_base_corretora[["Conta_Limpa", "Carteira", "Status", "Cliente"]], on="Conta_Limpa", how="inner")
                if status_selecionados:
                    df_disp_btg = df_disp_btg[df_disp_btg["Status"].astype(str).isin(status_selecionados)]
                if carteiras_selecionadas:
                    df_disp_btg = df_disp_btg[df_disp_btg["Carteira"].astype(str).isin(carteiras_selecionadas)]
                if contas_selecionadas:
                    df_disp_btg = df_disp_btg[df_disp_btg["Conta_Limpa"].isin(limpar_conta(pd.Series(contas_selecionadas)))]

                if not df_disp_btg.empty and colunas_posicao:
                    df_melt_btg = df_disp_btg.melt(
                        id_vars=["Conta_Limpa", "Carteira", "Status", "Cliente"],
                        value_vars=colunas_posicao,
                        var_name="Tipo_Aplicacao",
                        value_name="Valor_Aplicado"
                    )
                    df_melt_btg = df_melt_btg[df_melt_btg["Valor_Aplicado"] > 0]

                    if not df_melt_btg.empty:
                        media_geral_aplicacao = df_melt_btg.groupby("Tipo_Aplicacao")["Valor_Aplicado"].mean().reset_index()
                        media_geral_aplicacao.rename(columns={"Valor_Aplicado": "Media_Geral"}, inplace=True)
                        
                        df_melt_btg = df_melt_btg.merge(media_geral_aplicacao, on="Tipo_Aplicacao", how="left")
                        df_melt_btg["Desvio_Media"] = (df_melt_btg["Valor_Aplicado"] - df_melt_btg["Media_Geral"]).abs()

                        fig_scatter = px.scatter(
                            df_melt_btg,
                            x="Tipo_Aplicacao",
                            y="Valor_Aplicado",
                            color="Carteira",
                            size="Desvio_Media",
                            hover_data=["Conta_Limpa", "Cliente", "Status", "Media_Geral"],
                            title="Dispersão de Aplicações por Carteira (Tamanho = Descolamento da Média Geral)",
                            labels={"Tipo_Aplicacao": "Tipo de Aplicação", "Valor_Aplicado": "Valor Aplicado (R$)"}
                        )
                        fig_scatter.update_layout(
                            plot_bgcolor="#0e1117",
                            paper_bgcolor="#0e1117",
                            font=dict(color="white"),
                            title_font=dict(color="white"),
                            xaxis=dict(tickangle=-45)
                        )
                        st.plotly_chart(fig_scatter, use_container_width=True)

                st.markdown("### 📅 Vencimentos Mensais - BTG (Até Dezembro de 2030)")
                if not posicao_global_btg.empty:
                    df_venc = posicao_global_btg.copy()
                    
                    contas_validas_venc = df_btg_view_pos["Conta_Limpa"].dropna().unique()
                    df_venc = df_venc[df_venc["Conta_Limpa"].isin(contas_validas_venc)]

                    if produtos_selecionados:
                        df_venc = df_venc[df_venc["Produto"].astype(str).isin(produtos_selecionados)]

                    col_venc = next((c for c in df_venc.columns if "venc" in c.lower()), None)
                    if col_venc:
                        df_venc["Data_Vencimento"] = pd.to_datetime(df_venc[col_venc], errors="coerce")
                        df_venc = df_venc.dropna(subset=["Data_Vencimento"])

                        data_inicio = pd.Timestamp.now().normalize()
                        data_fim = pd.to_datetime("2030-12-31")
                        df_venc = df_venc[(df_venc["Data_Vencimento"] >= data_inicio) & (df_venc["Data_Vencimento"] <= data_fim)]

                        if not df_venc.empty:
                            df_venc["Ano_Mes"] = df_venc["Data_Vencimento"].dt.to_period("M").astype(str)
                            venc_agrupado = df_venc.groupby("Ano_Mes")["Valor_Bruto_Num"].sum() / 1e6
                            idx_completo = pd.period_range(start=data_inicio.strftime("%Y-%m"), end="2030-12", freq="M").astype(str)
                            venc_agrupado = venc_agrupado.reindex(idx_completo, fill_value=0)

                            fig_venc, ax_venc = plt.subplots(figsize=(14, 5))
                            fig_venc.patch.set_facecolor("#0e1117")
                            ax_venc.set_facecolor("#0e1117")

                            ax_venc.bar(
                                venc_agrupado.index,
                                venc_agrupado.values,
                                color="#2ca02c",
                                edgecolor="#ffffff",
                                linewidth=0.5,
                                width=0.7
                            )
                            ax_venc.set_title("Valor Bruto dos Vencimentos por Mês (R$ Milhões)", color="white", fontsize=12, pad=15)
                            ax_venc.set_ylabel("Valor Bruto (R$ Milhões)", color="white", fontsize=10)
                            plt.xticks(rotation=90, ha="center", color="white", fontsize=8)
                            ax_venc.tick_params(colors="white")
                            ax_venc.xaxis.label.set_color("white")
                            ax_venc.yaxis.label.set_color("white")
                            for spine in ax_venc.spines.values():
                                spine.set_edgecolor("white")
                            ax_venc.grid(True, linestyle="--", alpha=0.2, color="white")
                            plt.tight_layout()
                            st.pyplot(fig_venc)

    elif corretora_selecionada == "XP":
        # =========================================================================
        # SEÇÃO XP
        # =========================================================================
        st.subheader("🟡 Corretora XP - Contratos & Posições")

        df_base_corretora_xp = contratos_df[contratos_df["Corretora"] == "XP"].copy()

        st.sidebar.header("Filtros XP")

        status_xp_disp = sorted([str(s) for s in df_base_corretora_xp["Status"].dropna().unique()]) if "Status" in df_base_corretora_xp.columns else []
        status_xp_sel = st.sidebar.multiselect("Filtrar por Status", status_xp_disp, key="status_xp_sel")
        
        df_f1_xp = df_base_corretora_xp.copy()
        if status_xp_sel:
            df_f1_xp = df_f1_xp[df_f1_xp["Status"].astype(str).isin(status_xp_sel)]

        carteiras_xp_disp = sorted([str(c) for c in df_f1_xp["Carteira"].dropna().unique()]) if "Carteira" in df_f1_xp.columns else []
        carteiras_xp_sel = st.sidebar.multiselect("Filtrar por Carteira", carteiras_xp_disp, key="cart_xp_sel")
        
        df_f2_xp = df_f1_xp.copy()
        if carteiras_xp_sel:
            df_f2_xp = df_f2_xp[df_f2_xp["Carteira"].astype(str).isin(carteiras_xp_sel)]

        contas_xp_disp = sorted([str(c) for c in df_f2_xp["Conta"].dropna().unique()]) if "Conta" in df_f2_xp.columns else []
        contas_xp_sel = st.sidebar.multiselect("Filtrar por Conta", contas_xp_disp, key="conta_xp_sel")
        
        df_xp_view = df_f2_xp.copy()
        if contas_xp_sel:
            df_xp_view = df_xp_view[df_xp_view["Conta"].astype(str).isin(contas_xp_sel)]

        st.sidebar.markdown("---")
        st.sidebar.header("Filtro Vencimentos")
        
        produtos_xp_disp = []
        if not posicao_global_xp.empty:
            df_rf_apenas = posicao_global_xp[posicao_global_xp["Aba_Origem"] == "Renda Fixa"]
            if "Categoria" in df_rf_apenas.columns:
                produtos_xp_disp = sorted([str(p) for p in df_rf_apenas["Categoria"].dropna().unique()])
            
            if not produtos_xp_disp:
                df_venc_opcoes = posicao_global_xp[posicao_global_xp["Aba_Origem"].isin(["Renda Fixa", "Coe", "Tesouro Direto"])]
                for col_prod in ["Categoria", "Produto", "Ativo", "SubMercado"]:
                    if col_prod in df_venc_opcoes.columns:
                        produtos_xp_disp = sorted([str(p) for p in df_venc_opcoes[col_prod].dropna().unique()])
                        if produtos_xp_disp:
                            break
        
        produtos_xp_sel = st.sidebar.multiselect("Filtrar por Produto/Ativo", produtos_xp_disp, key="venc_prod_sel_xp")

        st.metric("Total de Contratos (XP)", len(df_xp_view))

        colunas_posicao_xp = []
        if not xp_pos_pivot.empty:
            df_xp_view_pos = df_base_corretora_xp.merge(xp_pos_pivot, on="Conta_Limpa", how="left")
            if status_xp_sel:
                df_xp_view_pos = df_xp_view_pos[df_xp_view_pos["Status"].astype(str).isin(status_xp_sel)]
            if carteiras_xp_sel:
                df_xp_view_pos = df_xp_view_pos[df_xp_view_pos["Carteira"].astype(str).isin(carteiras_xp_sel)]
            if contas_xp_sel:
                df_xp_view_pos = df_xp_view_pos[df_xp_view_pos["Conta"].astype(str).isin(contas_xp_sel)]

            colunas_posicao_xp = [c for c in xp_pos_pivot.columns if c != "Conta_Limpa"]
            for col in colunas_posicao_xp:
                df_xp_view_pos[col] = df_xp_view_pos[col].fillna(0)
        else:
            df_xp_view_pos = df_xp_view

        df_exibir_xp = df_xp_view.copy()
        if not xp_pos_pivot.empty:
            df_exibir_xp = df_exibir_xp.merge(xp_pos_pivot, on="Conta_Limpa", how="left")
            for col in colunas_posicao_xp:
                if col in df_exibir_xp.columns:
                    df_exibir_xp[col] = df_exibir_xp[col].apply(formatar_real)

        if "Conta_Limpa" in df_exibir_xp.columns:
            df_exibir_xp = df_exibir_xp.drop(columns=["Conta_Limpa"])

        st.dataframe(df_exibir_xp, use_container_width=True)

        st.download_button(
            label="📥 Baixar Planilha XP em Excel (com PL Total e %)",
            data=gerar_excel_corretora(df_xp_view, colunas_posicao_xp, xp_pos_pivot),
            file_name="posicao_contratos_xp.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("---")

        if colunas_posicao_xp:
            st.markdown("### 📊 Consolidado por Abas/Produtos (XP)")
            soma_pos_xp = df_xp_view_pos[colunas_posicao_xp].sum() / 1e6

            if not soma_pos_xp.empty and soma_pos_xp.sum() > 0:
                fig_xp, ax_xp = plt.subplots(figsize=(12, 5))
                fig_xp.patch.set_facecolor("#0e1117")
                ax_xp.set_facecolor("#0e1117")

                ax_xp.bar(
                    soma_pos_xp.index,
                    soma_pos_xp.values,
                    color="#ff7f0e",
                    edgecolor="#ffffff",
                    linewidth=0.5,
                    width=0.6
                )

                ax_xp.set_title("Valor Bruto por Aba / Produto - XP (R$ Milhões)", color="white", fontsize=12, pad=15)
                ax_xp.set_ylabel("Valor Bruto (R$ Milhões)", color="white", fontsize=10)
                plt.xticks(rotation=45, ha="right", color="white", fontsize=9)
                ax_xp.tick_params(colors="white")
                ax_xp.xaxis.label.set_color("white")
                ax_xp.yaxis.label.set_color("white")

                for spine in ax_xp.spines.values():
                    spine.set_edgecolor("white")

                ax_xp.grid(True, linestyle="--", alpha=0.2, color="white")
                plt.tight_layout()
                st.pyplot(fig_xp)

                st.markdown("### 🥧 Distribuição Percentual por Categoria (XP)")
                soma_pizza_xp = soma_pos_xp[soma_pos_xp > 0]

                if not soma_pizza_xp.empty:
                    fig_pie_xp, ax_pie_xp = plt.subplots(figsize=(7, 4))
                    fig_pie_xp.patch.set_facecolor("#0e1117")
                    ax_pie_xp.set_facecolor("#0e1117")

                    cores_suaves_xp = [
                        "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", 
                        "#59a14f", "#edc948", "#b07aa1", "#ff9da7", 
                        "#9c755f", "#bab0ac", "#86bcb6", "#e1974c", "#849ba8"
                    ]

                    wedges_xp, texts_xp, autotexts_xp = ax_pie_xp.pie(
                        soma_pizza_xp.values,
                        autopct='%1.1f%%',
                        startangle=140,
                        colors=cores_suaves_xp[:len(soma_pizza_xp)],
                        textprops=dict(color="white", fontsize=8),
                        pctdistance=0.75
                    )

                    for autotext in autotexts_xp:
                        autotext.set_color('white')
                        autotext.set_fontsize(7.5)

                    legend_pie_xp = ax_pie_xp.legend(
                        wedges_xp, 
                        soma_pizza_xp.index,
                        title="Categorias XP",
                        loc="center left",
                        bbox_to_anchor=(1, 0, 0.5, 1),
                        facecolor="#161b22",
                        edgecolor="white",
                        labelcolor="white",
                        fontsize=8
                    )
                    legend_pie_xp.get_title().set_color('white')

                    ax_pie_xp.set_title("Participação Percentual por Categoria (XP)", color="white", fontsize=11, pad=12)
                    plt.tight_layout()
                    st.pyplot(fig_pie_xp)

                st.markdown("### 📉 Gráfico de Dispersão - Descolamento das Aplicações (XP)")
                
                df_disp_xp = xp_pos_pivot.merge(df_base_corretora_xp[["Conta_Limpa", "Carteira", "Status", "Cliente"]], on="Conta_Limpa", how="inner")
                if status_xp_sel:
                    df_disp_xp = df_disp_xp[df_disp_xp["Status"].astype(str).isin(status_xp_sel)]
                if carteiras_xp_sel:
                    df_disp_xp = df_disp_xp[df_disp_xp["Carteira"].astype(str).isin(carteiras_xp_sel)]
                if contas_xp_sel:
                    df_disp_xp = df_disp_xp[df_disp_xp["Conta_Limpa"].isin(limpar_conta(pd.Series(contas_xp_sel)))]

                if not df_disp_xp.empty and colunas_posicao_xp:
                    df_melt_xp = df_disp_xp.melt(
                        id_vars=["Conta_Limpa", "Carteira", "Status", "Cliente"],
                        value_vars=colunas_posicao_xp,
                        var_name="Tipo_Aplicacao",
                        value_name="Valor_Aplicado"
                    )
                    df_melt_xp = df_melt_xp[df_melt_xp["Valor_Aplicado"] > 0]

                    if not df_melt_xp.empty:
                        media_geral_aplicacao_xp = df_melt_xp.groupby("Tipo_Aplicacao")["Valor_Aplicado"].mean().reset_index()
                        media_geral_aplicacao_xp.rename(columns={"Valor_Aplicado": "Media_Geral"}, inplace=True)
                        
                        df_melt_xp = df_melt_xp.merge(media_geral_aplicacao_xp, on="Tipo_Aplicacao", how="left")
                        df_melt_xp["Desvio_Media"] = (df_melt_xp["Valor_Aplicado"] - df_melt_xp["Media_Geral"]).abs()

                        fig_scatter_xp = px.scatter(
                            df_melt_xp,
                            x="Tipo_Aplicacao",
                            y="Valor_Aplicado",
                            color="Carteira",
                            size="Desvio_Media",
                            hover_data=["Conta_Limpa", "Cliente", "Status", "Media_Geral"],
                            title="Dispersão de Aplicações por Carteira (Tamanho = Descolamento da Média Geral)",
                            labels={"Tipo_Aplicacao": "Tipo de Aplicação", "Valor_Aplicado": "Valor Aplicado (R$)"}
                        )
                        fig_scatter_xp.update_layout(
                            plot_bgcolor="#0e1117",
                            paper_bgcolor="#0e1117",
                            font=dict(color="white"),
                            title_font=dict(color="white"),
                            xaxis=dict(tickangle=-45)
                        )
                        st.plotly_chart(fig_scatter_xp, use_container_width=True)

                st.markdown("### 📅 Vencimentos Mensais - XP (Até Dezembro de 2030)")
                if not posicao_global_xp.empty:
                    df_venc_xp = posicao_global_xp.copy()
                    
                    contas_validas_venc_xp = df_xp_view_pos["Conta_Limpa"].dropna().unique()
                    df_venc_xp = df_venc_xp[df_venc_xp["Conta_Limpa"].isin(contas_validas_venc_xp)]

                    df_venc_xp = df_venc_xp[df_venc_xp["Aba_Origem"].isin(["Renda Fixa", "Coe", "Tesouro Direto"])].copy()
                    
                    if produtos_xp_sel:
                        col_filtragem = "Categoria" if "Categoria" in df_venc_xp.columns else next((c for c in ["Produto", "Ativo", "SubMercado"] if c in df_venc_xp.columns), None)
                        if col_filtragem:
                            df_venc_xp = df_venc_xp[df_venc_xp[col_filtragem].astype(str).isin(produtos_xp_sel)]

                    if "Data_Vencimento_Calculada" in df_venc_xp.columns:
                        df_venc_xp["Data_Vencimento"] = pd.to_datetime(df_venc_xp["Data_Vencimento_Calculada"], errors="coerce")
                        if hasattr(df_venc_xp["Data_Vencimento"].dt, "tz") and df_venc_xp["Data_Vencimento"].dt.tz is not None:
                            df_venc_xp["Data_Vencimento"] = df_venc_xp["Data_Vencimento"].dt.tz_localize(None)

                        df_venc_xp = df_venc_xp.dropna(subset=["Data_Vencimento"])

                        data_inicio = pd.Timestamp.now().normalize()
                        data_fim = pd.to_datetime("2030-12-31")
                        df_venc_xp = df_venc_xp[(df_venc_xp["Data_Vencimento"] >= data_inicio) & (df_venc_xp["Data_Vencimento"] <= data_fim)]

                        if not df_venc_xp.empty:
                            df_venc_xp["Ano_Mes"] = df_venc_xp["Data_Vencimento"].dt.to_period("M").astype(str)
                            venc_agrupado_xp = df_venc_xp.groupby("Ano_Mes")["Valor_Calculado"].sum() / 1e6
                            idx_completo_xp = pd.period_range(start=data_inicio.strftime("%Y-%m"), end="2030-12", freq="M").astype(str)
                            venc_agrupado_xp = venc_agrupado_xp.reindex(idx_completo_xp, fill_value=0)

                            fig_venc_xp, ax_venc_xp = plt.subplots(figsize=(14, 5))
                            fig_venc_xp.patch.set_facecolor("#0e1117")
                            ax_venc_xp.set_facecolor("#0e1117")

                            ax_venc_xp.bar(
                                venc_agrupado_xp.index,
                                venc_agrupado_xp.values,
                                color="#ff7f0e",
                                edgecolor="#ffffff",
                                linewidth=0.5,
                                width=0.7
                            )
                            ax_venc_xp.set_title("Valor Bruto dos Vencimentos por Mês - XP (R$ Milhões)", color="white", fontsize=12, pad=15)
                            ax_venc_xp.set_ylabel("Valor Bruto (R$ Milhões)", color="white", fontsize=10)
                            plt.xticks(rotation=90, ha="center", color="white", fontsize=8)
                            ax_venc_xp.tick_params(colors="white")
                            ax_venc_xp.xaxis.label.set_color("white")
                            ax_venc_xp.yaxis.label.set_color("white")
                            for spine in ax_venc_xp.spines.values():
                                spine.set_edgecolor("white")
                            ax_venc_xp.grid(True, linestyle="--", alpha=0.2, color="white")
                            plt.tight_layout()
                            st.pyplot(fig_venc_xp)

    elif corretora_selecionada == "Base":
        # =========================================================================
        # SEÇÃO BASE (BTG e XP em sequência, com colunas de contrato nas primeiras posições)
        # =========================================================================
        st.subheader("📋 Base Consolidada - Posições Brutas com Dados de Contrato")

        st.markdown(
            "Abaixo estão disponíveis as planilhas unificadas com as colunas do Controle de Contratos "
            "posicionadas nas primeiras colunas, seguindo a ordem solicitada (Primeiro BTG, Depois XP)."
        )

        cols_contrato_desejadas = ["Corretora", "Cliente", "Conta", "Status", "Situação", "Carteira", "Observações"]

        # --- SEÇÃO BTG NA BASE ---
        st.markdown("### 🔵 Posição Bruta BTG + Contratos")
        if not base_pos_btg.empty:
            df_base_btg_exibir = base_pos_btg.drop(columns=["Conta_Limpa"], errors="ignore")
            cols_atuais_btg = [c for c in df_base_btg_exibir.columns]
            cols_extras_btg = [c for c in cols_atuais_btg if c not in cols_contrato_desejadas]
            cols_ordenadas_btg = [c for c in cols_contrato_desejadas if c in cols_atuais_btg] + cols_extras_btg
            df_base_btg_exibir = df_base_btg_exibir[cols_ordenadas_btg]

            st.metric("Total de Registros (BTG)", len(df_base_btg_exibir))
            st.dataframe(df_base_btg_exibir, use_container_width=True)

            st.download_button(
                label="📥 Baixar Planilha Base BTG em Excel",
                data=gerar_excel_base(base_pos_btg),
                file_name="base_posicao_btg_com_contratos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum dado de posição BTG encontrado.")

        st.markdown("---")

        # --- SEÇÃO XP NA BASE ---
        st.markdown("### 🟡 Posição Bruta XP + Contratos")
        if not base_pos_xp.empty:
            df_base_xp_exibir = base_pos_xp.drop(columns=["Conta_Limpa"], errors="ignore")
            cols_atuais_xp = [c for c in df_base_xp_exibir.columns]
            cols_extras_xp = [c for c in cols_atuais_xp if c not in cols_contrato_desejadas]
            cols_ordenadas_xp = [c for c in cols_contrato_desejadas if c in cols_atuais_xp] + cols_extras_xp
            df_base_xp_exibir = df_base_xp_exibir[cols_ordenadas_xp]

            st.metric("Total de Registros (XP)", len(df_base_xp_exibir))
            st.dataframe(df_base_xp_exibir, use_container_width=True)

            st.download_button(
                label="📥 Baixar Planilha Base XP em Excel",
                data=gerar_excel_base(base_pos_xp),
                file_name="base_posicao_xp_com_contratos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Nenhum dado de posição XP encontrado.")

else:
    st.warning("Nenhum dado encontrado. Verifique se os arquivos Excel estão na pasta.")









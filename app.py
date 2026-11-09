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

def gerar_excel_base_multi_abas(path_pos, df_contratos_ref, cols_contrato):
    """Gera um Excel mantendo todas as abas originais de posição e adicionando as colunas de contrato nas primeiras posições."""
    output = io.BytesIO()
    if not os.path.exists(path_pos):
        return output.getvalue()

    try:
        xls = pd.ExcelFile(path_pos, engine="openpyxl")
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for aba_nome in xls.sheet_names:
                df_aba = pd.read_excel(xls, sheet_name=aba_nome)
                df_aba.columns = df_aba.columns.str.strip()

                col_conta = None
                for candidato in ["CodigoCliente", "Código do Cliente", "Codigo Cliente", "Conta"]:
                    if candidato in df_aba.columns:
                        col_conta = candidato
                        break

                if col_conta:
                    df_aba["Conta_Limpa"] = limpar_conta(df_aba[col_conta])
                    df_merged = df_aba.merge(
                        df_contratos_ref.drop(columns=["Conta"], errors="ignore"),
                        on="Conta_Limpa",
                        how="left"
                    )
                    df_merged = df_merged.drop(columns=["Conta_Limpa"], errors="ignore")

                    # Organiza as colunas de contrato nas primeiras posições
                    cols_atuais = [c for c in df_merged.columns]
                    cols_extras = [c for c in cols_atuais if c not in cols_contrato]
                    cols_ordenadas = [c for c in cols_contrato if c in cols_atuais] + cols_extras
                    df_final = df_merged[cols_ordenadas]
                else:
                    df_final = df_aba

                df_final.to_excel(writer, index=False, sheet_name=aba_nome)
    except Exception as e:
        pass

    output.seek(0)
    wb = openpyxl.load_workbook(output)
    font_header = Font(name="Calibri", size=11, bold=True)
    border_bottom = Border(bottom=Side(style="medium", color="000000"))

    for ws in wb.worksheets:
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
@st.cache_data(show_spinner="Carregando dados...")
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
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), None, None

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
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), None, None

    # Caminhos de Posição
    path_pos_btg = "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
    
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

    # -------------------------
    # POSIÇÃO BTG
    # -------------------------
    btg_pos_pivot = pd.DataFrame()
    posicao_global_btg = pd.DataFrame()

    if os.path.exists(path_pos_btg):
        try:
            posicao = pd.read_excel(path_pos_btg, sheet_name="Export", engine="openpyxl")
            posicao.columns = posicao.columns.str.strip()

            col_conta_btg = "CodigoCliente" if "CodigoCliente" in posicao.columns else "Conta"
            if col_conta_btg in posicao.columns:
                posicao["Conta_Limpa"] = limpar_conta(posicao[col_conta_btg])

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
            pass

    # -------------------------
    # POSIÇÃO XP
    # -------------------------
    xp_pos_pivot = pd.DataFrame()
    posicao_global_xp = pd.DataFrame()

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
        except Exception as e:
            pass

    return contratos_geral, btg_pos_pivot, posicao_global_btg, xp_pos_pivot, posicao_global_xp, path_pos_btg, path_pos_xp

# =========================
# MONTAGEM DA INTERFACE
# =========================

contratos_df, btg_pos_pivot, posicao_global_btg, xp_pos_pivot, posicao_global_xp, path_pos_btg, path_pos_xp = carregar_dados()

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

    elif corretora_selecionada == "XP":
        # =========================================================================
        # SEÇÃO XP (MANTIDA ORIGINAL SEM ALTERAÇÃO)
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

    elif corretora_selecionada == "Base":
        # =========================================================================
        # SEÇÃO BASE (APENAS BOTÕES DE DOWNLOAD - PRIMEIRO BTG, DEPOIS XP)
        # =========================================================================
        st.subheader("📋 Base Consolidada - Downloads dos Arquivos Brutos")

        st.markdown(
            "Utilize os botões abaixo para baixar as planilhas completas de posições (com todas as abas originais) "
            "contendo as colunas do Controle de Contratos inseridas nas primeiras colunas."
        )

        cols_contrato_desejadas = ["Corretora", "Cliente", "Conta", "Status", "Situação", "Carteira", "Observações"]

        # 1. BTG
        st.markdown("### 🔵 Posição Bruta BTG + Contratos")
        df_contratos_btg_ref = contratos_df[contratos_df["Corretora"] == "BTG"].copy()
        df_contratos_btg_ref["Conta_Limpa"] = limpar_conta(df_contratos_btg_ref["Conta"])

        excel_btg_base = gerar_excel_base_multi_abas(path_pos_btg, df_contratos_btg_ref, cols_contrato_desejadas)
        st.download_button(
            label="📥 Baixar Planilha Base BTG (Multi-abas com Contratos)",
            data=excel_btg_base,
            file_name="base_posicao_btg_completa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("---")

        # 2. XP
        st.markdown("### 🟡 Posição Bruta XP + Contratos")
        df_contratos_xp_ref = contratos_df[contratos_df["Corretora"] == "XP"].copy()
        df_contratos_xp_ref["Conta_Limpa"] = limpar_conta(df_contratos_xp_ref["Conta"])

        excel_xp_base = gerar_excel_base_multi_abas(path_pos_xp, df_contratos_xp_ref, cols_contrato_desejadas)
        st.download_button(
            label="📥 Baixar Planilha Base XP (Multi-abas com Contratos)",
            data=excel_xp_base,
            file_name="base_posicao_xp_completa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.warning("Nenhum dado encontrado. Verifique se os arquivos Excel estão na pasta.")









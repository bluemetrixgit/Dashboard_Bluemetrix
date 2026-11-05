import os
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

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

# =========================
# CARREGAR E PROCESSAR DADOS (COM CACHE INTELIGENTE)
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
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

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

        if "Conta" in contratos_geral.columns:
            contratos_geral["Conta_Limpa"] = limpar_conta(contratos_geral["Conta"])
    except Exception as e:
        st.error(f"Erro ao ler Controle de Contratos: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # -------------------------
    # POSIÇÃO BTG
    # -------------------------
    path_pos_btg = "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
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
                        df_aba["Aba_Origem"] = aba_nome

                        if col_alvo in df_aba.columns:
                            df_aba["Valor_Calculado"] = limpar_coluna_numerica(df_aba[col_alvo])
                        else:
                            cands = [c for c in df_aba.columns if col_alvo.lower() in c.lower()]
                            if cands:
                                df_aba["Valor_Calculado"] = limpar_coluna_numerica(df_aba[cands[0]])
                            else:
                                df_aba["Valor_Calculado"] = 0

                        # Mapear DataVencimento especificamente para Coe, Renda Fixa e Tesouro Direto
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

                        df_res_aba = df_aba.groupby("Conta_Limpa")["Valor_Calculado"].sum().reset_index()
                        df_res_aba.rename(columns={"Valor_Calculado": aba_nome}, inplace=True)
                        dfs_por_aba.append(df_res_aba)

            if dfs_por_aba:
                from functools import reduce
                xp_pos_pivot = reduce(lambda left, right: pd.merge(left, right, on="Conta_Limpa", how="outer"), dfs_por_aba).fillna(0)

            if dfs_globais:
                posicao_global_xp = pd.concat(dfs_globais, ignore_index=True)

        except Exception as e:
            st.warning(f"Aviso ao ler relatório de posições da XP: {e}")

    return contratos_geral, btg_pos_pivot, posicao_global_btg, xp_pos_pivot, posicao_global_xp

# =========================
# MONTAGEM DA INTERFACE
# =========================

contratos_df, btg_pos_pivot, posicao_global_btg, xp_pos_pivot, posicao_global_xp = carregar_dados()

if not contratos_df.empty:
    st.sidebar.title("Navegação")
    corretora_selecionada = st.sidebar.radio("Selecione a Corretora", ["BTG", "XP"])
    st.sidebar.markdown("---")

    if corretora_selecionada == "BTG":
        # =========================================================================
        # SEÇÃO BTG
        # =========================================================================
        st.subheader("🔵 Corretora BTG - Contratos & Posições")

        df_btg_view = contratos_df[contratos_df["Corretora"] == "BTG"].copy()

        colunas_posicao = []
        if not btg_pos_pivot.empty:
            df_btg_view = df_btg_view.merge(btg_pos_pivot, on="Conta_Limpa", how="left")
            colunas_posicao = [c for c in btg_pos_pivot.columns if c != "Conta_Limpa"]
            for col in colunas_posicao:
                df_btg_view[col] = df_btg_view[col].fillna(0)

        st.sidebar.header("Filtros BTG")

        contas_selecionadas = []
        carteiras_selecionadas = []
        status_selecionados = []
        produtos_selecionados = []

        if "Conta" in df_btg_view.columns:
            contas_disponiveis = sorted([str(c) for c in df_btg_view["Conta"].dropna().unique()])
            contas_selecionadas = st.sidebar.multiselect("Filtrar por Conta", contas_disponiveis, key="conta_btg_sel")
            if contas_selecionadas:
                df_btg_view = df_btg_view[df_btg_view["Conta"].astype(str).isin(contas_selecionadas)]

        if "Carteira" in df_btg_view.columns:
            carteiras_disponiveis = sorted([str(c) for c in df_btg_view["Carteira"].dropna().unique()])
            carteiras_selecionadas = st.sidebar.multiselect("Filtrar por Carteira", carteiras_disponiveis, key="cart_btg_sel")
            if carteiras_selecionadas:
                df_btg_view = df_btg_view[df_btg_view["Carteira"].astype(str).isin(carteiras_selecionadas)]

        if "Status" in df_btg_view.columns:
            status_disponiveis = sorted([str(s) for s in df_btg_view["Status"].dropna().unique()])
            status_selecionados = st.sidebar.multiselect("Filtrar por Status", status_disponiveis, key="status_btg_sel")
            if status_selecionados:
                df_btg_view = df_btg_view[df_btg_view["Status"].astype(str).isin(status_selecionados)]

        st.sidebar.markdown("---")
        st.sidebar.header("Filtro Vencimentos")
        produtos_disponiveis = []
        if not posicao_global_btg.empty and "Produto" in posicao_global_btg.columns:
            produtos_disponiveis = sorted([str(p) for p in posicao_global_btg["Produto"].dropna().unique()])
        
        produtos_selecionados = st.sidebar.multiselect("Filtrar por Produto", produtos_disponiveis, key="venc_prod_sel_btg")

        st.metric("Total de Contratos (BTG)", len(df_btg_view))

        df_exibir_btg = df_btg_view.copy()
        if not btg_pos_pivot.empty:
            for col in colunas_posicao:
                if col in df_exibir_btg.columns:
                    df_exibir_btg[col] = df_exibir_btg[col].apply(formatar_real)

        if "Conta_Limpa" in df_exibir_btg.columns:
            df_exibir_btg = df_exibir_btg.drop(columns=["Conta_Limpa"])

        st.dataframe(df_exibir_btg, use_container_width=True)

        if colunas_posicao:
            st.markdown("### 📊 Consolidado de Valor Bruto por Mercado/Produto (BTG)")
            soma_posicoes = df_btg_view[colunas_posicao].sum() / 1e6

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

                    ax_pie.legend(
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

                    ax_pie.set_title("Participação Percentual por Categoria", color="white", fontsize=11, pad=12)
                    plt.tight_layout()
                    st.pyplot(fig_pie)

                st.markdown("### 📅 Vencimentos Mensais - BTG (Até Dezembro de 2030)")
                if not posicao_global_btg.empty:
                    df_venc = posicao_global_btg.copy()
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

    else:
        # =========================================================================
        # SEÇÃO XP
        # =========================================================================
        st.subheader("🟡 Corretora XP - Contratos & Posições")

        df_xp_view = contratos_df[contratos_df["Corretora"] == "XP"].copy()

        colunas_posicao_xp = []
        if not xp_pos_pivot.empty:
            df_xp_view = df_xp_view.merge(xp_pos_pivot, on="Conta_Limpa", how="left")
            colunas_posicao_xp = [c for c in xp_pos_pivot.columns if c != "Conta_Limpa"]
            for col in colunas_posicao_xp:
                df_xp_view[col] = df_xp_view[col].fillna(0)

        st.sidebar.header("Filtros XP")

        contas_xp_sel = []
        carteiras_xp_sel = []
        status_xp_sel = []
        produtos_xp_sel = []

        if "Conta" in df_xp_view.columns:
            contas_xp_disp = sorted([str(c) for c in df_xp_view["Conta"].dropna().unique()])
            contas_xp_sel = st.sidebar.multiselect("Filtrar por Conta", contas_xp_disp, key="conta_xp_sel")
            if contas_xp_sel:
                df_xp_view = df_xp_view[df_xp_view["Conta"].astype(str).isin(contas_xp_sel)]

        if "Carteira" in df_xp_view.columns:
            carteiras_xp_disp = sorted([str(c) for c in df_xp_view["Carteira"].dropna().unique()])
            carteiras_xp_sel = st.sidebar.multiselect("Filtrar por Carteira", carteiras_xp_disp, key="cart_xp_sel")
            if carteiras_xp_sel:
                df_xp_view = df_xp_view[df_xp_view["Carteira"].astype(str).isin(carteiras_xp_sel)]

        if "Status" in df_xp_view.columns:
            status_xp_disp = sorted([str(s) for s in df_xp_view["Status"].dropna().unique()])
            status_xp_sel = st.sidebar.multiselect("Filtrar por Status", status_xp_disp, key="status_xp_sel")
            if status_xp_sel:
                df_xp_view = df_xp_view[df_xp_view["Status"].astype(str).isin(status_xp_sel)]

        st.sidebar.markdown("---")
        st.sidebar.header("Filtro Vencimentos")
        
        # Extrair opções estritamente da coluna "Categoria" da aba "Renda Fixa" (com fallbacks de segurança)
        produtos_xp_disp = []
        if not posicao_global_xp.empty:
            df_rf_apenas = posicao_global_xp[posicao_global_xp["Aba_Origem"] == "Renda Fixa"]
            if "Categoria" in df_rf_apenas.columns:
                produtos_xp_disp = sorted([str(p) for p in df_rf_apenas["Categoria"].dropna().unique()])
            
            # Se por acaso a aba Renda Fixa não tiver a coluna Categoria, busca nas outras abas permitidas
            if not produtos_xp_disp:
                df_venc_opcoes = posicao_global_xp[posicao_global_xp["Aba_Origem"].isin(["Renda Fixa", "Coe", "Tesouro Direto"])]
                for col_prod in ["Categoria", "Produto", "Ativo", "SubMercado"]:
                    if col_prod in df_venc_opcoes.columns:
                        produtos_xp_disp = sorted([str(p) for p in df_venc_opcoes[col_prod].dropna().unique()])
                        if produtos_xp_disp:
                            break
        
        produtos_xp_sel = st.sidebar.multiselect("Filtrar por Produto/Ativo", produtos_xp_disp, key="venc_prod_sel_xp")

        st.metric("Total de Contratos (XP)", len(df_xp_view))

        df_exibir_xp = df_xp_view.copy()
        if not xp_pos_pivot.empty:
            for col in colunas_posicao_xp:
                if col in df_exibir_xp.columns:
                    df_exibir_xp[col] = df_exibir_xp[col].apply(formatar_real)

        if "Conta_Limpa" in df_exibir_xp.columns:
            df_exibir_xp = df_exibir_xp.drop(columns=["Conta_Limpa"])

        st.dataframe(df_exibir_xp, use_container_width=True)

        if colunas_posicao_xp:
            st.markdown("### 📊 Consolidado por Abas/Produtos (XP)")
            soma_pos_xp = df_xp_view[colunas_posicao_xp].sum() / 1e6

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

                    cores_neutras_xp = [
                        "#ff7f0e", "#ffbb78", "#aec7e8", "#1f77b4", 
                        "#ff9896", "#98df8a", "#c5b0d5", "#c49c94", "#f7b6d2"
                    ]

                    wedges_xp, texts_xp, autotexts_xp = ax_pie_xp.pie(
                        soma_pizza_xp.values,
                        autopct='%1.1f%%',
                        startangle=140,
                        colors=cores_neutras_xp[:len(soma_pizza_xp)],
                        textprops=dict(color="white", fontsize=8),
                        pctdistance=0.75
                    )

                    for autotext in autotexts_xp:
                        autotext.set_color('white')
                        autotext.set_fontsize(7.5)

                    ax_pie_xp.legend(
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

                    ax_pie_xp.set_title("Participação Percentual por Categoria (XP)", color="white", fontsize=11, pad=12)
                    plt.tight_layout()
                    st.pyplot(fig_pie_xp)

                st.markdown("### 📅 Vencimentos Mensais - XP (Até Dezembro de 2030)")
                if not posicao_global_xp.empty:
                    df_venc_xp = posicao_global_xp[posicao_global_xp["Aba_Origem"].isin(["Renda Fixa", "Coe", "Tesouro Direto"])].copy()
                    
                    if produtos_xp_sel:
                        # Filtra utilizando a coluna Categoria prioritariamente ou demais colunas de identificação
                        col_filtragem = "Categoria" if "Categoria" in df_venc_xp.columns else next((c for c in ["Produto", "Ativo", "SubMercado"] if c in df_venc_xp.columns), None)
                        if col_filtragem:
                            df_venc_xp = df_venc_xp[df_venc_xp[col_filtragem].astype(str).isin(produtos_xp_sel)]

                    if "Data_Vencimento_Calculada" in df_venc_xp.columns:
                        df_venc_xp["Data_Vencimento"] = df_venc_xp["Data_Vencimento_Calculada"]
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

else:
    st.warning("Nenhum dado encontrado. Verifique se os arquivos Excel estão na pasta.")









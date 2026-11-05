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

# =========================
# CARREGAR E PROCESSAR DADOS
# =========================

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
        return pd.DataFrame(), pd.DataFrame()

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
        return pd.DataFrame(), pd.DataFrame()

    path_pos_btg = "data/Posição.xlsx" if os.path.exists("data/Posição.xlsx") else "Posição.xlsx"
    btg_pos_pivot = pd.DataFrame()

    if os.path.exists(path_pos_btg):
        try:
            posicao = pd.read_excel(path_pos_btg, sheet_name="Export", engine="openpyxl")
            posicao.columns = posicao.columns.str.strip()

            if "Conta" in posicao.columns:
                posicao["Conta_Limpa"] = limpar_conta(posicao["Conta"])

                def definir_coluna(row):
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

                posicao["Coluna_Mercado"] = posicao.apply(definir_coluna, axis=1)

                col_vb = "Valor Bruto" if "Valor Bruto" in posicao.columns else [c for c in posicao.columns if "valor" in c.lower()][0]
                if posicao[col_vb].dtype == object:
                    posicao["Valor_Bruto_Num"] = pd.to_numeric(
                        posicao[col_vb].astype(str)
                        .str.replace("R$", "", regex=False)
                        .str.replace(".", "", regex=False)
                        .str.replace(",", ".", regex=False)
                        .str.strip(),
                        errors="coerce"
                    ).fillna(0)
                else:
                    posicao["Valor_Bruto_Num"] = pd.to_numeric(posicao[col_vb], errors="coerce").fillna(0)

                btg_pos_pivot = posicao.pivot_table(
                    index="Conta_Limpa",
                    columns="Coluna_Mercado",
                    values="Valor_Bruto_Num",
                    aggfunc="sum",
                    fill_value=0
                ).reset_index()

                # Removendo colunas indesejadas
                colunas_para_remover = [c for c in btg_pos_pivot.columns if pd.isna(c) or str(c).lower() in ["nan", "none", "valor em trânsito", "valor em transito"]]
                btg_pos_pivot = btg_pos_pivot.drop(columns=colunas_para_remover, errors="ignore")

                # =========================
                # AGRUPAR COLUNAS EM "Outros"
                # =========================
                colunas_para_agrupar = [
                    "Conta Corrente", "COE", "Criptoativos", 
                    "Derivativos", "Fundos Cetipados", "OPCAO", "Tesouro Direto"
                ]
                # Encontrar quais dessas colunas realmente existem no pivot
                cols_existentes = [c for c in colunas_para_agrupar if c in btg_pos_pivot.columns]
                
                if cols_existentes:
                    # Somar as colunas e criar a coluna "Outros"
                    btg_pos_pivot["Outros"] = btg_pos_pivot[cols_existentes].sum(axis=1)
                    # Remover as colunas individuais originais
                    btg_pos_pivot = btg_pos_pivot.drop(columns=cols_existentes)

                    # Garantir que "Outros" fique no final (junto com Conta_Limpa)
                    conta_limpa_col = btg_pos_pivot.pop("Conta_Limpa")
                    outros_col = btg_pos_pivot.pop("Outros")
                    btg_pos_pivot["Conta_Limpa"] = conta_limpa_col
                    btg_pos_pivot["Outros"] = outros_col

        except Exception as e:
            st.warning(f"Aviso ao ler Posição BTG: {e}")

    return contratos_geral, btg_pos_pivot

# =========================
# MONTAGEM DA INTERFACE
# =========================

contratos_df, btg_pos_pivot = carregar_dados()

if not contratos_df.empty:
    aba_btg, aba_xp = st.tabs(["BTG", "XP"])

    # -------------------------
    # ABA BTG
    # -------------------------
    with aba_btg:
        st.subheader("🔵 Corretora BTG - Contratos & Posições")

        df_btg_view = contratos_df[contratos_df["Corretora"] == "BTG"].copy()

        colunas_posicao = []
        if not btg_pos_pivot.empty:
            df_btg_view = df_btg_view.merge(btg_pos_pivot, on="Conta_Limpa", how="left")
            colunas_posicao = [c for c in btg_pos_pivot.columns if c != "Conta_Limpa"]
            for col in colunas_posicao:
                df_btg_view[col] = df_btg_view[col].fillna(0)

        # -------------------------
        # FILTROS LATERAIS
        # -------------------------
        st.sidebar.header("Filtros BTG")

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

        st.metric("Total de Contratos", len(df_btg_view))

        # Formatação Visual da Tabela
        df_exibir_btg = df_btg_view.copy()

        if not btg_pos_pivot.empty:
            for col in colunas_posicao:
                if col in df_exibir_btg.columns:
                    df_exibir_btg[col] = df_exibir_btg[col].apply(formatar_real)

        if "Conta_Limpa" in df_exibir_btg.columns:
            df_exibir_btg = df_exibir_btg.drop(columns=["Conta_Limpa"])

        st.dataframe(df_exibir_btg, use_container_width=True)

        # -------------------------
        # GRÁFICO DE BARRAS
        # -------------------------
        if colunas_posicao:
            st.markdown("### 📊 Consolidado de Valor Bruto por Mercado/Produto")

            soma_posicoes = df_btg_view[colunas_posicao].sum() / 1e6  # Convertido para Milhões

            if not soma_posicoes.empty and soma_posicoes.sum() > 0:
                fig, ax = plt.subplots(figsize=(12, 5))

                fig.patch.set_facecolor("#0e1117")
                ax.set_facecolor("#0e1117")

                bars = ax.bar(
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

                # Legenda explicativa para a coluna "Outros"
                st.markdown(
                    "> **ℹ️ Legenda da categoria 'Outros':** A coluna *Outros* consolida os valores de: "
                    "*Conta Corrente, COE, Criptoativos, Derivativos, Fundos Cetipados, OPCAO e Tesouro Direto*."
                )

                # -------------------------
                # GRÁFICO DE PIZZA (Compacto e com Legenda Lateral Organizada)
                # -------------------------
                st.markdown("### 🥧 Distribuição Percentual por Categoria")
                
                soma_pizza = soma_posicoes[soma_posicoes > 0]

                if not soma_pizza.empty:
                    # Tamanho compacto
                    fig_pie, ax_pie = plt.subplots(figsize=(7, 4))
                    fig_pie.patch.set_facecolor("#0e1117")
                    ax_pie.set_facecolor("#0e1117")

                    cores_neutras = [
                        "#1f77b4", "#4682b4", "#5f9ea0", "#708090", 
                        "#5c6b73", "#3a506b", "#4a5568", "#2d3748", "#4a607a"
                    ]

                    # Geramos a pizza sem os rótulos diretos nas fatias para evitar poluição
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

                    # Adicionando uma legenda lateral organizada e separada
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
    # -------------------------
    # ABA XP
    # -------------------------
    with aba_xp:
        st.subheader("🟡 Corretora XP - Contratos")

        df_xp_view = contratos_df[contratos_df["Corretora"] == "XP"].copy()

        if "Carteira" in df_xp_view.columns:
            carteiras_xp = ["Todas"] + sorted([str(c) for c in df_xp_view["Carteira"].dropna().unique()])
            sel_cart_xp = st.selectbox("Filtrar por Carteira", carteiras_xp, key="cart_xp_sel")
            if sel_cart_xp != "Todas":
                df_xp_view = df_xp_view[df_xp_view["Carteira"].astype(str) == sel_cart_xp]

        if "Conta_Limpa" in df_xp_view.columns:
            df_xp_view = df_xp_view.drop(columns=["Conta_Limpa"])

        st.metric("Total de Contratos", len(df_xp_view))
        st.dataframe(df_xp_view, use_container_width=True)

else:
    st.warning("Nenhum dado encontrado. Verifique se os arquivos Excel estão na pasta.")

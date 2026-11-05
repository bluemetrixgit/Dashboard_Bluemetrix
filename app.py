st.markdown("### 📅 Vencimentos Mensais - XP (Até Dezembro de 2030)")
                if not posicao_global_xp.empty:
                    df_venc_xp = posicao_global_xp[posicao_global_xp["Aba_Origem"].isin(["Renda Fixa", "Coe", "Tesouro Direto"])].copy()
                    
                    if produtos_xp_sel:
                        col_filtragem = "Categoria" if "Categoria" in df_venc_xp.columns else next((c for c in ["Produto", "Ativo", "SubMercado"] if c in df_venc_xp.columns), None)
                        if col_filtragem:
                            df_venc_xp = df_venc_xp[df_venc_xp[col_filtragem].astype(str).isin(produtos_xp_sel)]

                    if "Data_Vencimento_Calculada" in df_venc_xp.columns:
                        # Converte para datetime e remove qualquer fuso horário (tz-aware -> tz-naive) para evitar conflitos
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









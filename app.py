import streamlit as st
import pandas as pd
import plotly.express as px
import io
import unicodedata


# 1. Configuração de página
st.set_page_config(page_title="Pendências de Faturamento | Bradesco — Solar Cuidados", page_icon="🏦", layout="wide")


# 2. Injeção da Paleta de Cores Exata (Bordô, Dourado e Off-White) — mesmo padrão do dashboard de Prorrogações
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght=300;400;500;600;700;800&display=swap');
    
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #FAFAF9 !important; 
        color: #1A1714 !important; 
        font-family: 'Inter', sans-serif !important;
    }
    
    .topbar-header {
        background-color: #3D0B16 !important; 
        padding: 12px 20px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 14px;
        margin-bottom: 5px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .brand-mark {
        width: 38px;
        height: 38px;
        background: linear-gradient(135deg, #C07C20, #E09A30); 
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 18px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.2);
    }
    .brand-title {
        color: #FFFFFF !important;
        font-size: 22px !important;
        font-weight: 800 !important;
        margin: 0 !important;
        letter-spacing: -.2px;
    }
    .brand-title span {
        color: #E09A30 !important;
    }
    .subtitle {
        font-size: 13px;
        color: #8A7D72 !important;
        margin-top: 5px;
        margin-bottom: 25px;
        font-weight: 500;
    }

    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #F3F0EC !important; 
        border: 1px solid #E2DDD6 !important; 
        color: #52473E !important; 
        border-radius: 8px 8px 0px 0px !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        font-size: 11px !important;
        letter-spacing: .2px;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background-color: #5C1220 !important; 
        color: #FFFFFF !important;
        border-color: #5C1220 !important;
    }

    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2DDD6 !important;
        border-radius: 12px !important; 
        padding: 14px 18px !important;
        box-shadow: 0 2px 8px rgba(28,14,10,.08) !important; 
        border-left: 4px solid #C07C20 !important; 
    }
    div[data-testid="stMetricLabel"] {
        color: #5C1220 !important; 
        font-weight: 700 !important;
        font-size: 12px !important;
        text-transform: uppercase !important;
        letter-spacing: .5px;
    }
    div[data-testid="stMetricValue"] {
        color: #1A1714 !important;
        font-weight: 800 !important;
        font-size: 24px !important;
        letter-spacing: -1px;
    }
    
    .stDataFrame {
        background-color: #FFFFFF !important;
        border: 1px solid #E2DDD6 !important;
        border-radius: 12px !important;
    }

    /* Box de Insights */
    .insight-card {
        background-color: #FFFFFF;
        border-left: 5px solid #5C1220;
        padding: 15px 20px;
        border-radius: 4px 12px 12px 4px;
        margin-bottom: 12px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.03);
    }

    .badge-origem {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 700;
    }
    .badge-doc { background-color: #FDEBD0; color: #8A5A00; }
    .badge-semdoc { background-color: #F8D7DA; color: #721C24; }
    </style>
""", unsafe_allow_html=True)


# Cabeçalho Oficial
st.markdown("""
    <div class="topbar-header">
        <div class="brand-mark">🏦</div>
        <h1 class="brand-title">Bradesco — <span>Pendências de Faturamento</span></h1>
    </div>
""", unsafe_allow_html=True)
st.markdown('<p class="subtitle">Cruzamento entre documentos de faturamento e base de orçamento/pendências (IW) para identificar atendimentos que ainda não passaram pelo processo de faturamento.</p>', unsafe_allow_html=True)


# --- ÁREA DE UPLOAD ---
col_up1, col_up2 = st.columns(2)
with col_up1:
    arquivo_faturamento = st.file_uploader("1 - FATURAMENTO (.csv/.xlsx)", type=["csv", "xlsx"])
with col_up2:
    arquivo_bradesco = st.file_uploader("2 - BRADESCO — Base de Orçamento/Pendências (.csv/.xlsx)", type=["csv", "xlsx"])


def carregar_arquivo(arquivo):
    if arquivo.name.endswith('.csv'):
        try:
            arquivo.seek(0)
            return pd.read_csv(arquivo, sep=';', encoding='utf-8')
        except Exception:
            arquivo.seek(0)
            return pd.read_csv(arquivo, sep=';', encoding='iso-8859-1')
    return pd.read_excel(arquivo)


def normalizar_texto(texto):
    """Remove acentos e caixa para comparações robustas de texto."""
    texto = str(texto) if not pd.isna(texto) else ''
    texto = texto.strip().lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    return texto


def parse_valor_brl(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    v = str(valor).strip().replace('R$', '').strip()
    if not v:
        return None
    v = v.replace('.', '').replace(',', '.')
    try:
        return float(v)
    except ValueError:
        return None


def parse_data(valor):
    if pd.isna(valor):
        return None
    return pd.to_datetime(valor, dayfirst=True, errors='coerce')


def formatar_para_exibicao(df):
    """Pré-formata datas e valores monetários como texto simples (sem usar pandas Styler),
    o que evita o bug de renderização 'removeChild' do Streamlit em alguns navegadores/versões."""
    df_fmt = df.copy()
    for col in ['Data Início', 'Data Fim']:
        if col in df_fmt.columns:
            df_fmt[col] = df_fmt[col].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notna(d) else '')
    for col in ['Valor Orçado', 'Valor Faturado']:
        if col in df_fmt.columns:
            df_fmt[col] = df_fmt[col].apply(lambda v: f"R$ {v:,.2f}" if pd.notna(v) else '')
    return df_fmt


if arquivo_faturamento and arquivo_bradesco:
    try:
        fat = carregar_arquivo(arquivo_faturamento)
        bra = carregar_arquivo(arquivo_bradesco)
        fat.columns = fat.columns.str.strip()
        bra.columns = bra.columns.str.strip()

        # --- PARTE A: documentos gerados no faturamento, mas nunca processados ---
        parte_a = fat[fat['Status Doc'].isna() & fat['Valor Faturado (Conta)'].isna()].copy()
        parte_a_view = pd.DataFrame({
            'ID Paciente': parte_a['ID de Atendimento'],
            'Nome Paciente': parte_a['Nome'],
            'Data Início': parte_a['Data Início'].apply(parse_data),
            'Data Fim': parte_a['Data Fim'].apply(parse_data),
            'Valor Orçado': parte_a['Valor Orçado'].apply(parse_valor_brl),
            'Valor Faturado': parte_a['Valor Faturado (Conta)'].apply(parse_valor_brl),
            'Comentários do Faturamento': parte_a['Comentários Faturamento'],
            'Justificativa Pendência': parte_a['Justificativa Pendência'],
            'Origem': 'Documento gerado, não processado'
        })

        # --- PARTE B: atendimentos do bradesco que nunca geraram documento de faturamento ---
        atend_fat_ids = set(fat['ID de Atendimento'].dropna().astype(int))
        bra_atend_num = pd.to_numeric(bra['Nº Atend.'], errors='coerce')
        mask_nao_gerou_doc = bra_atend_num.notna() & (~bra_atend_num.astype('Int64').isin(atend_fat_ids))
        mask_nao_ja_faturado = bra['Justificativa Pendência'].apply(normalizar_texto) != 'faturado'

        parte_b = bra[mask_nao_gerou_doc & mask_nao_ja_faturado].copy()
        parte_b_view = pd.DataFrame({
            'ID Paciente': bra_atend_num[mask_nao_gerou_doc & mask_nao_ja_faturado],
            'Nome Paciente': parte_b['Nome Paciente'],
            'Data Início': parte_b['Dt Inicio'].apply(parse_data),
            'Data Fim': parte_b['Dt Fim'].apply(parse_data),
            'Valor Orçado': parte_b['Vr. Cobrar'].apply(parse_valor_brl),
            'Valor Faturado': None,
            'Comentários do Faturamento': parte_b['Obs Faturamento'],
            'Justificativa Pendência': parte_b['Justificativa Pendência'],
            'Origem': 'Nunca gerou documento de faturamento'
        })

        df_final = pd.concat([parte_a_view, parte_b_view], ignore_index=True)
        df_final['ID Paciente'] = df_final['ID Paciente'].astype('Int64')
        df_final = df_final.sort_values(by='Valor Orçado', ascending=False, na_position='last').reset_index(drop=True)
        df_final['Justificativa Pendência'] = df_final['Justificativa Pendência'].fillna('(Não informado)')
        df_final['Mês'] = df_final['Data Início'].dt.to_period('M').astype(str)

        # Métricas globais
        total_pendencias = len(df_final)
        valor_total_orcado = df_final['Valor Orçado'].sum(skipna=True)
        qtd_parte_a = len(parte_a_view)
        qtd_parte_b = len(parte_b_view)
        setor_top = df_final['Justificativa Pendência'].value_counts().idxmax() if total_pendencias > 0 else '-'

        # --- ABAS DO DASHBOARD ---
        aba1, aba2, aba3 = st.tabs([
            "📌 Visão Geral", "📋 Lista Detalhada", "🔍 Por Justificativa (Setor)"
        ])

        with aba1:
            st.markdown("### 📌 Indicadores Gerais")
            card1, card2, card3, card4 = st.columns(4)
            card1.metric("Total de Pendências", f"{total_pendencias}")
            card2.metric("Valor Total Orçado", f"R$ {valor_total_orcado:,.2f}")
            card3.metric("Doc. Gerado, Não Processado", f"{qtd_parte_a}")
            card4.metric("Nunca Gerou Documento", f"{qtd_parte_b}")

            st.markdown("---")
            st.markdown("### 📊 Gráficos")
            col_g1, col_g2 = st.columns(2)

            with col_g1:
                st.markdown("#### Pendências por Justificativa (Setor)")
                df_setor = df_final.groupby('Justificativa Pendência').agg(
                    Valor=('Valor Orçado', 'sum'), Quantidade=('ID Paciente', 'count')
                ).reset_index().sort_values(by='Valor', ascending=True)
                fig_setor = px.bar(
                    df_setor, y='Justificativa Pendência', x='Valor', orientation='h',
                    color_discrete_sequence=['#C07C20'], custom_data=['Quantidade']
                )
                fig_setor.update_traces(hovertemplate='<b>%{y}</b><br>Valor: R$ %{x:,.2f}<br>Qtd: %{customdata[0]}<extra></extra>')
                fig_setor.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis=dict(tickprefix="R$ ", tickformat=",.0f"))
                st.plotly_chart(fig_setor, use_container_width=True)

            with col_g2:
                st.markdown("#### Pendências por Mês (Data Início)")
                df_mes = df_final.groupby('Mês').agg(
                    Valor=('Valor Orçado', 'sum'), Quantidade=('ID Paciente', 'count')
                ).reset_index().sort_values(by='Mês')
                fig_mes = px.bar(
                    df_mes, x='Mês', y='Quantidade', color_discrete_sequence=['#5C1220'],
                    custom_data=['Valor']
                )
                fig_mes.update_traces(hovertemplate='<b>%{x}</b><br>Qtd: %{y}<br>Valor: R$ %{customdata[0]:,.2f}<extra></extra>')
                fig_mes.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_mes, use_container_width=True)

            st.markdown("---")
            st.markdown("### 💡 Insight Automático")
            st.markdown(f"""<div class="insight-card"><b>🎯 Principal Justificativa Pendente:</b><br>
                        <b>{setor_top}</b> concentra a maior parte das pendências de faturamento.</div>""", unsafe_allow_html=True)

        with aba2:
            st.markdown("### 📋 Lista Detalhada de Pendências")

            col_f1, col_f2 = st.columns(2)
            with col_f1:
                busca = st.text_input("🔎 Buscar por nome, ID, justificativa ou comentário:")
            with col_f2:
                filtro_origem = st.selectbox("Filtrar por origem:", ["Todas", "Documento gerado, não processado", "Nunca gerou documento de faturamento"])

            df_view = df_final.copy()
            if busca:
                termo = normalizar_texto(busca)
                df_view = df_view[
                    df_view['Nome Paciente'].apply(normalizar_texto).str.contains(termo, na=False) |
                    df_view['ID Paciente'].astype(str).str.contains(termo, na=False) |
                    df_view['Justificativa Pendência'].apply(normalizar_texto).str.contains(termo, na=False) |
                    df_view['Comentários do Faturamento'].apply(normalizar_texto).str.contains(termo, na=False)
                ]
            if filtro_origem != "Todas":
                df_view = df_view[df_view['Origem'] == filtro_origem]

            st.markdown(f"**Mostrando {len(df_view)} de {len(df_final)} registros | Valor: R$ {df_view['Valor Orçado'].sum():,.2f}**")

            df_view_export = df_view[['ID Paciente', 'Nome Paciente', 'Data Início', 'Data Fim', 'Valor Orçado', 'Valor Faturado', 'Comentários do Faturamento', 'Justificativa Pendência', 'Origem']].copy()

            buffer_export = io.BytesIO()
            with pd.ExcelWriter(buffer_export, engine='xlsxwriter') as writer:
                df_view_export.to_excel(writer, sheet_name='Não Faturado', index=False)
            st.download_button(
                label="📥 Baixar Planilha Estruturada: Não Faturado",
                data=buffer_export.getvalue(),
                file_name="bradesco_nao_faturado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.dataframe(
                formatar_para_exibicao(df_view_export),
                use_container_width=True, hide_index=True
            )

        with aba3:
            st.markdown("### 🔍 Detalhamento por Justificativa (Setor)")
            setores_disponiveis = df_final['Justificativa Pendência'].value_counts().index.tolist()
            setor_selecionado = st.selectbox("Selecione a Justificativa Pendência:", setores_disponiveis)

            df_setor_detalhe = df_final[df_final['Justificativa Pendência'] == setor_selecionado].copy()
            st.markdown(f"**Total: {len(df_setor_detalhe)} pendências | Valor: R$ {df_setor_detalhe['Valor Orçado'].sum():,.2f}**")

            df_setor_export = df_setor_detalhe[['ID Paciente', 'Nome Paciente', 'Data Início', 'Data Fim', 'Valor Orçado', 'Valor Faturado', 'Comentários do Faturamento', 'Origem']].copy()
            st.dataframe(
                formatar_para_exibicao(df_setor_export),
                use_container_width=True, hide_index=True
            )

    except Exception as e:
        st.error(f"Erro ao processar os arquivos. Detalhe técnico: {e}")
else:
    st.info("💡 Selecione os dois arquivos acima (Faturamento e Bradesco) para carregar o cruzamento.")

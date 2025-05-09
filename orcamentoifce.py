import streamlit as st
import pandas as pd
import io
import plotly.express as px

# ============================
# 1️⃣ CARREGAMENTO DA PLANILHA
# ============================

# Caminho do arquivo
file1 = 'planilhaifce2025.xlsx'

# Função para carregar a planilha

def carregar_planilha(caminho):
    try:
        return pd.read_excel(caminho)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo {caminho}: {e}")
        return None

# Carregar os dados
planilha1_df = carregar_planilha(file1)

# ============================
# 2️⃣ EXIBIÇÃO COMPLETA DA PLANILHA (FORMATADA EM R$)
# ============================

if planilha1_df is not None:
    st.title("DEPARTAMENTO DE ADMINSTRAÇÃO E PLANEJAMENTO (DAP)\n ")
    st.subheader("ORÇAMENTO 'CAMPUS' TAUÁ-CE 2025 \n AÇÕES: 20RL (CUSTEIO) e 2994 (ASSISTẼNCIA)")

    # Conversão para moeda brasileira
    colunas_valores = ['VALOR PAGAMENTO MÊS', 'VALOR PAGAMENTO ANO', 'PAGAMENTO REALIZADO', 'SALDO DE EMPENHO']
    for coluna in colunas_valores:
        # Verifica se é numérico antes de tentar formatar
        planilha1_df[coluna] = pd.to_numeric(planilha1_df[coluna], errors='coerce')
        planilha1_df[coluna] = planilha1_df[coluna].apply(lambda x: f"R$ {x:,.2f}" if pd.notnull(x) else "-")

    # Calcular totais corretamente
    totais = pd.DataFrame({
        'EMPRESAS CREDORAS': ['Total Geral'],
        'VALOR PAGAMENTO MÊS': [pd.to_numeric(planilha1_df['VALOR PAGAMENTO MÊS'].str.replace('R\$', '').str.replace(',', ''), errors='coerce').sum()],
        'VALOR PAGAMENTO ANO': [pd.to_numeric(planilha1_df['VALOR PAGAMENTO ANO'].str.replace('R\$', '').str.replace(',', ''), errors='coerce').sum()],
        'PAGAMENTO REALIZADO': [pd.to_numeric(planilha1_df['PAGAMENTO REALIZADO'].str.replace('R\$', '').str.replace(',', ''), errors='coerce').sum()],
        'SALDO DE EMPENHO': [pd.to_numeric(planilha1_df['SALDO DE EMPENHO'].str.replace('R\$', '').str.replace(',', ''), errors='coerce').sum()]
    })

    # Adicionar a linha de total ao final da planilha
    planilha1_df = pd.concat([planilha1_df, totais], ignore_index=True)

    # ============================
    # 🎨 ESTILIZAÇÃO DAS LINHAS
    # ============================

    def highlight_rows(row):
        if 25 <= row.name <= 35:
            return ['color: #BDB76B'] * len(row)  
        elif 0 <= row.name <= 24:
            return ['color: #006400'] * len(row)  
        elif row.name == 36:
            return ['color: white'] * len(row) 
        else:
            return ['color: black'] * len(row)

    # Aplicar estilo e exibir no Streamlit
    st.dataframe(planilha1_df.style.apply(highlight_rows, axis=1), height=250)

    # ============================
    # 3️⃣ BOTÃO PARA EXPORTAR EM EXCEL
    # ============================

    try:
        import xlsxwriter
    except ImportError:
        st.error("Pacote 'xlsxwriter' não está instalado. Rode: pip install xlsxwriter")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        planilha1_df.to_excel(writer, index=False, sheet_name='Orcamento 2025')
    buffer.seek(0)

    st.download_button(
        label="📥 Baixar Planilha Completa em Excel",
        data=buffer,
        file_name='Orcamento_Publico_2025.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# Carregar os dados
file_path = 'Planilhasegunda.xlsx'
df = pd.read_excel(file_path, sheet_name='Página3')

# Remover linhas com valores NaN
df = df.dropna()

# Guardar os valores numéricos para o gráfico
df_numerico = df.copy()

# Função para formatar em Real Brasileiro (R$)
def formatar_moeda(valor):
    return f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

# Aplicar a formatação em Real para visualização e manter os dados numéricos para o gráfico
for coluna in df.columns[1:]:
    df_numerico[coluna] = pd.to_numeric(df_numerico[coluna], errors='coerce')
    df[coluna] = df_numerico[coluna].apply(lambda x: formatar_moeda(x) if pd.notnull(x) else x)

# Adicionar rótulos apenas com valor em Real (R$)
df_numerico['A RECEBER (R$)'] = df_numerico['A RECEBER'].apply(formatar_moeda)
df_numerico['RECEBIDO (R$)'] = df_numerico['RECEBIDO'].apply(formatar_moeda)
df_numerico['FALTANDO (R$)'] = df_numerico['FALTANDO'].apply(formatar_moeda)

# Configuração das abas
tab1, tab2 = st.tabs(["Gráfico Interativo", "Planilha Completa"])

# Aba Gráfico Interativo
with tab1:
    st.title('Visualização Interativa')
    categoria = st.selectbox(
        'Selecione uma categoria para o gráfico:', 
        ['A RECEBER', 'RECEBIDO', 'FALTANDO', 'Comparativo: Recebido vs Faltando']
    )
    
    # Verificar se é um comparativo ou uma categoria única
    if categoria == 'Comparativo: Recebido vs Faltando':
        # Preparar DataFrame para exibir os valores no topo
        df_comparativo = pd.melt(df_numerico, id_vars=[df.columns[0]], value_vars=['RECEBIDO', 'FALTANDO'])
        df_comparativo['Valor Formatado'] = df_comparativo['value'].apply(formatar_moeda)
        
        # Plotando o gráfico com os valores em R$ no topo
        fig = px.bar(
            df_comparativo, 
            x=df.columns[0], 
            y='value', 
            color='variable',
            barmode='group',
            text='Valor Formatado',
            title='Comparativo de Valores - Recebido vs Faltando'
        )
        fig.update_traces(textposition='outside')
    else:
        rotulo = f"{categoria} (R$)"
        fig = px.bar(
            df_numerico, 
            x=df.columns[0], 
            y=categoria, 
            text=rotulo,
            title=f'Gráfico de Barras - {categoria}'
        )
        fig.update_traces(textposition='outside')
    
    st.plotly_chart(fig)

# Aba Planilha Completa
with tab2:
    st.title('Visualização da Planilha Completa')
    st.dataframe(df)

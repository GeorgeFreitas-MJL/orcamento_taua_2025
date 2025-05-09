import streamlit as st
import pandas as pd
import io
import plotly.express as px

# ============================
# 1️⃣ CARREGAMENTO DAS PLANILHAS
# ============================

# Caminhos dos arquivos
files = {
    '20RL (CUSTEIO)': 'planilha20rl.xlsx',
    '2994 (ASSISTÊNCIA)': 'planilha2994.xlsx',
    'CAPACITACÃO': 'planilhacapacita.xlsx',
    'DEMANDA 2025': 'planilhanescessaria.xlsx'
}

# Função para carregar a planilha
def carregar_planilha(caminho):
    try:
        return pd.read_excel(caminho)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo {caminho}: {e}")
        return None

# Dicionário para armazenar os DataFrames
planilhas_dfs = {}

# Carregar todas as planilhas
for nome, caminho in files.items():
    df = carregar_planilha(caminho)
    if df is not None:
        # Identificar colunas numéricas e formatar como moeda (R$)
        colunas_numericas = df.select_dtypes(include=['float64', 'int64']).columns
        for coluna in colunas_numericas:
            df[coluna] = df[coluna].apply(lambda x: f"R$ {x:,.2f}" if pd.notnull(x) else "-")
        planilhas_dfs[nome] = df

# ============================
# 2️⃣ EXIBIÇÃO COMPLETA DAS PLANILHAS
# ============================

if planilhas_dfs:
    st.title("DEPARTAMENTO DE ADMINISTRAÇÃO E PLANEJAMENTO (DAP)")
    st.subheader("ORÇAMENTO 'CAMPUS' TAUÁ-CE 2025")

    for nome, df in planilhas_dfs.items():
        st.header(f"📌 {nome}")
        st.dataframe(df, height=50)

# ============================
# 3️⃣ BOTÃO PARA EXPORTAR EM EXCEL
# ============================

try:
    import xlsxwriter
except ImportError:
    st.error("Pacote 'xlsxwriter' não está instalado. Rode: pip install xlsxwriter")

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    for nome, df in planilhas_dfs.items():
        df.to_excel(writer, index=False, sheet_name=nome)
buffer.seek(0)

st.download_button(
    label="📥 Baixar Planilhas Completas em Excel",
    data=buffer,
    file_name='Orcamento_Publico_2025.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)


# Caminho do arquivo atualizado
file_path = 'planilhatabela.xlsx'

# Carregar os dados da nova planilha
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
df_numerico['FALTANDO RECEBER (R$)'] = df_numerico['FALTANDO RECEBER'].apply(formatar_moeda)
df_numerico['NECESSÁRIO PARA 2025 (R$)'] = df_numerico['NECESSÁRIO PARA 2025'].apply(formatar_moeda)

# Configuração das abas
tab1, tab2 = st.tabs(['Gráfico Interativo', 'Planilha Completa'])

# Aba Gráfico Interativo
with tab1:
    st.title('Visualização Interativa')
    categoria = st.selectbox(
        'Selecione uma categoria para o gráfico:', 
        ['A RECEBER', 'RECEBIDO', 'FALTANDO RECEBER', 'NECESSÁRIO PARA 2025', 
         'Comparativo: RECEBIDO vs FALTANDO RECEBER vs NECESSÁRIO']
    )
    
    # Verificar se é um comparativo ou uma categoria única
    if categoria == 'Comparativo: RECEBIDO vs FALTANDO RECEBER vs NECESSÁRIO':
        df_comparativo = pd.melt(df_numerico, id_vars=[df.columns[0]], value_vars=['RECEBIDO', 'FALTANDO RECEBER', 'NECESSÁRIO PARA 2025'])
    elif categoria == 'Comparativo: NECESSÁRIO vs FALTANDO RECEBER':
        df_comparativo = pd.melt(df_numerico, id_vars=[df.columns[0]], value_vars=['NECESSÁRIO PARA 2025', 'FALTANDO RECEBER'])
    else:
        rotulo = f'{categoria} (R$)'
        fig = px.bar(
            df_numerico, 
            x=df.columns[0], 
            y=categoria, 
            text=rotulo,
            title=f'Gráfico de Barras - {categoria}'
        )
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)
        st.stop()

    # Plotando o gráfico comparativo
    df_comparativo['Valor Formatado'] = df_comparativo['value'].apply(formatar_moeda)
    fig = px.bar(
        df_comparativo, 
        x=df.columns[0], 
        y='value', 
        color='variable',
        barmode='group',
        text='Valor Formatado',
        title=f'Comparativo de Valores - {categoria}'
    )
    fig.update_traces(textposition='outside')
    st.plotly_chart(fig)

# Aba Planilha Completa
with tab2:
    st.title('Visualização da Planilha Completa')
    st.dataframe(df)

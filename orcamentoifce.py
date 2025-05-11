import streamlit as st
import plotly.express as px
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(layout="wide")


st.markdown(
    '''
    <style>
        /* Modo Claro */
        @media (prefers-color-scheme: light) {
            body {
                background-color: #eaeaea;
                color: #000000;
            }
            .stApp {
                background-color: transparent;
            }
        }

        /* Modo Escuro */
        @media (prefers-color-scheme: dark) {
            body {
                background-color: #2c2c2c;
                color: #ffffff;
            }
            .stApp {
                background-color: transparent;
            }
        }
    </style>
    ''',
    unsafe_allow_html=True
)



# ============================
#PRIMEIRA PARTE
# ============================

# ============================
# 1️⃣ CARREGAMENTO DAS PLANILHAS
# ============================

# Caminhos dos arquivos

files = {
    'AÇÃO 20RL - CUSTEIO': 'planilha20rl.xlsx', 
    'AÇÃO 2994 - ASSISTÊNCIA': 'planilha2994.xlsx',
    'AÇÃO 4572 - CAPACITACÃO': 'planilhacapacita.xlsx',
    'DEMANDA NECESSÁRIA PARA 2025': 'planilhanescessaria.xlsx'
}

# Função para carregar a planilha
def carregar_planilha(caminho):
    try:
        return pd.read_excel(caminho)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo {caminho}: {e}")
        return None

# Função para estilizar a linha 'Total'
def highlight_total(row):
    if row.astype(str).str.contains(r'(?i)total', regex=True).any():
        return ['background-color: #1C1C1C; font-weight: bold; color: white'] * len(row)
    return ['font-weight: bold; color: black'] * len(row)

# Função para aplicar estilo zebra adaptado para modo escuro e claro
def zebra_style(df):
    styles = pd.DataFrame('font-weight: bold', index=df.index, columns=df.columns)
    for index, row in df.iterrows():
        if not row.astype(str).str.contains(r'(?i)total', regex=True).any():
            if index % 2 == 0:
                styles.loc[index, :] += '; background-color: #708090; color: white'
            else:
                styles.loc[index, :] += '; background-color: #e6e6e6; color: black'
    return styles

# Função para aplicar os estilos
def header_style(df):
    styled = df.style.apply(highlight_total, axis=1)
    styled = styled.apply(zebra_style, axis=None)
    return styled

# Dicionário para armazenar os DataFrames
planilhas_dfs = {}

# Carregar todas as planilhas
for nome, caminho in files.items():
    df = carregar_planilha(caminho)
    if df is not None:
        # Identificar colunas numéricas e tentar converter para float
        colunas_numericas = df.select_dtypes(include=['float64', 'int64']).columns
        
        # Forçar a conversão para números; se falhar, coloca NaN
        for coluna in colunas_numericas:
            df[coluna] = pd.to_numeric(df[coluna], errors='coerce')
        
        # Substituir NaN e Inf
        df.replace([np.inf, -np.inf], 0, inplace=True)
        df.fillna("-", inplace=True)

        # Aplicar a formatação de moeda com ponto em vez de vírgula
        for coluna in colunas_numericas:
            df[coluna] = df[coluna].apply(lambda x: f"R$ {x:,.2f}".replace(",", ".") if isinstance(x, (int, float)) else x)

        # Guardar no dicionário
        planilhas_dfs[nome] = df

# ============================
# 2️⃣ EXIBIÇÃO COMPLETA DAS PLANILHAS
# ============================

if planilhas_dfs:
    # Título principal em verde e centralizado
    st.markdown("""
    <h1 style='color: ; margin-bottom: 80px; text-align: center;'>
        DEPARTAMENTO DE ADMINISTRAÇÃO E PLANEJAMENTO (DAP) - <em>CAMPUS</em> TAUÁ-CE <br> ORÇAMENTO 2025
    </h1>
    """, unsafe_allow_html=True)


    # Subtítulo em verde, um pouco menor, centralizado
st.markdown("""
   <h6 style='color: ; margin-bottom: 0px; text-align: right;'>Passe o mouse sobre a borda para visualizar as opções e acessar todas as planilhas e os gráficos abaixo ⤵️</h6>
   """, unsafe_allow_html=True)

st.divider()
st.title("AÇÕES - RECURSOS:")

for nome, df in planilhas_dfs.items():
        st.header(f"📗 {nome}")
        st.dataframe(header_style(df), height=5)

# ============================
# 3️⃣ BOTÃO PARA EXPORTAR EM EXCEL COM ESTILO
# ============================
 
# Estilo para o botão
st.markdown("""
    <style>
    .stDownloadButton button {
        background-color: #28a745;
        color: white;
        border-radius: 5px;
        border: none;
    }
    .stDownloadButton button:hover {
        color: red !important;
    }
    </style>
    """, unsafe_allow_html=True) 
 
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


# ============================
#SEGUNDA PARTE!
# ============================

# ============================
# 1️⃣ CARREGAMENTO DAS PLANILHAS
# ============================

# Caminhos dos arquivos
files = {
    'AÇÃO 20RL - CUSTEIO': 'planilha20rl.xlsx', 
    'AÇÃO 2994 - ASSISTÊNCIA': 'planilha2994.xlsx',
    'AÇÃO 4572 - CAPACITACÃO': 'planilhacapacita.xlsx',
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
planilhas_dfs_graficos = {}

# Carregar todas as planilhas
for nome, caminho in files.items():
    df = carregar_planilha(caminho)
    if df is not None:
        # Identificar colunas numéricas e tentar converter para float
        colunas_numericas = df.select_dtypes(include=['float64', 'int64']).columns

        # Cópia para gráficos (sem formatação)
        df_grafico = df.copy()

        # Forçar a conversão para números; se falhar, coloca NaN
        for coluna in colunas_numericas:
            df[coluna] = pd.to_numeric(df[coluna], errors='coerce')

        # Substituir NaN e Inf
        df.replace([np.inf, -np.inf], 0, inplace=True)
        df.fillna("-", inplace=True)

        # Formatação dos valores para R$ (Real Brasileiro)
        for coluna in colunas_numericas:
            df[coluna] = df[coluna].map(lambda x: f"R$ {x:.2f}".replace(",", ".") if isinstance(x, (int, float)) else x)

        # Guardar no dicionário
        planilhas_dfs[nome] = df
        planilhas_dfs_graficos[nome] = df_grafico



#GRÁFICOS INTERATIVOS POR PLANILHA

st.divider()
st.title("AÇÕES - PAGAMENTOS:")


# Seleciona a planilha para visualização do gráfico
selected_sheet = st.selectbox("Escolha a AÇÃO para visualização do gráfico ⤵️", list(planilhas_dfs.keys()))

if selected_sheet:
    selected_df = planilhas_dfs_graficos[selected_sheet]

    # Remover linhas vazias, apenas se a coluna existir
    if 'PAGAMENTO REALIZADO' in selected_df.columns:
        selected_df = selected_df[selected_df['PAGAMENTO REALIZADO'] > 0]

    if 'SALDO NECESSÁRIO' in selected_df.columns:
        selected_df = selected_df[selected_df['SALDO NECESSÁRIO'] > 0]

    # Definindo rótulos e valores com base na planilha e calculando percentuais
    if selected_sheet == 'AÇÃO 20RL - CUSTEIO':
        selected_df['Percentual'] = selected_df['PAGAMENTO REALIZADO'] / selected_df['PAGAMENTO REALIZADO'].sum() * 100
        selected_df['Label'] = selected_df.apply(lambda row: f"{row['AÇÃO 20RL - CUSTEIO']} - R$ {row['PAGAMENTO REALIZADO']:.2f} ({row['Percentual']:.2f}%)", axis=1)
    elif selected_sheet == 'AÇÃO 2994 - ASSISTÊNCIA':
        selected_df['Percentual'] = selected_df['PAGAMENTO REALIZADO'] / selected_df['PAGAMENTO REALIZADO'].sum() * 100
        selected_df['Label'] = selected_df.apply(lambda row: f"{row['AÇÃO 2994 - ASSISTÊNCIA']} - R$ {row['PAGAMENTO REALIZADO']:.2f} ({row['Percentual']:.2f}%)", axis=1)
    elif selected_sheet == 'AÇÃO 4572 - CAPACITACÃO':
        selected_df['Percentual'] = selected_df['PAGAMENTO REALIZADO'] / selected_df['PAGAMENTO REALIZADO'].sum() * 100
        selected_df['Label'] = selected_df.apply(lambda row: f"{row['AÇÃO 4572 - CAPACITAÇÃO']} - R$ {row['PAGAMENTO REALIZADO']:.2f} ({row['Percentual']:.2f}%)", axis=1)


    # Ajuste de layout do gráfico
    cores_personalizadas = px.colors.qualitative.Set3 + px.colors.qualitative.Pastel + px.colors.qualitative.Safe
    fig = px.pie(selected_df, values='Percentual', names='Label',
                 hole=0.3, color_discrete_sequence=cores_personalizadas)

    st.subheader(f"Distribuição Percentual de valores - {selected_sheet}")
    st.plotly_chart(fig, use_container_width=True)





# ============================
#TERCEIRA PARTE!
# ============================

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
df_numerico['ORÇAMENTO (R$)'] = df_numerico['ORÇAMENTO'].apply(formatar_moeda)
df_numerico['RECEBIDO (R$)'] = df_numerico['RECEBIDO'].apply(formatar_moeda)
df_numerico['FALTANDO RECEBER (R$)'] = df_numerico['FALTANDO RECEBER'].apply(formatar_moeda)
df_numerico['NECESSÁRIO PARA 2025 (R$)'] = df_numerico['NECESSÁRIO PARA 2025'].apply(formatar_moeda)


# Configuração das abas
st.divider()
st.title("AÇÕES - FLUXO DE RECURSOS:")

tab1, tab2 = st.tabs(['Gráfico Interativo', 'Planilha Completa'])

# Estilo CSS para os botões
st.markdown("""
    <style>
    .stButton button {
        background-color: #28a745 !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px;
        border: none;
        height: 70px;
    }
    .stButton button:hover {
        color: red !important;
    }
    </style>
    """, unsafe_allow_html=True)



# Aba Gráfico Interativo
with tab1:
    st.markdown('Clique nas opções abaixo ⤵️')
    
    # Botões horizontais
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        btn1 = st.button('COMPARAR FLUXO DE RECURSO')
    with col2:
        btn2 = st.button('ORÇAMENTO 2025')
    with col3:
        btn3 = st.button('RECURSO RECEBIDO')
    with col4:
        btn4 = st.button('FALTANDO RECEBER')
    with col5:
        btn5 = st.button('NECESSÁRIO PARA 2025')
    
    # Renderização dos gráficos com base no botão clicado
    if btn1:
        df_comparativo = pd.melt(df_numerico, id_vars=[df.columns[0]], value_vars=['RECEBIDO', 'FALTANDO RECEBER', 'NECESSÁRIO PARA 2025'])
        df_comparativo['Valor Formatado'] = df_comparativo['value'].apply(formatar_moeda)
        fig = px.bar(
            df_comparativo, 
            x=df.columns[0], 
            y='value', 
            color='variable',
            barmode='group',
            text='Valor Formatado',
            title='Comparativo de Valores - RECEBIDO vs FALTANDO RECEBER vs NECESSÁRIO'
        )
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)
    
    elif btn2:
        fig = px.bar(
            df_numerico, 
            x=df.columns[0], 
            y='ORÇAMENTO', 
            text='ORÇAMENTO (R$)',
            title='Gráfico de Barras - ORÇAMENTO'
        )
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)

    elif btn3:
        fig = px.bar(
            df_numerico, 
            x=df.columns[0], 
            y='RECEBIDO', 
            text='RECEBIDO (R$)',
            title='Gráfico de Barras - RECEBIDO'
        )
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)

    elif btn4:
        fig = px.bar(
            df_numerico, 
            x=df.columns[0], 
            y='FALTANDO RECEBER', 
            text='FALTANDO RECEBER (R$)',
            title='Gráfico de Barras - FALTANDO RECEBER'
        )
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)

    elif btn5:
        fig = px.bar(
            df_numerico, 
            x=df.columns[0], 
            y='NECESSÁRIO PARA 2025', 
            text='NECESSÁRIO PARA 2025 (R$)',
            title='Gráfico de Barras - NECESSÁRIO PARA 2025'
        )
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)

# Aba Planilha Completa
with tab2:
    st.title('Visualização da Planilha Completa')
    st.dataframe(df)
""


# LINK CLICÁVEL ESTILIZADO
st.markdown("---")
st.markdown("""
    <div style='text-align: left; margin-top: 20px;'>
        <a href='https://orcamento.ifce.edu.br/' target='_blank' 
        style='font-size: 18px; color: #6495ED; text-decoration: none; font-weight: bold;'>
            🔗 Clique para acessar: ORÇAMENTO DA REDE.
        </a>
    </div>
    """, unsafe_allow_html=True)
    
  # RODAPÉ ESTILIZADO 
st.markdown("""
    <hr style='margin-top: 50px;'>
    <div style='text-align: center; color: red !important; font-size: 14px;'>
        Desenvolvido pelo DAP-TAUÁ-2025 - Todos os direitos reservados.<br>
        💡 Tem dúvidas ou precisa de suporte? Estamos à disposição para ajudar!
    </div>
    """, unsafe_allow_html=True)  
    
    
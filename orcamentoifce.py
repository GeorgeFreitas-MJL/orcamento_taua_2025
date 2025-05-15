
import streamlit as st
import plotly.express as px
import pandas as pd
import io
import numpy as np
import plotly.graph_objects as go


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
# CARREGAMENTO DAS PLANILHAS

# Caminhos dos arquivos

files = {
    'AÇÃO 20RL - CUSTEIO': 'planilha20rl.xlsx', 
    'AÇÃO 2994 - ASSISTÊNCIA': 'planilha2994.xlsx',
    'AÇÃO 4572 - CAPACITACÃO': 'planilhacapacita.xlsx',
    'DEMANDA NECESSÁRIA PARA 2025': 'planilhanescessaria.xlsx',
    'AÇÕES EM PROCESSAMENTO': 'planilhanegativa.xlsx'   
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
# EXIBIÇÃO COMPLETA DAS PLANILHAS

if planilhas_dfs:
    # Título principal em verde e centralizado
    st.markdown("""
    <h1 style='color: ; margin-bottom: 80px; text-align: center;'>
        DEPARTAMENTO DE ADMINISTRAÇÃO E PLANEJAMENTO (DAP) <br>
        <em>CAMPUS</em> TAUÁ-CE - ORÇAMENTO 2025
        </h1>
    """, unsafe_allow_html=True)
 
    # Subtítulo em verde, um pouco menor, centralizado
st.markdown("""
   <h6 style='color: ; margin-bottom: 0px; text-align: right;'>🖱️ Passe o mouse sobre as bordas para visualizar as opções e acessar todas as planilhas e gráficos listados abaixo ⤵️</h6>
   """, unsafe_allow_html=True)

st.divider()
st.title("➡️ AÇÕES - RECURSOS")

for nome, df in planilhas_dfs.items():
        st.header(f" {nome}")
        st.dataframe(header_style(df), height=5)

# ============================
# BOTÃO PARA EXPORTAR EM EXCEL COM ESTILO

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
# Interface Streamlit
st.divider()
st.title('➡️ AÇÕES - PAGAMENTOS')

# ============================
# CARREGAMENTO DAS PLANILHAS

# Caminhos dos arquivos
file_20rl = 'planilha20rl.xlsx'
file_2994 = 'planilha2994.xlsx'

# Ler as planilhas
df_20rl = pd.read_excel(file_20rl)
df_2994 = pd.read_excel(file_2994)

# ============================
# PREPARAÇÃO DOS DADOS

# Ajuste nos DataFrames para o gráfico de pizza
df_20rl['Label'] = df_20rl.apply(lambda x: f"{x['AÇÃO 20RL - CUSTEIO']} - R$ {x['PAGAMENTO REALIZADO']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), axis=1)
df_20rl['Percentual'] = df_20rl['PAGAMENTO REALIZADO'] / df_20rl['PAGAMENTO REALIZADO'].sum() * 100

df_2994['Label'] = df_2994.apply(lambda x: f"{x['AÇÃO 2994 - ASSISTÊNCIA']} - R$ {x['PAGAMENTO REALIZADO']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'), axis=1)
df_2994['Percentual'] = df_2994['PAGAMENTO REALIZADO'] / df_2994['PAGAMENTO REALIZADO'].sum() * 100

# ============================
# INTERFACE STREAMLIT

# Exibir título principal
#st.title("Gráficos de Distribuição Percentual")

# Filtro interativo para escolha da Ação
opcoes = ['AÇÃO 20RL - CUSTEIO', 'AÇÃO 2994 - ASSISTÊNCIA']
escolha = st.selectbox("🖱️ Selecione a AÇÃO para visualizar o gráfico ⤵️", opcoes)

if escolha == 'AÇÃO 20RL - CUSTEIO':
    st.subheader("Distribuição Percentual:")
    fig = px.pie(df_20rl, values='Percentual', names='Label', hole=0.5)
    st.plotly_chart(fig, use_container_width=True)

elif escolha == 'AÇÃO 2994 - ASSISTÊNCIA':
    st.subheader("Distribuição Percentual:")
    fig = px.pie(df_2994, values='Percentual', names='Label', hole=0.5)
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
st.title("➡️ AÇÕES - FLUXO DE RECURSOS")

tab1, tab2 = st.tabs(['🖱️ Gráfico Interativo', '🖱️ Planilha Completa'])

# Estilo CSS para os botões e abas
st.markdown('''<style>
.stButton button {
    background-color: #28a745 !important;
    color: white !important;
    font-weight: bold !important;
    border-radius: 8px;
    border: none;
    height: 40px;
}
.stButton button:hover {
    color: red !important;
}

/* Estilização das abas para limitar a linha vermelha */
div[data-baseweb="tab-list"] {
    display: flex;
    justify-content: center;
    border-bottom: 2px solid transparent;
}

div[data-baseweb="tab"] {
    flex: 1;
    text-align: center;
}

div[data-baseweb="tab"][aria-selected="true"] {
    border-bottom: 2px solid red !important;
}
</style>
''', unsafe_allow_html=True)

# Aba Gráfico Interativo
with tab1:
    st.markdown('🖱️ Clique nas opções abaixo ⤵️')
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        btn1 = st.button('FLUXO DE RECURSO')
    with col2:
        btn2 = st.button('ORÇAMENTO 2025')
    with col3:
        btn3 = st.button('RECURSO RECEBIDO')
    with col4:
        btn4 = st.button('FALTANDO RECEBER')
    with col5:
        btn5 = st.button('NECESSÁRIO PARA 2025')

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
        fig = px.bar(df_numerico, x=df.columns[0], y='ORÇAMENTO', text='ORÇAMENTO (R$)', title='Gráfico de Barras - ORÇAMENTO')
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)
    elif btn3:
        fig = px.bar(df_numerico, x=df.columns[0], y='RECEBIDO', text='RECEBIDO (R$)', title='Gráfico de Barras - RECEBIDO')
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)
    elif btn4:
        fig = px.bar(df_numerico, x=df.columns[0], y='FALTANDO RECEBER', text='FALTANDO RECEBER (R$)', title='Gráfico de Barras - FALTANDO RECEBER')
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)
    elif btn5:
        fig = px.bar(df_numerico, x=df.columns[0], y='NECESSÁRIO PARA 2025', text='NECESSÁRIO PARA 2025 (R$)', title='Gráfico de Barras - NECESSÁRIO PARA 2025')
        fig.update_traces(textposition='outside')
        st.plotly_chart(fig)

# Aba Planilha Completa
with tab2:
    st.markdown('Visualização da Planilha Completa ⤵️')
    st.dataframe(df)


# ============================
#QUARTA PARTE!
# ============================

st.divider()
st.title("➡️ AÇÕES - NEGATIVADAS")

# Título da aplicação
# st.markdown("💰 Game Life Financeiro - Recebimentos vs Total Negativado")

# Carregar os dados
try:
    # Carregando a planilha
    data = pd.read_excel('planilhanegativa.xlsx', sheet_name='Página1')
    
    # Remover espaços em branco dos nomes das colunas
    data.columns = data.columns.str.strip()
    
    # Filtrar os dados necessários
    final_data = pd.DataFrame({
        'Ação': data['AÇÃO'],
        'Recurso Recebido': data['RECURSO RECEBIDO'],
        'Total Negativado': data['TOTAL NEGATIVADO']
    })

    # Adicionando colunas para o gráfico
    final_data = final_data.melt(id_vars=['Ação'], var_name='Tipo', value_name='Valor')
    final_data['Cor'] = final_data['Tipo'].map({
        'Recurso Recebido': 'blue',
        'Total Negativado': 'red'
    })
    
    # Formatando os valores para Real Brasileiro
    final_data['Valor Formatado'] = final_data['Valor'].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

except FileNotFoundError:
    st.error("⚠️ Arquivo 'planilhanegativa.xlsx' não encontrado. Coloque-o no mesmo diretório do script.")
    st.stop()

# ============================
# 🔄 GRÁFICO INTERATIVO
# ============================

fig = px.bar(
    final_data,
    x='Valor',
    y='Ação',
    color='Tipo',
    text='Valor Formatado',
    orientation='h',
    color_discrete_map={
        'Recurso Recebido': 'blue',
        'Total Negativado': 'red'
    },
    title="💰 RECURSOS POSITIVOS x RECURSOS NEGATIVOS"
)

fig.update_traces(textposition='inside')
fig.update_layout(
    xaxis_title="Valor em R$",
    yaxis_title="Ação",
    legend_title="Categoria",
    height=400
)

# Exibir o gráfico no Streamlit
st.plotly_chart(fig, use_container_width=True)

# Mostrar a tabela de dados
with st.expander("🔎 Visualizar Dados"):
    st.dataframe(final_data)

# ============================
#FINAL
# ============================

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
        🖥️ Desenvolvido pelo DAP-TAUÁ-2025 - Todos os direitos reservados.<br>
        💡 Tem dúvidas ou precisa de suporte? Estamos à disposição para ajudar!<br>
        📨 E-mail: george.luiz@ifce.edu.br
    </div>
    """, unsafe_allow_html=True)  
    
    




import plotly.express as px
import streamlit as st
import pandas as pd
import io
import xlsxwriter
import numpy as np

st.set_page_config(layout="wide")


st.markdown(
    """
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
    """,
    unsafe_allow_html=True
)


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

        # Aplicar a formatação de moeda apenas para valores válidos
        for coluna in colunas_numericas:
            df[coluna] = df[coluna].apply(lambda x: f"R$ {x:,.2f}" if isinstance(x, (int, float)) else x)

        # Guardar no dicionário
        planilhas_dfs[nome] = df

# ============================
# 2️⃣ EXIBIÇÃO COMPLETA DAS PLANILHAS
# ============================

if planilhas_dfs:
    # st.title("DEPARTAMENTO DE ADMINISTRAÇÃO E PLANEJAMENTO (DAP)")
    # st.subheader("ORÇAMENTO 'CAMPUS' TAUÁ-CE 2025")
    st.markdown("<h1 style='color: #4CAF50; margin-bottom: 40px;;'>DEPARTAMENTO DE ADMINISTRAÇÃO E PLANEJAMENTO (DAP)</h1>", unsafe_allow_html=True)
    #st.markdown("<h2 style='color: #4CAF50;'>Orçamento do *campus* Tauá-CE em 2025</h2>", unsafe_allow_html=True)
    #st.markdown("<h2 style='color: #4CAF50; margin-bottom: 40px;;'>Orçamento do <em>campus</em> Tauá-CE em 2025</h2>", unsafe_allow_html=True)
    #st.markdown("""<h2 style='color: black; margin-bottom: 05px; text-align: center;'>Orçamento do <em>campus</em> Tauá-CE em 2025</h2>""", unsafe_allow_html=True)
    st.markdown("""
    <style>
        /* Ajuste para modo claro e escuro */
        .titulo-campus {
            color: black;
        }

        @media (prefers-color-scheme: dark) {
            .titulo-campus {
                color: white;
            }
        }
    </style>

    <h2 class='titulo-campus' style='margin-bottom: 5px; text-align: center;'>
        Orçamento do <em>campus</em> Tauá-CE em 2025
    </h2>
""", unsafe_allow_html=True)

    for nome, df in planilhas_dfs.items():
        st.header(f"📌 {nome}")
        st.dataframe(header_style(df), height=5)

# ============================
# 3️⃣ BOTÃO PARA EXPORTAR EM EXCEL COM ESTILO
# ============================

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    for nome, df in planilhas_dfs.items():
        df.to_excel(writer, index=False, sheet_name=nome)

        # Aplicar estilos no Excel
        workbook = writer.book
        worksheet = writer.sheets[nome]

        format_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
        format_zebra1 = workbook.add_format({'bg_color': '#f0f0f0'})
        format_zebra2 = workbook.add_format({'bg_color': '#e6e6e6'})
        format_total = workbook.add_format({'bold': True, 'bg_color': '#4B0082', 'font_color': 'white'})

        for idx, col in enumerate(df.columns):
            worksheet.write(0, idx, col, format_header)

        for row_num, row_data in df.iterrows():
            style = format_total if "Total" in row_data.to_string() else (format_zebra1 if row_num % 2 == 0 else format_zebra2)
            for col_num, value in enumerate(row_data):
                # Tratamento para NaN e Inf
                if pd.isna(value) or value == "-" or value == np.inf or value == -np.inf:
                    worksheet.write(row_num + 1, col_num, "", style)
                else:
                    worksheet.write(row_num + 1, col_num, value, style)

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
    st.title('Visualização Interativa \n Clique nas opções abaixo:')
    
    # Menu interativo na horizontal com estilo melhorado
    st.markdown(
        '''
        <style>
        .stButton button {
            margin: 5px;
            background-color: #4CAF50;
            color: white;
            border-radius: 5px;
        }
        </style>
        ''',
        unsafe_allow_html=True
    )
    
    # Botões horizontais
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        btn1 = st.button('COMPARAR')
    with col2:
        btn2 = st.button('ORÇAMENTO')
    with col3:
        btn3 = st.button('RECEBIDO')
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
            y='A RECEBER', 
            text='A RECEBER (R$)',
            title='Gráfico de Barras - A RECEBER'
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



# ============================
# 4️⃣ RODAPÉ ESTILIZADO
# ============================

st.markdown(
    """
    <style>
        .footer {
            width: 100%;
            background-color: #4CAF50;
            color: white;
            text-align: center;
            padding: 10px;
            font-size: 14px;
            margin-top: 20px;
            border-radius: 5px;
        }
    </style>
    
    <div class="footer">
        💡 Tem dúvidas ou precisa de suporte? O DAP está à disposição. Estamos aqui para ajudar!
    </div>
    """,
    unsafe_allow_html=True
)

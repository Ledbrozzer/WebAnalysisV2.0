# pip install streamlit
# pip install openpyxl
# streamlit run Analysis_test.py
# $ streamlit run yourscript.py
import streamlit as st
import pandas as pd
import plotly.express as px
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
#Read Excel File
tabela = pd.read_excel("Abastecimento_Caminhao.xlsx")
tabelaPlaca = pd.read_excel("RELAÇÃO FROTA ATUALIZADO 251124.xlsx")
#Convert column'Data Req.'t/typ datetime
tabela['Data Req.'] = pd.to_datetime(tabela['Data Req.'], errors='coerce', dayfirst=True)
#Exclu colun indesejd
columns_exclud = ["Combustível", "Vlr. Unitário", "Hora Abast.", "Obs.", "Abast. Externo"]
current_columns = [col for col in columns_exclud if col in tabela.columns]
if current_columns:
    tabela = tabela.drop(columns=current_columns)
else:
    print("Nenhuma das colunas especificadas foi encontrada no DataFrame.")
#Order data pela'Data Req.'d-form ascend
tabela = tabela.sort_values(by=['Data Req.'])
#Define datas p/filtr init|final
dtini = "01/06/2024"  #Data inic p/cálculr
dtfinal = "01/07/2024"  #Data inic p/exib
data_inicial_calculo = pd.to_datetime(dtini, format='%d/%m/%Y')
data_inicial_exibicao = pd.to_datetime(dtfinal, format='%d/%m/%Y')
data_final = pd.Timestamp.now()
#Apply t-filter from dates t/calculate
tabela_calculo = tabela[(tabela['Data Req.'] >= data_inicial_calculo) & (tabela['Data Req.'] <= data_final)].copy()
#Calc dif->Km|Horim p/kdá Veículo/Equip.
tabela_calculo['Diferença de Km'] = tabela_calculo.groupby('Veículo/Equip.')['Km Atual'].diff().abs()
tabela_calculo['Diferença de Horim'] = tabela_calculo.groupby('Veículo/Equip.')['Horim. Equip.'].diff().abs()
#GET value->Litros d-abastecimento befor t/each Veículo/Equip.
tabela_calculo['Litros Anterior'] = tabela_calculo.groupby('Veículo/Equip.')['Litros'].shift(1)
#Calc Km&Horim/Litro
tabela_calculo['Km por Litro'] = tabela_calculo['Diferença de Km'] / tabela_calculo['Litros Anterior']
tabela_calculo['Horim por Litro'] = tabela_calculo['Diferença de Horim'] / tabela_calculo['Litros Anterior']
#Arredond values
tabela_calculo['Km por Litro'] = tabela_calculo['Km por Litro'].round(3)
tabela_calculo['Horim por Litro'] = tabela_calculo['Horim por Litro'].round(3)
#Filtr data t/exibt only after'01/07/2024'
tabela_exibicao = tabela_calculo[tabela_calculo['Data Req.'] >= data_inicial_exibicao]
#Reformat colun'Data Req.'p/exib format desejd using .loc[]
tabela_exibicao.loc[:, 'Data Req.'] = tabela_exibicao['Data Req.'].dt.strftime('%d/%m/%Y')
#Mescl tabl'tabela_exibicao'c/tabl'tabelaPlaca'p/includ colun"PLACA/"
tabela_exibicao = tabela_exibicao.merge(tabelaPlaca[['Placa TOPCON', 'PLACA/']], left_on='Veículo/Equip.', right_on='Placa TOPCON', how='left')
#Reorganize columns'Order
colunas_ordem = ["Requisição", "Data Req.", "Requisitante", "PLACA/", "Diferença de Km", "Km por Litro", "Vlr. Total", "Km Atual", "Horim por Litro", "Horim. Equip."]
colunas_restantes = [col for col in tabela_exibicao.columns if col not in colunas_ordem]
tabela_exibicao = tabela_exibicao[colunas_ordem + colunas_restantes]
#Config d-aplic Streamlit
st.title('Análise de Abastecimento')
st.sidebar.header('Filtrar os Dados')
#Entrad d-user p/filtrgm
requisitante = st.sidebar.text_input('Requisitante', '')
veiculo = st.sidebar.text_input('Veículo', '')
data_inicial = st.sidebar.date_input('Data inicial', pd.to_datetime('2024-07-01'))
data_final = st.sidebar.date_input('Data final', pd.Timestamp.now())
#Apply filtr t/t-data
filtro = tabela_exibicao[(tabela_exibicao['Requisitante'].str.contains(requisitante, case=False, na=False)) &
                         (tabela_exibicao['Veículo/Equip.'].str.contains(veiculo, case=False, na=False)) &
                         (pd.to_datetime(tabela_exibicao['Data Req.'], format='%d/%m/%Y') >= pd.to_datetime(data_inicial)) &
                         (pd.to_datetime(tabela_exibicao['Data Req.'], format='%d/%m/%Y') <= pd.to_datetime(data_final))]
#Show data filtrd
st.write("Dados Filtrados:")
st.write(filtro)
#Seleç análise
analise = st.sidebar.selectbox(
    'Selecione a Análise',
    ('Análise 1: Diferença de Km', 'Análise 2: Km por Litro', 'Análise 3: Horim por Litro', 'Análise 4: Km/Litro por Data', 'Análise 5: Performance Requisitante')
)
#Analysis'Functions
def analise1(filtro):
    fig = px.histogram(filtro, x='Diferença de Km', color='Diferença de Km', title='Análise 1: Diferença de Km')
    return fig
def analise2(filtro):
    fig = px.histogram(filtro, x='Km por Litro', color='Km por Litro', title='Análise 2: Km por Litro')
    return fig
def analise3(filtro):
    fig = px.histogram(filtro, x='Horim por Litro', color='Horim por Litro', title='Análise 3: Horim por Litro')
    return fig
def analise4(filtro):
    fig = px.histogram(filtro, x='Data Req.', y='Km por Litro', color='Km por Litro', hover_data=['Veículo/Equip.'], title='Análise 4: Km/Litro por Data')
    return fig
def analise5(filtro):
    fig = px.histogram(filtro, x='Km por Litro', y=['Requisitante'], color='Km por Litro', hover_data=['Data Req.'], title='Análise 5: Performance Requisitante')
    return fig
#Show gráfc selectd
if analise == 'Análise 1: Diferença de Km':
    fig = analise1(filtro)
elif analise == 'Análise 2: Km por Litro':
    fig = analise2(filtro)
elif analise == 'Análise 3: Horim por Litro':
    fig = analise3(filtro)
elif analise == 'Análise 4: Km/Litro por Data':
    fig = analise4(filtro)
elif analise == 'Análise 5: Performance Requisitante':
    fig = analise5(filtro)
st.plotly_chart(fig)
#Buttn p/export data filtrad p/Excel
if st.button('Exportar Dados Filtrados para Excel'):
    with pd.ExcelWriter('dados_filtrados.xlsx', engine='openpyxl') as writer:
        filtro.to_excel(writer, index=False, sheet_name='Dados Filtrados')
    st.write('Dados exportados para Excel com sucesso!')
#Link p/download arqv Excel
with open('dados_filtrados.xlsx', 'rb') as f:
    st.download_button('Baixar Dados Filtrados', f, file_name='dados_filtrados.xlsx')
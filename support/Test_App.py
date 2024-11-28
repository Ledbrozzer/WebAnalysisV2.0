# pip install streamlit
# pip install openpyxl
# streamlit run SL_App.py
# $ streamlit run yourscript.py
import streamlit as st
import pandas as pd
import plotly.express as px
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
#Ler Excel
tabela = pd.read_excel("Abastecimento_Caminhao.xlsx")
tabelaPlaca = pd.read_excel("RELAÇÃO FROTA ATUALIZADO 251124.xlsx")
#Convert colun'Data Req.'t/type datetime
tabela['Data Req.'] = pd.to_datetime(tabela['Data Req.'], errors='coerce', dayfirst=True)
#Exclu coluns indesejds
columns_exclud = ["Combustível", "Vlr. Unitário", "Hora Abast.", "Abast. Externo"]
current_columns = [col for col in columns_exclud if col in tabela.columns]
if current_columns:
    tabela = tabela.drop(columns=current_columns)
#Order data pel'Data Req.'form ascend
tabela = tabela.sort_values(by=['Data Req.'])
#Calc dif Km&&Horim p/cad Veículo/Equip.
tabela['Diferença de Km'] = tabela.groupby('Veículo/Equip.')['Km Atual'].diff().abs()
tabela['Diferença de Horim'] = tabela.groupby('Veículo/Equip.')['Horim. Equip.'].diff().abs()
#Obte value-Litros-abasteciment befor p/cad-Veículo/Equip.
tabela['Litros Anterior'] = tabela.groupby('Veículo/Equip.')['Litros'].shift(1)
#Calc Km&&Horim/Litro
tabela['Km por Litro'] = tabela['Diferença de Km'] / tabela['Litros Anterior']
tabela['Horim por Litro'] = tabela['Diferença de Horim'] / tabela['Litros Anterior']
#Arredond values
tabela['Km por Litro'] = tabela['Km por Litro'].round(3)
tabela['Horim por Litro'] = tabela['Horim por Litro'].round(3)
#Cal Combustív Restant
tabela['Combustível Restante'] = tabela['Diferença de Km'] % tabela['Litros Anterior']
tabela['Combustível Restante'] = tabela['Combustível Restante'].round(3)
#Reformat colun'Data Req.'p/exib-format desejd usng .loc[]
tabela.loc[:, 'Data Req.'] = tabela['Data Req.'].dt.strftime('%d/%m/%Y')
#Mescl tabl'tabela'c/tabl'tabelaPlaca't/include colun"PLACA/"
tabela = tabela.merge(tabelaPlaca[['Placa TOPCON', 'PLACA/']], left_on='Veículo/Equip.', right_on='Placa TOPCON', how='left')
#Reorganize coluns'order
colunas_ordem = ["Requisição", "Data Req.", "Requisitante", "PLACA/", "Diferença de Km", "Km por Litro", "Combustível Restante", "Vlr. Total", "Km Atual", "Km Rodados", "Horim por Litro", "Horim. Equip.", "Litros Anterior", "Litros", "Diferença de Horim", "Veículo/Equip.", "Obs."]
tabela = tabela[colunas_ordem]
#Ajust data type p/compatibility c/Arrow
tabela['Obs.'] = tabela['Obs.'].astype(str)
#Config-application Streamlit
st.title('Análise de Abastecimento')
st.sidebar.header('Filtrar os Dados')
#Entrad user p/filtrgm
requisitante = st.sidebar.text_input('Requisitante', '')
veiculo = st.sidebar.text_input('Veículo', '')
data_inicial = st.sidebar.date_input('Data inicial', pd.to_datetime('2024-01-01'))
data_final = st.sidebar.date_input('Data final', pd.Timestamp.now())
km_litro_min = st.sidebar.number_input('Km por Litro (Mínimo)', value=0.0, step=0.1)
km_litro_max = st.sidebar.number_input('Km por Litro (Máximo)', value=100.0, step=0.1)
#ApplyFiltr t/t-data
filtro = tabela[(tabela['Requisitante'].str.contains(requisitante, case=False, na=False)) &
                (tabela['Veículo/Equip.'].str.contains(veiculo, case=False, na=False)) &
                (pd.to_datetime(tabela['Data Req.'], format='%d/%m/%Y') >= pd.to_datetime(data_inicial)) &
                (pd.to_datetime(tabela['Data Req.'], format='%d/%m/%Y') <= pd.to_datetime(data_final)) &
                (tabela['Km por Litro'] >= km_litro_min) &
                (tabela['Km por Litro'] <= km_litro_max)]
#Order data filtrad By'Data Req.'ascending order
filtro = filtro.sort_values(by=['Data Req.'])
#Show data filtrd
st.write("Dados Filtrados:")
st.write(filtro)
#Selection analys
analise = st.sidebar.selectbox(
    'Selecione a Análise',
    ('Análise 1: Diferença de Km(x)', 'Análise 2: Km por Litro(x)', 'Análise 3: Horim por Litro(x)', 'Análise 4: Km/Litro por Data', 'Análise 5: Performance Requisitante', 'Análise 6: Performance por Veículo', 'Análise 7: Km/Litro por Vlr Total', 'Análise 8: Top5|Bottom10 Km/Litro')
)
#Functns analysis
def analise1(filtro):
    fig = px.histogram(filtro, x='Diferença de Km', color='Diferença de Km', hover_data=['Veículo/Equip.'], title='Análise 1: Diferença de Km(x)')
    return fig
def analise2(filtro):
    fig = px.histogram(filtro, x='Km por Litro', color='Km por Litro', hover_data=['Veículo/Equip.', 'Data Req.'], title='Análise 2: Km por Litro(x)')
    return fig
def analise3(filtro):
    fig = px.histogram(filtro, x='Horim por Litro', color='Horim por Litro', hover_data=['Veículo/Equip.', 'Data Req.'], title='Análise 3: Horim por Litro(x)')
    return fig
def analise4(filtro):
    fig = px.histogram(filtro, x='Data Req.', y='Km por Litro', color='Km por Litro', hover_data=['Veículo/Equip.'], title='Análise 4: Km/Litro por Data')
    return fig
def analise5(filtro):
    fig = px.histogram(filtro, x='Km por Litro', y='Requisitante', color='Requisitante', hover_data=['Data Req.'], title='Análise 5: Performance Requisitante')
    return fig
def analise6(filtro):
    fig = px.histogram(filtro, x='Km por Litro', y='Veículo/Equip.', color='Data Req.', hover_data=['Km por Litro'], title='Análise 6: Performance por Veículo')
    return fig
def analise7(filtro):
    fig = px.histogram(filtro, x='Vlr. Total', y='Km por Litro', color='Vlr. Total', hover_data=['Veículo/Equip.'], title='Análise 7: Km/Litro por Vlr Total')
    return fig
def analise8(filtro):
    agrupado = filtro.groupby(['Veículo/Equip.', 'Requisitante']).agg({
        'Data Req.': 'max',
        'PLACA/': 'first',
        'Km por Litro': 'mean',
        'Km Atual': 'max'
    }).reset_index()
    #Selectng Top||Bottom10
    top5 = agrupado.nlargest(5, 'Km por Litro')
    bottom10 = agrupado.nsmallest(10, 'Km por Litro')
    #Gráfc Top5
    fig_top5 = px.bar(top5, x='Veículo/Equip.', y='Km por Litro', color='Km por Litro', hover_data=['Requisitante'])
    fig_top5.update_layout(title="Veículos/Equipamentos com MAIOR Km por Litro", xaxis_title="Veículo/Equip.", yaxis_title="Km por Litro", xaxis_tickangle=-45)
    #Gráfc Bottom10
    fig_bottom10 = px.bar(bottom10, x='Veículo/Equip.', y='Km por Litro', color='Km por Litro', hover_data=['Requisitante'])
    fig_bottom10.update_layout(title="Veículos/Equipamentos com MENOR Km por Litro", xaxis_title="Veículo/Equip.", yaxis_title="Km por Litro", xaxis_tickangle=-45)
    return fig_top5, fig_bottom10
#Init varibl'fig't/avoid'NameError'
fig = None
#Show gráfc selectd
if analise == 'Análise 1: Diferença de Km(x)':
    fig = analise1(filtro)
elif analise == 'Análise 2: Km por Litro(x)':
    fig = analise2(filtro)
elif analise == 'Análise 3: Horim por Litro(x)':
    fig = analise3(filtro)
elif analise == 'Análise 4: Km/Litro por Data':
    fig = analise4(filtro)
elif analise == 'Análise 5: Performance Requisitante':
    fig = analise5(filtro)
elif analise == 'Análise 6: Performance por Veículo':
    fig = analise6(filtro)
elif analise == 'Análise 7: Km/Litro por Vlr Total':
    fig = analise7(filtro)
elif analise == 'Análise 8: Top5|Bottom10 Km/Litro':
    fig_top5, fig_bottom10 = analise8(filtro)
    st.plotly_chart(fig_top5)
    st.plotly_chart(fig_bottom10)
#Plot gráfc
if fig:
    st.plotly_chart(fig)
#Bottn export data filtrd p/Excel
if st.button('Exportar Dados Filtrados para Excel', key='export_button'):
    with pd.ExcelWriter('dados_filtrados.xlsx', engine='openpyxl') as writer:
        filtro.to_excel(writer, index=False, sheet_name='Dados Filtrados')
    st.write('Dados exportados para Excel com sucesso!')
    #Link download-Excel-file
    with open('dados_filtrados.xlsx', 'rb') as f:
        st.download_button('Baixar Dados Filtrados', f, file_name='dados_filtrados.xlsx', key='download_button')
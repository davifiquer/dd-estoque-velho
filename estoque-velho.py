# -*- coding: utf-8 -*-
"""
Created on Wed May  3 16:07:16 2023

@author: davi.fiquer
"""

#Carregando as bibliotecas necessárias
import pandas as pd
import glob
from datetime import date

#Configuração PANDAS
pd.set_option('display.float_format', lambda x: '%.3f' % x)
pd.options.display.float_format = '{:.2f}'.format

#Criando as colunas
colunas = ['gerencia', 'comprador', 'fornecedor', 'produto', 'filial', 
           'entradamenor7dias', 'entradamenor28dias', 'entradamenor56dias',
           'entradaentre2a6meses', 'entradamaior6meses', 'entradaacima2meses',
           'estoquetotalpv', '%1', '%2', '%3', '%4', '%5']

#Guardar os dataframes aqui
array_df = []

# Função para ler o arquivo exportado da Ellévti
def gerar_dataframe(arquivo, situacao):
    #Criando o dataframe principal
    df_create = pd.read_excel(arquivo)
    df_create.columns = colunas
    df_create = df_create.drop([0,1])
    if situacao == 'linha':
        df_create['status'] = 'Linha'
    else:
        df_create['status'] = 'Fora Linha'
    df_create.drop(['%1', '%2', '%3', '%4', '%5'], axis=1,
                   inplace=True)
    df_create['data'] = date.today().strftime('%d/%m/%Y')
    
    return df_create


#Base de dados de Produtos em Linha
folder_path = 'leitura-arquivos\linha'
#Base de dados de Produtos Fora de Linha
folder_path_fora_linha = 'leitura-arquivos\\fora'
#Lendo todos os arquivos em linha e fora de linha para realizar o looping
file_list = glob.glob(folder_path + "/*.xlsx")
file_list_fora = glob.glob(folder_path_fora_linha + "/*.xlsx")

#Criando os dataframes em linha
print('Lendo os arquivos em linha...')
for i in range(0, len(file_list)):
    print('Gerando o arquivo número: ' + str(i+1))
    #Criando o dataframe
    dataframe = gerar_dataframe(file_list[i], 'linha')
    #Gravando o dataframe no Array de dataframes
    array_df.append(dataframe)
print('Arquivos em linha concluído!')

#Criando os dataframes fora de linha
print('Lendo os arquivos fora de  linha...')
for i in range(0, len(file_list_fora)):
    print('Gerando o arquivo número: ' + str(i+1))
    #Criando o dataframe
    dataframe = gerar_dataframe(file_list_fora[i], 'fora')
    #Gravando o dataframe no Array de dataframes
    array_df.append(dataframe)
print('Arquivos fora de linha concluído!')

df = pd.concat(array_df, axis=0)

#Selecionando apenas as colunas que eu quero
df = df[['gerencia', 'comprador', 'fornecedor', 'produto', 'filial', 
           'entradamenor7dias', 'entradamenor28dias', 'entradamenor56dias',
           'entradaentre2a6meses', 'entradamaior6meses', 'entradaacima2meses',
           'estoquetotalpv', 'status', 'data']]

print('Base gerada com sucesso!')

#Path onde vai salvar os arquivos
path_to_save = 'G:\\Davi Fiquer\\Relatório - Estoque Velho\\Datenbank'

#Filtrando o estoque velho
df = df[df['entradaacima2meses'] > 0]
#Data de hoje para ficar no nome do arquivo
data_hoje = date.today().strftime('%d/%m/%Y')
data_hoje = data_hoje.replace('/', '.')

path_completo = path_to_save + '\\' + data_hoje + '.xlsx'

#Exportando para Excel
df.to_excel(path_completo, sheet_name='base', index=False,
            encoding='ISO-8859-1')


import pandas as pd 
import numpy as np
from datetime import datetime 

arquivo = pd.ExcelFile("/Users/marciojunior/Documents/telecine/z-rz/vendas-combustiveis-m3.xls")

df_tb1 = pd.read_excel(arquivo, sheet_name='tabela1') #Venda de Combustíveis Derivados de Petróleo por UF e Produto
df_tb2 = pd.read_excel(arquivo, sheet_name='tabela2') #Venda de Diesel por UF e Tipo

#renomeado as colunas 
df_rname_tb1 = df_tb1.rename(columns={'COMBUSTÍVEL':'PRODUCT', 'ESTADO':'UF','UNIDADE':'UNIT','Jan':'2003-01', 'Fev':'2003-02','Mar':'2003-03','Abr':'2003-04','Mai':'2003-05','Jun':'2003-06','Jul':'2003-07','Ago':'2003-08','Set':'2003-09','Out':'2003-10','Nov':'2003-11','Dez':'2003-12'})
df_rname_tb2 = df_tb2.rename(columns={'COMBUSTÍVEL':'PRODUCT', 'ESTADO':'UF','UNIDADE':'UNIT','Jan':'2003-01', 'Fev':'2003-02','Mar':'2003-03','Abr':'2003-04','Mai':'2003-05','Jun':'2003-06','Jul':'2003-07','Ago':'2003-08','Set':'2003-09','Out':'2003-10','Nov':'2003-11','Dez':'2003-12'})

#agrupa e somas as colunas 
soma_tb1 = df_rname_tb1.groupby(['PRODUCT','ANO', 'UF', 'UNIT']).sum()
soma_tb2 = df_rname_tb2.groupby(['PRODUCT','ANO', 'UF', 'UNIT']).sum()

#transforma as colunas em linha
pivot_tb1 = pd.melt(soma_tb1.reset_index(), id_vars=['PRODUCT', 'ANO', 'UF', 'TOTAL', 'UNIT'], 
var_name='YEAR_MONTH', value_name='VOLUME')

pivot_tb2 = pd.melt(soma_tb2.reset_index(), id_vars=['PRODUCT', 'ANO', 'UF', 'TOTAL', 'UNIT'], 
var_name='YEAR_MONTH', value_name='VOLUME')

#criando um DF e sequenciando as colunas conforme teste e criando indice padrao
df_result_final_tb1 = pd.DataFrame(data=pivot_tb1, columns=['YEAR_MONTH','UF','PRODUCT','UNIT','VOLUME'])
df_result_final_tb2 = pd.DataFrame(data=pivot_tb2, columns=['YEAR_MONTH','UF','PRODUCT','UNIT','VOLUME'])

#trabalhando os tipos
df_result_final_tb1.astype({
    'YEAR_MONTH':'datetime64',
    'UF':'string',
    'PRODUCT':'string',
    'UNIT':'string',
    'VOLUME':'double'
})

df_result_final_tb2.astype({
    'YEAR_MONTH':'datetime64',
    'UF':'string',
    'PRODUCT':'string',
    'UNIT':'string',
    'VOLUME':'double'
})

final_tb1 = df_result_final_tb1.to_excel('/Users/marciojunior/Documents/telecine/z-rz/output_sales_of_oil_by_uf_and_product.xls')
final_tb2 = df_result_final_tb2.to_excel('/Users/marciojunior/Documents/telecine/z-rz/output_sales_of_diesel_by_uf_and_product.xls')
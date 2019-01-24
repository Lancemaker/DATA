# Import `os` 
import os
import pandas as pd
import xlrd
import xlsxwriter
import numpy as np
import lib.Carro

# Variaves das pastas e arquivos
raizArquivo = os.getcwd()+"/"
print(raizArquivo)
planilhasArquivo = raizArquivo+"planilhas/"
saida = raizArquivo+"saida/"
diagramar = "Planilha_Diagramar_Completa_v03.xlsx"
completa = "Planilha AUTOMÓVEIS REV_240918_v2.xlsx"
basica= "basica.xlsx"


#Abrindo Arquivos excel como dataframes : 
xl = pd.ExcelFile(planilhasArquivo+basica)
count=0
df1=[]
for name in xl.sheet_names:    
    df1.append(xl.parse(name))
planilha = df1[4]

#variaveis de escrita 
#writer = pd.ExcelWriter(saida+'example.xlsx', engine='xlsxwriter')


#para escrever o arquivo:
#df1.to_excel(writer, 'CONSOLIDADA')
#writer.save()

# Modelo final da planilha:
#MONTADORA MODELO TIPO ANO CAPACIDADE(L) RECOMENDACAO_CASTROL

print (planilha)
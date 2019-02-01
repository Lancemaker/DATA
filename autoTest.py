import os
from openpyxl import load_workbook
from openpyxl import Workbook
import openpyxl
from lib.Carro import Carro
import re


raizArquivo = os.getcwd()+"/"
print(raizArquivo)
planilhasArquivo = raizArquivo+"planilhas/"
saida = raizArquivo+"saida/"
diagramar = "Planilha_Diagramar_Completa_v03.xlsx"
completa = "Planilha AUTOMÓVEIS REV_240918_v2.xlsx"
basica= "planilha.xlsx"

workbook= load_workbook(planilhasArquivo+basica)

base=workbook.get_sheet_by_name('Planilha1')
planilha=workbook.get_sheet_by_name('Filtro_5')
test=workbook.create_sheet("Testlog")
teste=[]
hit=[]
error=[]
for num in range(2,547):
    carroTest=Carro(planilha['A'+str(num)].value,planilha['B'+str(num)].value,planilha['C'+str(num)].value.split(", "),planilha['D'+str(num)].value.split(" "),planilha['E'+str(num)].value.split(", "),planilha['F'+str(num)].value)
    teste.append(carroTest)
for num in range(2,7833):
    carro = Carro(planilha['A'+str(num)].value,planilha['B'+str(num)].value,planilha['C'+str(num)].value,planilha['D'+str(num)].value,planilha['E'+str(num)].value,planilha['F'+str(num)].value)  
    print(num)
    catch = re.findall(r"[-+]?\d*\.\d", carro.modelo)
    if(len(catch)>0):
        carro.tipo = catch[0] 
        carro.modelo = carro.modelo.replace(carro.tipo+" ","")
    else:
        carro.tipo = "NAO ENCONTRADO" 
    carro.capacidade.replace(",",".")
    for t in teste:        
        if(carro.montadora in t.montadora):
            if(t.modelo in carro.modelo):
                if(carro.tipo in t.tipo):
                    print(carro.ano[0],t.ano[0][0])
                    '''                 
                    if(carro.ano[0] > t.ano[0])or(carro.ano[0]<t.ano[2]):                        
                        if(carro.capacidade in t.capacidade):
                            if(t.recomendacao in carro.recomendacao):
                                hit.append(carro)'''
    if(carro not in hit):
        error.append(str(carro.Show())+"na linha :"+str(num))
for i in error:
    print(i)
  

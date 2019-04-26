"""
  Projeto: E-Social Claudino S/A

  Extrair dados de planilhas excel (E-Social/MasterSAF) e
  gerar arquivos no format CSV.

  Francisco Filho - 10.04.2019 - Teresina

  Contribuicao de Edytarcio na identificacao de valores inteiros/decimais

  Importando Base Calculo Descontos e Deducoes 17.04.2019

  OBS.: editor utilizado [ gvim ] vim grafico

  Executar: python Esocial2csv

"""
from xlrd import open_workbook
import csv
import sys
from sys import argv
import os.path
 
def isFile(fileName):
    if(not os.path.isfile(fileName)):
      raise ValueError("Informe Documento Valido")

def search (lista, valor):
    return [(lista.index(x), x.index(valor)) for x in lista if valor in x]


def processa():
    wb = open_workbook(sys.argv[1])

    seqarq  = 0
    valor11 = 0
    valor31 = 0
    compet  = " "
    cnpj    = " "

    for i in range(0, wb.nsheets):
        nomearq = "arquivo"
        sheet = wb.sheet_by_index(i)
        print(sheet.name)
        if sheet.name == 'Processos judiciais':
           nomearq = 'ProcessoJud'
        elif sheet.name == 'Cálculo previdenciário':
           nomearq = 'CalculoPrev'
        elif sheet.name == 'B. cálc., descontos e deduções':
           nomearq ='BCalcDescDedu' + cnpj
        elif sheet.name == 'Contrib. devidas a outras ent.':
           nomearq = 'ContrDevOutras'
        else:
           nomearq = 'arquivo' + str(seqarq)   # Outras pastas excel...
    
        seqarq += 1
        #with open("data/%s.csv" %(sheet.name.replace(" ","")), "w") as file:
        with open("data/%s.csv" %(nomearq), "w") as file:
            writer = csv.writer(file, delimiter = "|")
            print(sheet, sheet.name, sheet.ncols, sheet.nrows)
            if nomearq == 'BCalcDescDedu':  # forma cabecalho 
               header = ["COMPET","CPF","Matricula","IDValor","Valor-11","Valor-21","Valor-31"]
            else:
               header = [cell.value for cell in sheet.row(0)]
    
            writer.writerow(header)  # Grava cabecalho
    
            for row_idx in range(0, sheet.nrows):
                row = [int(cell.value) if isinstance(cell.value, float) and float(cell.value).is_integer() else cell.value
                       for cell in sheet.row(row_idx)]
    
                if 'CNPJ' in row[0]: # Raiz CNPJ
                   cnpj = row[4].replace(".","")
    
                lista = []
                if 'BCalcDescDedu' in nomearq:
                   if 'Período de apuração:' in row[0]: # Competencia
                       compet = row[1][3:] + row[1][:2]
    
                   if 'ID' in row[0]:
                      if row[10] == 21: # finaliza juncao de linhas
                         lista = [compet,row[2],row[6],row[10],valor11,row[11],valor31]
                         valor11 = 0
                         valor31 = 0
                         writer.writerow(lista) # grava registro
                      if row[10] == 11:
                         valor11 = row[11]
                      if row[10] == 31:
                         valor31 = row[11]
                else:
                   writer.writerow(row) # grava outros arquivos (pastas excel--csv)

fileName = " "
if __name__=="__main__":
  try:
    fileName=sys.argv[1]
    isFile(fileName)
    pass
  except Exception as e:
    print("Nome do Documento Nao Informado")
    raise

processa()

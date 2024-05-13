import xlrd
from datetime import datetime, timedelta
import locale
import pandas as pd

book = xlrd.open_workbook("venda.xls")
sh = book.sheet_by_index(0)

locale.setlocale(locale.LC_ALL, 'en_US.utf8')

valoresComDias = []
temp = [0, 0]
pagamentos = [0, 1, 2, 3, 4]
bandeiras = ["Visa", "Mastercard"]
tiposVenda = ["Crédito à vista", "Crédito conversor moedas"]

prev = float('-inf')
chave = 0
isConversor = False

for rownum in reversed(range(5, sh.nrows)):
    d1 = datetime.strptime(sh.cell(rownum, 11).value, "%d/%m/%Y")
    d2 = datetime.strptime(sh.cell(rownum, 1).value, "%d/%m/%Y")

    bandeira = sh.cell(rownum, 3).value
    tipoVenda = sh.cell(rownum, 4).value

    diferencaDias = (abs((d2 - d1).days))

    if(d1.day > prev):
        prev = d1.day
        somaValor = 0

    if((diferencaDias < 8) 
       & (tipoVenda in tiposVenda)
       & (bandeira in bandeiras)):
            valoresComDias.append([sh.cell(rownum, 1).value, sh.cell(rownum, 3).value, sh.cell(rownum, 4).value, sh.cell(rownum, 11).value, (locale.atof(sh.cell(rownum, 7).value.strip("R$"))/100)])

df = pd.DataFrame(data=valoresComDias)
df.to_excel('resumoA.xlsx')

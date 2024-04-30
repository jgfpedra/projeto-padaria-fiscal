import xlrd
from datetime import datetime
import locale
import pandas as pd

book = xlrd.open_workbook("venda.xls")
sh = book.sheet_by_index(0)

locale.setlocale(locale.LC_ALL, 'en_US.utf8')

dicionario = {}

somaValor = 0 

for rownum in reversed(range(5, sh.nrows)):
    d1 = datetime.strptime(sh.cell(rownum, 11).value, "%d/%m/%Y")
    d2 = datetime.strptime(sh.cell(rownum, 1).value, "%d/%m/%Y")
    if((abs((d1 - d2).days) == 5)):
        somaValor += (locale.atof(sh.cell(rownum, 7).value.strip("R$"))/100)
        dicionario[sh.cell(rownum, 1).value] = [sh.cell(rownum, 3).value, sh.cell(rownum, 11).value, somaValor]

df = pd.DataFrame(data=dicionario)
df = (df.T)
print(df)
df.to_excel('resumo.xlsx')

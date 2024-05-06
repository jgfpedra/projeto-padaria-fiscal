import xlrd
from datetime import datetime, timedelta
import locale
import pandas as pd

def daterange(start_date, end_date):
    for n in range(int((end_date - start_date).days)):
        yield start_date + timedelta(n)

book = xlrd.open_workbook("venda.xls")
sh = book.sheet_by_index(0)

locale.setlocale(locale.LC_ALL, 'en_US.utf8')

valoresComDias = []
diaFeriados = []

somaValor = 0 
quantidadeFeriado = 0

finalSemana = [5, 6]

possuiFeriado = input("Houve algum feriado entre o periodo de dias da tabela? (s | n)\nEntrada: ")

if possuiFeriado == 's':
    quantidadeFeriado = int(input("Quantos dias foram? (Somente numero inteiro)\n"))
    while(quantidadeFeriado != 0):
        diaFeriados.append(datetime.strptime(input("Qual o dia do feriado? (XX/XX/XXXX): "), "%d/%m/%Y"))
        quantidadeFeriado -= 1


for rownum in reversed(range(5, sh.nrows)):
    d1 = datetime.strptime(sh.cell(rownum, 11).value, "%d/%m/%Y")
    d2 = datetime.strptime(sh.cell(rownum, 1).value, "%d/%m/%Y")

    diasTotais = (abs((d1 - d2).days))

    for date in daterange(d2, d1):
        if date.weekday() in finalSemana:
            diasTotais -= 1
        if date in diaFeriados:
            diasTotais -= 1

    if possuiFeriado == 's':
        diasTotais = diasTotais - quantidadeFeriado

    if(diasTotais == 5):
        somaValor += (locale.atof(sh.cell(rownum, 7).value.strip("R$"))/100)
        valoresComDias.append([sh.cell(rownum, 1).value, sh.cell(rownum, 3).value, sh.cell(rownum, 4).value, sh.cell(rownum, 11).value, somaValor])

df = pd.DataFrame(data=valoresComDias)
df.to_excel('resumo.xlsx')

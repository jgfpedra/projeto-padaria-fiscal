import xlrd
from datetime import datetime, timedelta
import locale
import pandas as pd

book = xlrd.open_workbook("venda.xls")
sh = book.sheet_by_index(0)

locale.setlocale(locale.LC_ALL, 'en_US.utf8')

valoresComDias = []
temp = [0, 0, 0, 0, 0, 0, 0, 0]
bandeiras = ["Visa", "Mastercard"]
tiposVenda = ["Crédito à vista", "Crédito conversor moedas"]

prev = float('-inf')
maior = 0

for rownum in reversed(range(5, sh.nrows)):
    d1 = datetime.strptime(sh.cell(rownum, 11).value, "%d/%m/%Y")
    d2 = datetime.strptime(sh.cell(rownum, 1).value, "%d/%m/%Y")

    diferencaDias = (abs((d2 - d1).days))

    if(diferencaDias < 8):
        bandeira = sh.cell(rownum, 3).value
        tipoVenda = sh.cell(rownum, 4).value
        dataPagamento = sh.cell(rownum, 11).value
        valorVenda = (locale.atof(sh.cell(rownum, 7).value.strip("R$"))/100)

        if((bandeira == bandeiras[0]) & (tipoVenda == tiposVenda[0])):
            temp[0] += valorVenda
            temp[1] = dataPagamento
        elif((bandeira == bandeiras[0]) & (tipoVenda == tiposVenda[1])):
            temp[2] += valorVenda
            temp[3] = dataPagamento
        elif((bandeira == bandeiras[1]) & (tipoVenda == tiposVenda[0])):
            temp[4] += valorVenda
            temp[5] = dataPagamento
        elif((bandeira == bandeiras[1]) & (tipoVenda == tiposVenda[1])):
            temp[6] += valorVenda
            temp[7] = dataPagamento


        if((d2.day > prev) | (d2.day >= 30)):

            if((d2.day >= 30) & (((d2 + timedelta(1)).day + 1) != 31)):
                maior = 1

            if(temp[0] != 0):
                valoresComDias.append([(d2.date() - timedelta(1 - maior)).strftime("%d/%m/%Y"), temp[1], bandeiras[0], tiposVenda[0], temp[0]])
            if(temp[2] != 0):
                valoresComDias.append([(d2.date() - timedelta(1 - maior)).strftime("%d/%m/%Y"), temp[3], bandeiras[0], tiposVenda[1], temp[2]])
            if(temp[4] != 0):
                valoresComDias.append([(d2.date() - timedelta(1 - maior)).strftime("%d/%m/%Y"), temp[5], bandeiras[1], tiposVenda[0], temp[4]])
            if(temp[6] != 0):
                valoresComDias.append([(d2.date() - timedelta(1 - maior)).strftime("%d/%m/%Y"), temp[7], bandeiras[1], tiposVenda[1], temp[6]])
            prev = d2.day
            temp = [0, 0, 0, 0, 0, 0, 0, 0]


df = pd.DataFrame(data=valoresComDias)
df.to_excel('resumo.xlsx')

from datetime import datetime, timedelta
import locale
import pandas as pd

df = pd.read_excel("venda.xls", header=4)

locale.setlocale(locale.LC_ALL, 'en_US.utf8')

bandeiras = ["Visa", "Mastercard"]
tiposVenda = ["Crédito à vista", "Crédito conversor moedas"]

prev = float('-inf')
maior = 0

cols = [1, 3, 4, 7, 11]
valores = df[df.columns[cols]]

valores = valores[(valores['Bandeira'].isin(bandeiras)) &
             (valores['Forma de pagamento'].isin(tiposVenda)) &
             ((valores['Previsão de pagamento'].apply(lambda x: datetime.strptime(x, "%d/%m/%Y")) -
               valores['Data da autorização da venda'].apply(lambda x: datetime.strptime(x, "%d/%m/%Y"))).apply(lambda x: x.days) < 8)]

valores['Valor da Venda'] = valores['Valor da venda'].apply(lambda x: locale.atof(x.strip("R$"))/100)

valores = valores.groupby(['Data da autorização da venda', 'Bandeira', 'Forma de pagamento', 'Previsão de pagamento'])['Valor da Venda'].sum()


df = pd.DataFrame(data = valores)
df.to_excel('resumoB.xlsx')

'''

Data da autorizacao da venda --> data inicial
Previsão de pagamento --> data final
Bandeira --> bandeiras
Forma de Pagamento --> tiposVenda
Valor da venda
'''

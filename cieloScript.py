from datetime import datetime, timedelta
import locale
import pandas as pd

df = pd.read_excel("venda.xlsx", header=9)

locale.setlocale(locale.LC_ALL, 'en_US.utf8')

bandeiras = ["Visa", "Mastercard"]
tiposVenda = ["Crédito à vista", "Crédito conversor moedas"]

cols = [0, 1, 3, 4, 6, 11]
valores = df[df.columns[cols]]

valores = valores[(valores['Bandeira'].isin(bandeiras)) &
             (valores['Forma de pagamento'].isin(tiposVenda)) &
             ((valores['Data prevista do pagamento'].apply(lambda x: datetime.strptime(x, "%d/%m/%Y")) -
               valores['Data da venda'].apply(lambda x: datetime.strptime(x, "%d/%m/%Y"))).apply(lambda x: x.days) < 8)]

valores = valores.groupby(['Data da venda', 'Bandeira', 'Forma de pagamento', 'Data prevista do pagamento'])['Valor líquido'].sum()

df = pd.DataFrame(data = valores)
df.to_excel('resumo.xlsx')

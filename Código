!pip install yfinance==0.1.74
!pip install mplcyberpunk
!pip install pywin32
import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
import win32com.client as win32

codigos_de_negociacao = ["^BVSP", "BRL=X"]

hoje = datetime.datetime.now()
um_ano_atras = hoje - datetime.timedelta(days = 365)

dados_mercado = yf.download(codigos_de_negociacao, um_ano_atras, hoje)

dados_fechamento = dados_mercado['Adj Close']

dados_fechamento.columns = ['dolar', 'ibovespa']

dados_fechamento = dados_fechamento.dropna()

dados_anuais = dados_fechamento.resample("Y").last()

dados_mensais = dados_fechamento.resample("M").last()

retorno_anual = dados_anuais.pct_change().dropna()

retorno_mensal = dados_mensais.pct_change().dropna()

retorno_diario = dados_fechamento.pct_change().dropna()

retorno_diario_dolar = retorno_diario.iloc[-1, 0]
retorno_diario_ibov = retorno_diario.iloc[-1, 1]

retorno_mensal_dolar = retorno_mensal.iloc[-1, 0]
retorno_mensal_ibov = retorno_mensal.iloc[-1, 1]

retorno_anual_dolar = retorno_anual.iloc[-1, 0]
retorno_anual_ibov = retorno_anual.iloc[-1, 1]

retorno_diario_dolar = round((retorno_diario_dolar * 100), 2)

retorno_diario_ibov = round((retorno_diario_ibov * 100), 2)

retorno_mensal_dolar = round((retorno_mensal_dolar * 100), 2)

retorno_mensal_ibov = round((retorno_mensal_ibov* 100), 2)

retorno_anual_dolar = round((retorno_anual_dolar * 100), 2)

retorno_anual_ibov = round((retorno_anual_ibov * 100), 2)

plt.style.use("cyberpunk")

dados_fechamento.plot(y = "ibovespa", use_index = True, legend = False)
plt.title("Ibovespa")
plt.savefig('Ibovespa.png', dpi = 300)
plt.show()

plt.style.use("cyberpunk")

dados_fechamento.plot(y = "dolar", use_index = True, legend = False)
retorno_anual = dados_anuais.pct_change().dropna()


retorno_mensal = dados_mensais.pct_change().dropna()

retorno_diario= dados_fechamento.pct_change().dropna()plt.title("Dolar")
plt.savefig('dolar.png', dpi = 300)
plt.show()

outlook = win32.Dispatch("outlook.application")

email = outlook.CreateItem(0)

email.To = "tassandro412@gmail.com"
email.Subject = "Relatório Diário"
email.Body = f'''Fala, chefe, receba o relatório diário:

Bolsa:

No ano o Ibovespa tá com uma rentabilidade de {retorno_anual_ibov}%,
enquanto no mês ela é de {retorno_mensal_ibov}%.

No último dia útil, o fechamento do Ibovespa foi de {retorno_diario_ibov}%.

Dólar:

No ano o Dólar tá com uma rentabilidade de {retorno_anual_dolar}%,
enquanto no mês ela é de {retorno_mensal_dolar}%.

No último dia útil, o fechamento do Dólar foi de {retorno_diario_dolar}%.

Receba!

'''

anexo_ibovespa = r'C:\Users\Tassandro\ibovespa.png'
anexo_dolar = r'C:\Users\Tassandro\dolar.png'

email.Attachments.Add(anexo_ibovespa)
email.Attachments.Add(anexo_dolar)

email.Send()

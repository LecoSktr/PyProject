import pandas as pd

# importar a base de dados
tabela_vendas = pd.read_excel ('Vendas.xlsx')


# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# calcular faturamento por loja
faturamento =  tabela_vendas[['ID Loja', 'Valor Final']] .groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendido por loja
Quantidade = tabela_vendas [['ID Loja' , 'Quantidade']] .groupby ('ID Loja').sum()
print(Quantidade)

print('-' * 50)
# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / Quantidade['Quantidade']).to_frame()
print(ticket_medio)

# enviar um email com relatório
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail= outlook.CreateItem(0)
mail.To = 'insert your addres'
mail.Subject = 'Tabela das Lojas'
mail.HTMLBody = f'''
<p> Prezados </p>


<p> Segue o Relátorio das media das vendas por cada loja </p>

<p> Faturamento </p>
{faturamento.to_html()}

<p> Quantidade vendida por loja </p>
{Quantidade.to_html()}

<p> Ticket Medio Por Produto em cada loja </p>
{ticket_medio.to_html()}

<p> Qualquer duvida estou a disposição </p>
'''

mail.Send()

print('email enviado')
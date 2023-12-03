import pandas as pd
import win32com.client as win32

# Importar a base de Dados
table_vends = pd.read_excel('vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(table_vends)
# Faturamento por Loja
faturamento = table_vends[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
qtd_produtos_loja = table_vends[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_produtos_loja)

print('-' * 50)
# Ticket Médio de produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_produtos_loja['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Enviar e-mail com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'daniellaavila@outlook.com.br'
mail.Subject = 'Relatório com Python 1.0'
mail.HTMLBody = f'''
<p>Prezados,</p>


<p>Segue o relatório de Vendas por cada loja:</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}


<p>Quantidade Vendida:</p>
{qtd_produtos_loja.to_html()}


<p>Ticket Médio dos produtos em cada loja:</P>
{ticket_medio.to_html()}


<p>Qualquer dúvida, estou à disposição,</p> 

<p>Att..</p>
<p>Rodrigo Lopes Emidio da Costa</p>

'''
mail.Send()

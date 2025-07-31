
# Lógica 
# [x] 1.Importar a base de dados
# [x] 2.Visualizar a base dados
# [x] 3.Faturamento por Loja
# [x] 4.Quantidade de Produtos vendidos por Loja
# [x] 5.Ticket Médio por Produto em cada Loja 
#     -> faturamento/quantidade de produtos vendidos
# [x] 6.Enviar um Email com o Relatório

# importardo a biblioteca pandas
import pandas as pd
# importando a biblioteca win32com para enviar email
import win32com.client as win32

#   1.Importando a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# para visualizar todas as colunas
pd.set_option('display.max_columns', None)

divisor = '=' * 50  
#   2.Visualizar a base de dados
    #   primeiro método  
    #   print(tabela_vendas[['ID Loja','Valor Final']])

    #   segundo método
    #   print(tabela_vendas.groupby('ID Loja').sum())

#   3.Faturamento por Loja
#   Agrupando os dados por 'ID Loja' e somando o 'Valor Final
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#print(faturamento)
print(divisor)
#   4.Quantidade de Produtos vendidos por Loja 
#   Agrupando os dados por 'ID Loja' e somando a 'Quantidade' de produtos vendidos
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print(divisor)
#   5.Ticket Médio por Produto em cada Loja
#   Calculando o ticket médio dividindo o faturamento pela quantidade de produtos vendidos
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

ticket_medio.rename(columns={0: 'Ticket Médio'}) # renomeando a coluna para 'Ticket Médio'

#   6.Enviando um Email com o Relatório
#   Integrando o email com o pywin32

outlook = win32.Dispatch('outlook.application') 
email = outlook.CreateItem(0) 
email.To = 'email@email.com' 
email.Subject = 'Relatório de Vendas'
email.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja:</p>

<p>Faturamento por Loja: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2F}'.format})}

<p>Quantidade de Produtos Vendidos por Loja:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produto em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2F}'.format})}  

<p>Qualquer dúvida, estou à disposição. </p>

<p>Atenciosamente,</p>
<p>Seu Nome</p>
'''
email.send()

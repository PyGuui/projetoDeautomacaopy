import pandas as pd
import win32com.client as win32

#Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# Visualizar a base de dados
pd.set_option('display.max_columns', None)

#print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()



# quantidade de produtos vendidos por loja
Quantidade_Produtos = tabela_vendas [["ID Loja","Quantidade"]].groupby("ID Loja").sum()


# ticket médio por produto em cada loja
ticket_medio = (faturamento["Valor Final"] / Quantidade_Produtos["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns ={0: "Ticket Médio"})


# enviar um email com relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'guilhermesantospereira01@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters ={"Valor Final": "R${:,.2f}".format})} 

<p>Quantidade Vendida:</p>  
{Quantidade_Produtos.to_html()}


<p>Ticket Médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters ={"Ticket Médio": "R${:,.2f}".format})}

<p>Att...</p>

<p>Guilherme Santos</p>

'''

mail.Send()

print("Email enviado")
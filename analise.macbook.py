import pandas as pd

# Importar a base de dados
tabela = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print('=' * 60)
print('Tabela de Analise')
print('=' * 60)
print(tabela)
print('\n')


# faturamento por loja
faturamento = tabela[["ID Loja", 'Valor Final']].groupby('ID Loja').sum()
print('=' * 60)
print(f'Faturamento Total das Lojas:')
print('=' * 60)
print(faturamento)
print('\n')


# quantidade de produtos vendidos por loja
quantidade = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print('=' * 60)
print(f'Quantidade total por loja')
print('=' * 60)
print(quantidade)
print('\n')


# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print('*' * 60)
print(f'Ticket Medio por Loja')
print('*' * 60)
print(ticket_medio)
print('\n')

# Enviando relatório por email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Verificando se existe alguma oferta dentro da tabela de ofertas
# Configurações do servidor de email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configurações do servidor de email
servidor_smtp = 'smtp.office365.com'
porta = 587
seu_email = 'thiarly.cavalcante@live.com'
sua_senha = '@@security$$'



# Criando o objeto MIMEMultipart para enviar o e-mail

msg = MIMEMultipart()
msg['From'] = 'Thiarly Cavalcante <thiarly.cavalcante@live.com>'
destino = msg['To'] = 'thiarly.cavalcante@gmail.com'
print('-' * 65)
print(f'E-mail será enviado para: {msg["To"]}')
print('-' * 65)
print('\n')
msg['Subject'] = 'Relatório de Vendas por Loja'

# Adicionando o conteúdo do e-mail
# texto_html = f"""
# <!DOCTYPE html>
# <html>
# <head>
# <style>
#     table {{
#         border-collapse: collapse;
#         white-space: nowrap;
#     }}
#     table, th, td {{
#         border: 1px solid black;
#     }}
# </style>
# </head>
# <body>
# <p>Prezados,</p>
# <p>Segue o Relatório de Vendas por cada Loja.</p>
# <p>Faturamento:</p>
# {faturamento.to_html(formatters={'Valor Final': lambda x: 'R$ {:,.2f}'.format(x).replace(",", "|").replace(".", ",").replace("|", ".")}, border=1)}
# <p>Quantidade Vendida:</p>
# {quantidade.to_html(border=1)}
# <p>Ticket Médio dos Produtos em cada Loja:</p>
# {ticket_medio.to_html(formatters={'Ticket Medio': lambda x: 'R$ {:,.2f}'.format(x).replace(",", "|").replace(".", ",").replace("|", ".")}, border=1)}
# <p>Qualquer dúvida estou à disposição.</p>
# <p>Att.,</p>
# <p>Lira</p>
# </body>
# </html>
# """
texto_html = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': lambda x: 'R${:,.2f}'.format(x).replace(",", "|",).replace(".", ",").replace("|", ".")}, border=1)}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': lambda x: 'R${:,.2f}'.format(x).replace(",", "|",).replace(".", ",").replace("|", ".")}, border=1)}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Thiarly</p>
'''


msg.attach(MIMEText(texto_html, 'html'))


# Enviando o e-mail
with smtplib.SMTP(servidor_smtp, porta) as server:
    server.starttls()
    server.login(seu_email, sua_senha)
    server.sendmail(seu_email, 'thiarly.cavalcante@gmail.com', msg.as_string())
    print('*' * 65)
    print(f'Email enviado com sucesso para: {destino}')
    print('*' * 65)



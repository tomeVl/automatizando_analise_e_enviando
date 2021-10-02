from selenium import webdriver
import zipfile
import pandas as pd
import time
import win32com.client as win32
import os

driver = webdriver.Chrome(executable_path=r'./chromedriver.exe')
driver.get('https://www.kaggle.com/sakshigoyal7/credit-card-customers')
time.sleep(5)

driver.find_elements_by_css_selector("a.sc-EZqKI")[0].click()  # Botão de DOWNLOAD
time.sleep(2)

driver.find_elements_by_css_selector("a.sc-bwcZwS")[1].click()  # CLICANDO NO BOTÃO DE GMAIL
time.sleep(1)

driver.find_elements_by_css_selector("input.mdc-text-field__input")[0].send_keys("rubenstome15@gmail.com")  # preenchendo o gmail
driver.find_elements_by_css_selector("input.mdc-text-field__input")[1].send_keys(os.environ['SENHA'])  # preenchendo a senha
driver.find_elements_by_css_selector("button.sc-jXcxbT")[0].click()
time.sleep(3)

driver.find_elements_by_css_selector("a.sc-EZqKI")[0].click()  # botao de download
time.sleep(10)

# tirando o arquivo do zip
with zipfile.ZipFile(r'C:\Users\55119\Downloads\archive.zip', 'r') as zip_ref:
    zip_ref.extractall(r'C:\Users\55119\Downloads')

# lendo o arquivo
cliente_df = pd.read_csv(r'C:\Users\55119\Downloads\BankChurners.csv')

# Quantos clientes Ativos e Quantos Clientes em Disputa Temos?
resumo_status = cliente_df.groupby('Attrition_Flag')['Attrition_Flag'].count()

# Dos Clientes Ativos, como está a distribuição por tipo de cartão
clientes_ativo = cliente_df.loc[cliente_df['Attrition_Flag'] == 'Existing Customer', ['Attrition_Flag', 'Card_Category']]
resumo_cartoes = clientes_ativo.groupby('Card_Category')['Card_Category'].count()
resumo_cartoes.index.names = ['Categoria do card - Existing Customer']

# Qual o tempo médio de permanência e o limite de crédito médio dos nosso clientes (totais e ex-clientes)?
tempo_medio = cliente_df['Months_on_book'].mean()
limite_todos = cliente_df['Credit_Limit'].mean()
ex_cliente = cliente_df.loc[cliente_df['Attrition_Flag'] == 'Attrited Customer', ['Attrition_Flag', 'Credit_Limit']]
limite_excliente = ex_cliente['Credit_Limit'].mean()

# enviando o email


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

mail.To = 'rubenstome15@gmail.com'
mail.Subject = 'Relatório de Clientes - Análise de Attrited Customers'
mail.Body = f'''
Olá Lira, tudo bem?

Conforme solicitado, levantamos os principais indicadores dos nossos clientes para ver o impacto dos Attrited Customers.

Temos atualmente a seguinte divisão da base de clientes:

{resumo_status.to_string()}

Quando analisamos os clientes ativos (Existing Customers) percebemos a seguinte divisão de categorias de cartão:

{resumo_cartoes.to_string()}

Já quanto ao tempo médio de permanência dos clientes temo em média {tempo_medio:.1f} meses.

Por fim, quando comparamos o limite de crédito médio dos clientes atuais e antigos, não temos muita diferença, ficando com {limite_todos:.1f} para os atuais e {limite_excliente:.1f} para os antigos.

Segue em anexo o relatório completo para mais detalhes.

Qualquer dúvida estou à disposição.

Att.,
Tomé
'''

# Anexos (pode colocar quantos quiser):
attachment = r'C:\Users\55119\Downloads\BankChurners.csv'
mail.Attachments.Add(attachment)

mail.Send()

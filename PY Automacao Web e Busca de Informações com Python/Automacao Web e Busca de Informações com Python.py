#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# Passo 1 - entrar na internet 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

navegador = webdriver.Chrome()

# Passo 2 - pegar cotação dólar
# Entrar no Google
navegador.get("https://www.google.com/")
# Pesquisar cotação dolar 
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dólar")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
# Pegar valor
cotacao_dolar = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_dolar)

# Passo 3 - pegar cotação Euro
# Entrar no Google
navegador.get("https://www.google.com/")
# Pesquisar cotação euro 
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
navegador.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
# Pegar valor
cotacao_euro = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
print(cotacao_euro)

# Passo 4 - pegar cotação ouro
# Entrar no site
navegador.get("https://www.melhorcambio.com/ouro-hoje")
# Pegar a cotação ouro
cotacao_ouro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute("value")
cotacao_ouro = cotacao_ouro.replace(",",".")
print(cotacao_ouro)

navegador.quit()


# Passo 5 - Importar e atualizar a base de dados
import pandas as pd

tabela = pd.read_excel("Produtos.xlsx")
## Caso a planilha tenha mais de 1 aba 
# tabela = pd.read_excel("Produtos.xlsx", sheets=numero da aba ou nome da aba)
display(tabela)

# atualizar a cotação
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"]= float(cotacao_ouro)



# atualizar o preço de compra -> preço original * cotação
tabela["Preço Base Reais"] = tabela["Preço Base Original"] * tabela["Cotação"]

# atualizar o preço de venda -> preço de compra * margem
tabela["Preço Final"] = tabela["Preço Base Reais"] * tabela["Margem"]

# Formatando casa decimal
tabela["Preço Final"] = tabela["Preço Final"].map("{:.2f}".format)
tabela["Preço Base Reais"] = tabela["Preço Base Reais"].map("{:.2f}".format)
tabela["Cotação"] = tabela["Cotação"].map("{:.2f}".format)
display(tabela)

# Passo 6 - Exportar a base de dados atualizada
tabela.to_excel("Produtos Novo Final.xlsx", index=False)
# Para substituir basta colocar nome original
# index=False para o python não exportar a coluna de indice


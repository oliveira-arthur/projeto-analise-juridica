#!/usr/bin/env python
# coding: utf-8

# In[12]:


#criar navegador
from selenium import webdriver  #import para trabalhar com o selenio, permite criar o navegador
from selenium.webdriver.chrome.service import Service #vao idenficar a versão do chrome e ira fazer o download automaticamente
from webdriver_manager.chrome import ChromeDriverManager

servico = Service(ChromeDriverManager().install()) #criando o serviço que vai identificar e instalar o webdriver no navegador
navegador = webdriver.Chrome(service=servico) #criando o navegador

# entrar no site para fazer a busca juridica


# In[13]:


import os
import time


caminho = os.getcwd()
arquivo = caminho + r"\index.html"
navegador.get(arquivo)


#se eu quiser entrar em um site diretamente eu nao vou usar o caminho mas sim o site
#navegador.get("http www.site.com")


# In[14]:


import pandas as pd

tabela = pd.read_excel('Processos.xlsx')
display(tabela)


# In[15]:


#colocar o mouse em cima da opção para clicar
from selenium.webdriver.common.by import By  #import parar encontrar os elementos dentro do site
from selenium.webdriver import ActionChains

for linha in tabela.index:
    navegador.get(arquivo)

   
    botao = navegador.find_element(By.CLASS_NAME, 'dropdown-menu')
    ActionChains(navegador).move_to_element(botao).perform()
    
    cidade = tabela.loc[linha, "Cidade"]
    
   
    navegador.find_element(By.PARTIAL_LINK_TEXT, cidade).click()
    
    # mudar para a nova aba
    aba_original = navegador.window_handles[0]
    indice = 1 + linha
    nova_aba = navegador.window_handles[indice]
    
    navegador.switch_to.window(nova_aba)
    
    # preencher o formulário com os dados de busca
    navegador.find_element(By.ID, 'nome').send_keys(tabela.loc[linha, "Nome"])
    navegador.find_element(By.ID, 'advogado').send_keys(tabela.loc[linha, "Advogado"])
    navegador.find_element(By.ID, 'numero').send_keys(tabela.loc[linha, "Processo"])

    # clicar em pesquisar
    navegador.find_element(By.CLASS_NAME, 'registerbtn').click()
    
    # confirmar a pesquisa
    alerta = navegador.switch_to.alert
    alerta.accept()
    
    # esperar o resultado da pesquisa e agir de acordo com o resultado
    while True:
        try:
            alerta = navegador.switch_to.alert
            break
        except:
            time.sleep(1)
    texto_alerta = alerta.text

    if "Processo encontrado com sucesso" in texto_alerta:
        alerta.accept()
        tabela.loc[linha, "Status"] = "Encontrado"
    else:
        tabela.loc[linha, "Status"] = "Não encontrado"
        alerta.accept()
        
        
navegador.quit()


# In[17]:


display(tabela)
tabela.to_excel('Processos Atualizados.xlsx') #criando uma tabela dos processos em excel


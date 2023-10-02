#Automação de encaminhamento de mensagens no whatsapp
#Usando Funcionalidade Nativa
#!Instalar pip install selenium pyperclip webdriver-manager

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

import os
import pyperclip
import time
import pandas as pd

service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
nav = webdriver.Chrome(options=options, service=service)

nav.get("https://web.whatsapp.com")

time.sleep(1)

#Pega a mensagens no Arquivo CSV


#!Verifica se o arquivo está aberto
def openFile():
    #*Repetição enquanto estiver aberto
    isOpen = False
    while True:
        try: 
            os.rename('mensagem.xlsx', 'mensagem2.xlsx')
            os.rename('mensagem2.xlsx', 'mensagem.xlsx')
            isOpen = False
        except OSError:
            print("Arquivo aberto, feche-o")
            isOpen = True
        
        if(not isOpen):
            df = pd.read_excel(r"mensagem.xlsx")

            for idx, mensagem in enumerate(df.values,start = 0):
                
                contatoPlan = mensagem[0]
                mensagemPlan = mensagem[1]
                if(contatoPlan != None and mensagemPlan != None):                    
                    if(len(mensagem) > 2):
                        isSentPlan = mensagem[2]
                            
                    #! Verifica se a mensagem foi enviada
                    if(isSentPlan != "Enviado" or isSentPlan != "Fim"):
                        #Enviar a mensagem para o Meu Grupo para depois encaminhar
                        #Clicar na Lupa escrever "EU" e apertar ENTER
                        nav.find_element('xpath', '//*[@id="side"]/div[1]/div/div/button/div[2]/span').click()

                        nav.find_element('xpath', '//*[@id="side"]/div[1]/div/div/div[1]').send_keys(contatoPlan)

                        nav.find_element('xpath', '//*[@id="side"]/div[1]/div/div/div[1]').send_keys(Keys.ENTER)
                        time.sleep(3)
                        #Escreve-se a mensagem
                        pyperclip.copy(mensagemPlan)
                        nav.find_element('xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.CONTROL + "v")
                        nav.find_element('xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.ENTER)
                        time.sleep(2)
                        print('a')
                        df.loc[idx, 'Enviado'] = 'Enviado'
                        df.to_excel("mensagem.xlsx", index=False)
            time.sleep(30)
        else:
            time.sleep(30)

openFile()

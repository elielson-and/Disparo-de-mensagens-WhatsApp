import time
import xlrd
import pandas as pd
import datetime
from selenium import webdriver

x = pd.read_excel(r"C:/Users/eliel/Desktop/Bot python/bot_python/dados.xlsx")
book = xlrd.open_workbook("C:/Users/eliel/Desktop/Bot python/bot_python/dados.xlsx")
sh = book.sheet_by_index(0)
numero_de_linhas = sh.nrows - 1 

link_p1 = "https://api.whatsapp.com/send?phone=+55"
linkp_2 = "feliz%20aniversário!%20Muita%20saúde,%20paz,%20fé%20e%20esperança%20para%20você%20e%20sua%20família.%20Aproveito,%20também,%20para%20parabenizar%20por%20seu%20empenho%20e%20contribuição%20com%20a%20educação%20do%20Maranhão!%20Abraços,%20Felipe%20Camarão."

class WhatsappBot:
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument('lang=pt-br')
        self.driver = webdriver.Chrome(
            executable_path=r'./chromedriver.exe', chrome_options=options)

    def EnviarMensagens(self):
        self.driver.get('https://send-bot-links.netlify.app/')
        time.sleep(2)
        inicio = True
        while inicio == True:
            time.sleep(2)
            now = datetime.datetime.now()
            dia_hoje = str(now.day) +'.'+ str(now.strftime("%m")+'.'+str(now.strftime("%Y")))

            for index in range(numero_de_linhas):
                if x['Data'] [index] == dia_hoje: #Se a condicao bater, é anv da pessoa
                    nome_pessoa = x['Nome'][index]
                    num_f = x['Whatsapp'][index]
                    numero_pessoa = int(num_f)
                    #print(link_p1 + str(numero_pessoa) + linkp_2)
                    self.link_final = link_p1 +  str(numero_pessoa) + '&text=' + nome_pessoa + ', ' + linkp_2

                    campo_de_texto = self.driver.find_element_by_xpath('//*[@id="link_final"]')
                    campo_de_texto.click()
                    time.sleep(2)
                    campo_de_texto.send_keys(self.link_final)
                    bt_buscar = self.driver.find_element_by_xpath('//*[@id="bt_ex"]')
                    time.sleep(5)
                    bt_buscar.click()

                    bt_iniciar_conversa = self.driver.find_element_by_xpath('//*[@id="action-button"]')
                    time.sleep(5)
                    bt_iniciar_conversa.click()


                    time.sleep(5)

                    bt_wpp = self.driver.find_element_by_xpath('//*[@id="fallback_block"]/div/div/a')
                    time.sleep(5)
                    bt_wpp.click()

                    time.sleep(25)

                    bt_enviar = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button/span')
                    time.sleep(5)
                    bt_enviar.click()
                    
                    time.sleep(5)
                    self.driver.get('https://send-bot-links.netlify.app/')
        now = " "

bot = WhatsappBot()
bot.EnviarMensagens()
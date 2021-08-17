
import time
import xlrd
import pandas as pd
import datetime
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys



x = pd.read_excel(r" ''' passar onde está a planilha com os dados aqui '''")
book = xlrd.open_workbook(" ''' passar onde está a planilha com os dados aqui ''' ")
sh = book.sheet_by_index(0)
numero_de_linhas = sh.nrows - 1 

contatos = 0
for qtd_lines in range(numero_de_linhas):
    contatos = contatos + 1

print(' - INICIANDO OPERAÇÕES, AGUARDE -')
time.sleep(2)
print(' - Verificando diretórios...')  
time.sleep(1)
print(' - Verificando arquivos...')
time.sleep(2)
print(' - Obtendo base de dados...')
time.sleep(2)
print(' - FORAM ENCONTRADAS '+str(contatos)+' PESSOAS')
pyautogui.press("enter")
print(' - Preparando ambiente, por favor aguarde...')
print(' = Tudo Pronto para iniciar!')
time.sleep(2)
pyautogui.press("enter")


print('***********************************************')
print(' A PARTIR DESTE MOMENTO ESTÁ TUDO AUTOMATIZADO')
print(' POR FAVOR, NÃO REALIZE NEHUMA AÇÃO!')
print('***********************************************')
time.sleep(3)


link_p1 = "https://api.whatsapp.com/send?phone=+55"
linkp_2 = "Olá, Feliz aniversário"

class WhatsappBot:
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument('lang=pt-br')
        self.driver = webdriver.Chrome(executable_path=r'./chromedriver.exe', chrome_options=options)
        self.act = ActionChains(self.driver)
    def EnviarMensagens(self):
        self.driver.get('https://sendbot.netlify.app/')
        pyautogui.press("F11")
        time.sleep(2)
        
        repeticoes = 0
        warnning = False
        inicio = True
        while inicio == True:
            time.sleep(2)
            now = datetime.datetime.now()

            for index in range(numero_de_linhas):
                dia_hoje = str(now.strftime("%d") +'.'+ now.strftime("%m"))
                repeticoes = repeticoes + 1
                anv_pessoa = (x['Data'] [index])
                anv_pessoa_new = str(anv_pessoa[:5])

                if str(dia_hoje) in str(anv_pessoa_new):#Se a condicao bater, é anv da pessoa
                    num_f = x['Whatsapp'][index]
                    nome_pessoa = x['Nome'][index]
                    numero_pessoa = int(num_f)
                    #print(link_p1 + str(numero_pessoa) + linkp_2)
                    self.link_final = link_p1 +  str(numero_pessoa) + '&text=' + nome_pessoa + ', ' + linkp_2

                    if warnning == True:
                        #TEMPORARIO
                        time.sleep(2)
                        campo_de_texto = self.driver.find_element_by_xpath('//*[@id="w3review"]')
                        campo_de_texto.click()
                        time.sleep(2)
                        pyautogui.press("enter")
                        campo_de_texto.send_keys('[ OCORREU UM ERRO AO ENVIAR A ÚLTIMA MENSAGEM, O CONTATO NÃO POSSUI WHATSAPP ]')
                        time.sleep(1)
                        pyautogui.press("enter")
                        time.sleep(2)

                    warnning = False

                    #TEMPORARIO
                    #
                    time.sleep(2)

                    campo_de_texto = self.driver.find_element_by_xpath('//*[@id="w3review"]')
                    time.sleep(2)
                    campo_de_texto.click()
                    time.sleep(2)
                    campo_de_texto.send_keys('      --- STATUS DE OPERAÇÃO ---')
                    time.sleep(1)
                    
                    if repeticoes > 4:
                        print('Timer delay iniciado:')
                        for timer_delay in range(60, -1, -1):
                            campo_de_texto.send_keys('Retomando em: '+str(timer_delay))
                            campo_de_texto.send_keys('enter')
                            time.sleep(1)


                    pyautogui.press("enter")
                    campo_de_texto.send_keys('- Prosseguindo para o próximo envio -')
                    time.sleep(1)
                    pyautogui.press("enter")
                    campo_de_texto.send_keys('- Validando informações...')
                    pyautogui.press("enter")
                    time.sleep(1)
                    campo_de_texto.send_keys('- Verificando base de dados...')
                    pyautogui.press("enter")
                    time.sleep(2)
                    campo_de_texto.send_keys('- UM NOVO ANIVERSARIANTE FOI ENCONTRADO!')
                    pyautogui.press("enter")
                    time.sleep(0.3)
                    campo_de_texto.send_keys('- Coletando informações do aniversariante...')
                    pyautogui.press("enter")
                    time.sleep(1)
                    campo_de_texto.send_keys('- Contatando API do WhatsApp../')
                    pyautogui.press("enter")
                    time.sleep(2)
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    campo_de_texto.send_keys('- [STATUS: 200(OK)]')
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(1)
                    campo_de_texto.send_keys('- Um momento...')
                    pyautogui.press("enter")
                    time.sleep(2)
                    campo_de_texto.send_keys('=> Enviando mensagem para '+nome_pessoa)
                    time.sleep(1)

                    box_link = self.driver.find_element_by_xpath('//*[@id="link_final"]')
                    box_link.click()
                    time.sleep(3)

                    box_link.send_keys(self.link_final)
                    time.sleep(2)

                    btExecutarLink = self.driver.find_element_by_xpath('//*[@id="bt_ex"]') 
                    btExecutarLink.click()
                    time.sleep(2)
                    bt_iniciar_conversa = self.driver.find_element_by_xpath('//*[@id="action-button"]')
                    bt_iniciar_conversa.click()
                    time.sleep(2)

                    bt_wpp = self.driver.find_element_by_xpath('//*[@id="fallback_block"]/div/div/a')
                    bt_wpp.click()
                    time.sleep(10)
                    #failure_element = 
                    try:
                        bt_enviar = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button/span')
                        time.sleep(5)
                        bt_enviar.click()
                        time.sleep(5)
                        self.driver.get('https://sendbot.netlify.app/')
                    except:
                        time.sleep(5)
                        self.driver.get('https://sendbot.netlify.app/')
                        warnning = True

                    time.sleep(4)
                    
                    now = datetime.datetime.now()
                    dia_hoje = str(now.strftime("%d") +'.'+ now.strftime("%m"))
                else:
                    print(x['Nome'][index]+" não faz aniversário hoje!")
                
                if repeticoes >= numero_de_linhas:
                    print("")
                    print("")
                    print("                                                             ************* fim do laco****")
                    print("")
                    print("")
                    time.sleep(10) 

bot = WhatsappBot()
bot.EnviarMensagens()







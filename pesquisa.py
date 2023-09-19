#######################################################  Configuração  ############################################################
# [paciente] para nome do paciente
# [alta] para data da alta do paciente
# \n para quebras de linha

mensagemWhatsapp = "Prezado(a) [paciente]," + "\n" +"Gostaríamos de convidá-lo (a) a participar da pesquisa de satisfação do Ganep Lar. "+"\n"+"Sua opinião é extremamente valiosa para nós. "+"\n"+"Levará aproximadamente 1 minuto do seu tempo e suas respostas nos ajudarão a aprimorar continuamente nossos serviços para atender às suas necessidades de maneira mais eficaz."+"\n"+"Para acessar a pesquisa de satisfação, clique no link abaixo:"

mensagemEmail = "Olá Bom dia.\nEstou encaminhando abaixo uma lista com os equipamentos que devem ser retirados, referente aos pacientes que tiveram alta"

assuntoEmail = "Retirada de equipamentos"

# Credênciais WebMail GanepLar
webmailGanepLar="https://webmail.ganeplar.com.br/"
emailGanepLar="retiradas@ganeplar.com.br"
senhaGanepLar="5eSSS3294672@"

emailsFalhados = "daniel.torres@dstorres.com.br"

###################################################################################################################################



from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import os
import time
import requests
import math
import pyautogui
import pandas as pd

# API
agent = {
    "User-Agent": 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'}

# CHAVE    xgLNUFtZsAbhZZaxkRh5ofM6Z0YIXwwv
api = requests.get(
    "https://editacodigo.com.br/index/api-whatsapp/xgLNUFtZsAbhZZaxkRh5ofM6Z0YIXwwv",  headers=agent)
time.sleep(1)
api = api.text
api = api.split(".n.")
bolinha_notificacao = api[3].strip()
contato_cliente = api[4].strip()
caixa_msg = api[5].strip()
msg_cliente = api[6].strip()
caixa_msg2 = api[7].strip()
caixa_pesquisa = api[8].strip()


dir_path = os.getcwd()
chrome_options2 = Options()
# chrome_options2.add_argument("--kiosk")
chrome_options2.add_argument(r"user-data-dir=" + dir_path + "/pasta/sessao")
driver = webdriver.Chrome(chrome_options2)



# variaveis de contato falhado
falhasWhatsApp = []
falhasEmail = []

# dicionario de empresa -> equipamentos
equipamentos = {}
emailsDict={}
empresas=[]
pacientes = {}

# função de delay (só para escrever menos sempre que for dar delay rs)
def delay(tempo):
    time.sleep(tempo)

# função para gerar os arquivos csv's referentes a cada prestador
def gerarPlanilhas():
    # Lendo e salvando o csv de equipamentos/altas em uma variável
    df = pd.read_csv('planilhas/EquipamentosParaEmail.csv', header=0, delimiter=';', encoding="ansi", engine="python")
    equips = df.to_numpy()

    # Loop para pegar cada linha dentro do csv
    for equip in equips:
        #Atualizando o email na lista de prestador => email
        emailsDict.update({
                equip[0] : equip[1]
            })
        # Caso o dicionário de prestador => lista equipamentos ja tenha um equipamento dentro
        # Atualizando para ter mais equipamentos na lista
        if equipamentos.__contains__(equip[0]):
            equipamentos.update({
                equip[0]: equipamentos.get(equip[0])+ [list( equip[2:])]
            })
        else:
            # Caso o dicionário de prestador => lista equipamentos ainda não tenha um equipamento dentro
            # Colocando o equipamento numa lista dentro do dicionário prestador => lista equipamentos
            equipamentos.update({
                equip[0]: [
                    list( equip[2:])
                    ]})
    # Loop para salvar um Xlsx com ["EQUIPAMENTO","PACIENTE","ALTA","ENDERECO"] de cada paciente/equipamento na lista de cada prestador 
    for key in emailsDict:
        # Criando o arquivo Xlsx de cada prestador
        df = pd.DataFrame(equipamentos.get(str(key)), columns=["EQUIPAMENTO","PACIENTE","ALTA","ENDERECO"])
        df.to_excel("listas/"+str(key)+".xlsx",index=False)

# Função para enviar emails para todos os prestadores anexando a lista de equipamentos a serem retirados pelo mesmo
def emails():
    # função para gerar os arquivos csv's referentes a cada prestador
    gerarPlanilhas()
    # Enviando email para cada prestador com a lista de objetos a serem retirados
    ganep = abrirWebMail(
        # Link do WEBMAIL
        webmailGanepLar,
        # Email de login
        emailGanepLar,
        # Senha de login
        senhaGanepLar)
    for value in emailsDict:
        enviarEmail(
            # pegando o Email pela lista de emails -> prestador
            # "daniel.torres@dstorres.com.br",
            str(emailsDict.get(value)),
            # Assunto
            assuntoEmail,
            # Mensagem
            mensagemEmail,
            # Pegando a planilha do prestador
            value + ".xlsx",
            # Caso seja o webmail do ganep mudam os campos a serem clicados
            ganep)
    relatarFalhas()

# Função para esperar por um elemento até que ele seja renderizado
# Substituindo assim um delay fixo que pode ocasionar em tempos de espera desnecessários e ou travamentos/falhas por esperas curtas demais
def waitForElement(selector, value):
    # repetindo função até encontrar o elemento
    while True:
        try:
            # Seleção de selector
            match selector:
                # Caso o selector sea ID
                case "id":
                    return driver.find_element(By.ID,value)
                # Caso o selector sea XPATH
                case "xpath":
                    return driver.find_element(By.XPATH,value)
                # Caso o selector sea CLASS_NAME
                case "class":
                    return driver.find_element(By.CLASS_NAME,value)
        except:
            print("Esperando por elemento " + value)
            delay(1)

def abrirWebMail(link,emailLogin,senhaLogin):
    ganep = str(link).lower().__contains__("ganep")
    print(ganep)
    # Abrindo o WEBMAIL
    driver.get(link)
    delay(3)
    logar = True
    # função para verificar se ja estava logado
    try:
        loggedAs = driver.find_element(By.XPATH if ganep else By.CLASS_NAME, "/html/body/div[1]/div[3]/div[1]/span" if ganep else "lm-user-name")
        # Verificação se o email logado é o mesmo do email solicitado
        print(loggedAs.text.lower())
        if loggedAs.text.lower() == str(emailLogin).lower():
            print("Já estava logado com o email certo")
            logar = False
        else:
            # deslogando caso o email logado não seja o mesmo do solicitado
            print("Deslogando, email solicitado diferente do logado")
            # Abrindo o menu dropdown do usuário
            driver.find_element(By.ID,"userinfomenulink" if ganep else "lm-user-dropdown-options").click()
            delay(1)
            # Clicando no botão sair do sistema
            driver.find_element(By.ID,"rcmbtn112" if ganep else "rcmbtn105")
    except:
        print("")
    finally:
        if logar:
            #Logando no sistema
            print("Logando")

            # Selecionando o input EMAIL
            login = driver.find_element(By.ID,"rcmloginuser")
            login.clear()
            login.send_keys(emailLogin)

            # Selecionando o input SENHA
            senha = driver.find_element(By.ID,"rcmloginpwd")
            senha.clear()
            senha.send_keys((senhaLogin),Keys.ENTER)
    return ganep

# Função para enviar email definindo [destino, assunto, msg, emailLogin, senhaLogin, arquivo]
def enviarEmail(destino, assunto, msg, arquivo, ganep):
    delay(2)
    global falhasEmail
    if not destino or destino == math.nan or len(destino) < 7: 
        falhasEmail.append(str(arquivo))
        print(str(arquivo).replace(".xlsx",""))
        print("Email não inserido")
        return False
    # função para esperar pela renderização do botão de ESCREVER novo email
    # Troca de elemento caso seja do GANEP ou do DSTORRES
    waitForElement("id", "rcmbtn107" if ganep else "rcmbtn110").click()

    delay(4)

    # Verificando se existe algum email não finalizado e excluindo o mesmo
    try:
        driver.find_element(By.XPATH,"/html/body/div[4]/div[3]/div/button[2]").click()
        print("Fechando email anterior que fora deixado aberto")
    except:
        print("nenhum email anterior aberto")

    # usando metodo de espera por elemento para não ter que colocar um delay fixo o que poderia gerar esperas desnecessárias e ou falhas pelo elemento ainda não ter sido renderizado
    # Troca de elemento caso seja do GANEP ou do DSTORRES
    waitForElement("xpath" if ganep else "id",
                   "/html/body/div[1]/div[4]/div[2]/form/div[1]/div/div[2]/div/div/ul/li/input" if ganep else "_to").send_keys(destino)
    
    # Escrevendo o assunto e a mensagem
    driver.find_element(By.ID,"compose-subject").send_keys(assunto,Keys.TAB,msg)

    # anexando arquivo caso solicitado
    if arquivo: 
        if arquivo == "falhasEmails":
            arquivo = ""
            for prestador in falhasEmail:
                arquivo+="\""+ prestador+"\" "
        anexar(arquivo,ganep)
    delay(6)

    # função para esperar pela renderização do botão de ENVIO do email
    # Troca de elemento caso seja do GANEP ou do DSTORRES
    waitForElement("id", "rcmbtn116" if ganep else "rcmbtn108").click()
    delay(6)

def anexar(arquivo, ganep=False):
    print("Anexando " + arquivo)

    # Clicando em ANEXAR arquivo
    # Troca de elemento caso seja do GANEP ou do DSTORRES
    driver.find_element(By.XPATH if ganep else By.ID,
                        "/html/body/div[1]/div[3]/div[2]/div[1]/div/button[1]" if ganep else "fileuploadbtn").click()
    
    # Caso seja no webmail do ganep essas funções não são necessárias
    if not ganep:
        # função para esperar pela renderização do botão de UPLOAD de arquivo
        waitForElement("id","uploadformFrm").click()

    delay(2)

    # Escrevendo o nome do arquivo
    pyautogui.write(arquivo)
    pyautogui.press('enter')
    delay(2)

    # Caso seja no webmail do ganep essas funções não são necessárias
    if not ganep:
        # função para esperar pela renderização do botão de ENVIO do arquivo anexado
        waitForElement("id","rcmbtn110").click()

# Função para ler o csv de números de whatsapp dos pacientes e enviar mensagem
def mensagens():
    global pacientes
    df = pd.read_csv('planilhas/Whatsapp.csv', header=0,delimiter=';')
    listaPacientes = df.to_numpy()
    print(listaPacientes)

    # Enviando mensagem para cada paciente dentro da lista do csv
    for paciente in listaPacientes:
        pacientes.update({
            str(paciente[2]):[paciente[0],paciente[1]]
        })
        enviar_msg("+55"+str(paciente[2]), mensagemWhatsapp, paciente)
    gerarFalhasWhatsapp()

# Função para abrir o whatsapp em uma conversa no modo WhatsApp Web
def modoWhatsWeb(numero):
    # Loop infinito até que consiga passar por todas as condicoes
    try:
        condicoes = [
            # Abrindo o whatsapp na conversa do número
            driver.get('https://wa.me/' + numero),
            # Apertando enter (Usado para quando mandar trocar de tela conseguir passar pelo popup de "Deseja mesmo sair dessa tela")
            pyautogui.hotkey('enter'),
            delay(1),
            # Clicando no botão INICIAR CONVERSA
            driver.find_element(By.ID, 'action-button').click(),
            delay(2),
            # Clicando no botão de INICIAR VIA WEB
            driver.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[2]/div/section/div/div/div/div[3]/div/div/h4[2]/a').click()
        ]
        return True if all(condicoes) else False
    except:
        modoWhatsWeb(numero)

# Função para enviar mensagem via whatsapp
def enviar_msg(numero, msg,paciente=[]):
    # colocando o whatsapp em uma conversa no modo WhatsApp Web
    modoWhatsWeb(numero)
    # Loop de 30 tentativas de envio
    for i in range(30):
        try:
            # Enviando a mensagem de texto via whatsapp
            campo_de_texto = driver.find_element(By.XPATH, caixa_msg)
            campo_de_texto.clear()
            msg = str(msg).replace("[paciente]",paciente[0]).replace("[alta]",paciente[1])
            campo_de_texto.send_keys(msg, Keys.ENTER) 
            print('enviado')
            return True
        except:
            print('Error ' + str(i))
            delay(1)
    # em caso de falha printando a falha e salvando em uma variavel para enviar por email posteriormente
    print("Número: " + str(numero) + " não tem whatsapp ou está com o número errado")
    # pegando a variável falha global
    global falhasWhatsApp
    falhasWhatsApp.append(str(paciente[2]))
    return False

def gerarFalhasWhatsapp():
    
    # Criando um Xlsx com todas as falhas de whatsapp causadas por numeros errados ou pessoas sem whatsapp
    listaPacientes = []
    for item in falhasWhatsApp:
        paciente = pacientes.get(item)
        listaPacientes.append([paciente[0],paciente[1],item])

    df = pd.DataFrame(listaPacientes, columns=["PACIENTE","ALTA","CELULAR"])
    df.to_excel("listas/falhasWhatsapp.xlsx",index=False)
    



    
def relatarFalhas():
    ganep = abrirWebMail(
        # Link do WEBMAIL
        "https://webmail-seguro.com.br/dstorres.com.br/",
        # Email de login
        "daniel.marcondes@dstorres.com.br",
        # Senha de login
        "583492761Mt$")
    
    enviarEmail(
        emailsFalhados,
        "Falhas nos envios do whatsapp",
        "Olá Bom dia.\nSegue abaixo uma planilha com todas as falhas das mensagens via whatsapp",
        "falhasWhatsApp.xlsx",
        ganep
    )
    
    mensagem = "Olá Bom dia.\nSegue abaixo as planilhas que não foram enviadas aos prestadores por conterem emails incorretos e ou estavam sem email declarado, são elas:\n"
    for prestador in falhasEmail:
        mensagem += str(prestador).replace(".xlsx","") + " está sem email cadastrado\n"
    enviarEmail(
        emailsFalhados,
        "Falhas nos envios dos emails de retirada",
        mensagem,
        "falhasEmails",
        ganep
    )
    
mensagens()
emails()

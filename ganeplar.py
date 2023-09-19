# Importando bibliotecas necessárias
import time
import pyautogui
import subprocess
import pyperclip

# Função para adicionar um delay
def delay(tempo):
    time.sleep(tempo)

# Função para escrever texto usando o clipboard
def write(texto, tempo=1, enter=False):
    pyperclip.copy(texto)
    press('ctrl','v')
    if enter: 
        delay(1)
        press('enter')
    delay(tempo)
    return True

# Função para clicar em uma posição específica na tela
def click(x,y,duplo=False):
    old = pyautogui.PAUSE
    pyautogui.PAUSE = .05
    pyautogui.click(x, y)
    if duplo: pyautogui.click(x, y)
    pyautogui.PAUSE = old

# Função para pressionar teclas de atalho
def press(a,b=None,c=None,tempo=1):
    if c: pyautogui.hotkey(a,b,c)
    elif b: pyautogui.hotkey(a,b)
    else: pyautogui.hotkey(a)
    delay(tempo)
    print(pyautogui.PAUSE)
    return True
                
# Função para fechar janelas ou processos
def fechar(pasta = ""):
    if not pasta: 
        press('alt','f4')
    elif "exe" in pasta:
        subprocess.call(["taskkill","/F","/IM", pasta])
    elif pasta == "explorer":
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq Planilhas"])
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq updater"])
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq Mapa Atual"])
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq Explorador de Arquivos"])
    else:
        subprocess.call(["taskkill","/F","/IM", "explorer.exe","/FI", "WINDOWTITLE eq "+pasta+".csv - Resultados da Pesquisa em updater"])

# Função para abrir o aplicativo GanepLar
def abrirGanepLar():
    clicar('iw-ganeplar.png', 1, 0.85, duplo = True, confirmacao='login.png', esperaMax=60)
    clicar('login.png', 2, 0.9,pos=[497, 313, 1097, 713])
    write("Ganeplar@7823")
    delay(1)
    press('enter',tempo=10)
    
# Função para baixar uma planilha
def baixarPlanilha(path,tempo):
    delay(2)
    clicar('administrador.png', 1,.85,confirmacao='planilhas_rpa.png')
    if not clicar('planilhas_rpa.png', 1,.85,confirmacao='satisfacao.png',restart=False,):
        baixarPlanilha(path,tempo)
    else:
        clicar(path, tempo, .8)

# Função para salvar a planilha
def salvar(path,nome,tempo = 1):
    baixarPlanilha(path,tempo)
    clicar('exportar.png',5,.7,pos=[ 299, 0, 899, 413])
    clicar('desktop.png', conf=0.85,duplo = True, pos=[293, 81, 893, 681])
    clicar('ganep.png',duplo = True, pos=[284, 40, 884, 640])
    clicar('chatbot.png',duplo = True, pos=[285, 39, 885, 639])
    clicar('planilhas.png',duplo = True, pos=[285, 39, 885, 639])
    clicar('nomearquivo.png')
    write(nome)
    clicar('salvar.png', 3, .7)
    fechar()
    while not esperarTela('rel_fechado.png',vezes=80): fechar()

# Função para esperar por uma imagem na tela
def esperarTela(path, tempo=0, conf=.95, vezes=0, clicar=False, duplo=False, confirmacao='', restart=True, esperaMax=25, pos=[0,0,1600,900], posconf=[0,0,1600,900]):
    confiabilidade = conf
    tentativas = 0
    # Criando um loop de espera pela imagem
    while tentativas <= vezes or vezes == 0:
        # Colocando dentro de um try para caso a imagem ainda não esteja na tela
        try:
            tentativas+=1
            # Procurando pela imagem na tela
            print("Procurando " + str(path) + " pela " + str(tentativas) + "° vez")
            if tentativas > 20 and confiabilidade > .60: confiabilidade -=.005
            x, y = pyautogui.locateCenterOnScreen(path, confidence=confiabilidade, region=pos)
            print(str(path) + " encontrado na posição: " + str(x) + ", " + str(y))
            delay(1.5)
            if clicar or duplo: click(x,y,duplo)
            # Saindo do loop
            if confirmacao:
                if esperarTela(confirmacao, 1, .85, vezes = esperaMax*2, pos= posconf,restart=False):
                    print("Foi confirmado que o click em " + str(path) + " foi bem sucedido")
                    delay(tempo)
                    return True
                elif restart:
                    print("Confirmação de click mal sucedida reiniciando processo")
                    tentativas=0
                else:
                    return False
            else:
                delay(tempo)
                return True
        except :
            delay(.5)
            print("não achou " + str(path) + " confiabilidade = " + str(confiabilidade))
            if tentativas >= vezes > 0:
                print("tentativas encerradas")
                return False


# Função para clicar em uma imagem quando ela aparecer
def clicar(path, tempo=0, conf=.95, duplo = False, vezes=0,confirmacao='',restart=True,esperaMax=0, pos=[0,0,1600,900], posconf=[0,0,1600,900]):
    try:
        return esperarTela(path, tempo, conf, vezes=vezes, clicar=True, duplo=duplo,confirmacao=confirmacao,restart=restart,esperaMax=esperaMax, pos=pos, posconf=posconf)
    except:
        return False 

# Função para baixar várias planilhas
def baixarplanilhas():
    salvar("satisfacao.png","realwpp.csv")
    salvar("equipamentos.png","realwpp.csv")
    delay(1)
    fechar()
    delay(1)
    clicar('confirmar_fechar.png',pos=[ 375, 260, 1175, 660]) 

# Executando as funções principais
abrirGanepLar()
baixarplanilhas()

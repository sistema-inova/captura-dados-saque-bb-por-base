import pyautogui
import time
from comtypes.client import CreateObject
from datetime import datetime
import getpass
import os 
import sys 


class Sisfin:
    
    def __init__(self, matricula, senha):
        if getattr(sys, 'frozen', False):
            # The application is frozen
            nome_root = os.path.dirname(sys.executable)
        else:
            # The application is not frozen
            # Change this bit to match where you store your data files:
            nome_root = os.path.dirname(__file__)
        caminho_sisfin = os.path.join(nome_root, 'SISFIN.edp')
        os.startfile(caminho_sisfin)
        time.sleep(3)

        self.system = CreateObject("Extra.System", dynamic=True)
        self.session = self.system.ActiveSession
        self.screen = self.session.Screen

        # pyautogui.keyDown('win')
        # pyautogui.keyDown('up')
        # pyautogui.keyUp('win')
        # pyautogui.keyUp('up')

        time.sleep(1)
        self.sisfin_escrever(self.screen, 19, 11, matricula)
        time.sleep(1)
        self.screen.SendKeys(senha)
        self.screen.SendKeys("<Enter>")
        time.sleep(1)

    def sisfin_finalizar(self):  
        os.system("taskkill /f /im EXTRA.EXE")

    def sisfin_voltar_pagina_inicial_cardapio(self, screen, linha):
        print("")
        # VOLTAR PARA A PÁGINA INICIAL DO CARDÁPIO
        self.sisfin_escrever(screen, linha, 8, '.')
        time.sleep(0.5)

    def sisfin_ler(self, screen, row, col, length):
        return screen.Area(row, col, row, col+length).value  

    def sisfin_escrever(self, screen, row, col, text):
        screen.row = row
        screen.col = col
        screen.SendKeys(text)
        time.sleep(0.5)
        screen.SendKeys("<Enter>")
        time.sleep(0.2)

    def vmcy_acessar(self, screen, menu, periodo):
        time.sleep(0.5)
        self.sisfin_escrever(screen, 21,8,'VMCY')
        time.sleep(1)
        while (self.sisfin_ler(screen, 2, 1,7) == 'Cardapio'):
            time.sleep(1)

        self.sisfin_escrever(screen, 4, 13, 'T')
        time.sleep(0.5)
        self.sisfin_escrever(screen, 4, 44, menu)
        time.sleep(0.5)
        self.sisfin_escrever(screen, 5, 17, 'T')
        time.sleep(0.5)
        self.sisfin_escrever(screen, 5, 60, periodo)
        time.sleep(0.5)
        self.sisfin_escrever(screen, 5, 71, periodo)
        time.sleep(0.5)
        self.sisfin_escrever(screen, 6, 17, 'T')
        time.sleep(1.5)

    def vmcy_tratamento_sem_movimentacao(self, screen, celula_planilha):
        print("R$ -") 
        celula_planilha = 0
        self.sisfin_escrever(screen, 23, 50, '.')
        time.sleep(0.5)
        self.sisfin_escrever(screen, 6, 17, '.')
        time.sleep(0.5)

    def vmcb_acessar(self, screen, periodo, opcao):
        time.sleep(0.5)
        self.sisfin_escrever(screen, 21,8,'VMCB')
        time.sleep(1)
        while (self.sisfin_ler(screen, 2, 1,7) == 'Cardapio'):
            time.sleep(1)

        self.sisfin_escrever(screen, 4, 12, periodo)
        time.sleep(0.5)
        self.sisfin_escrever(screen, 5, 12, opcao)
        time.sleep(1.5)

        contagem = 0
        while self.sisfin_ler(screen, 9, 1, 2).strip() == '' and contagem <= 4:
            time.sleep(0.5)
            contagem += 1

    def vmcb_tratamento_sem_movimentacao(self, screen, celula_planilha):
        celula_planilha = 0
        print("Não há movimentações nesta data...")
        print("")
        self.sisfin_escrever(screen, 5, 12, '.')

    def vmcd_acessar_detalhando_menu_opcao(self, screen, periodo, opcao):
        self.sisfin_escrever(screen, 21,8,'VMCD')
        time.sleep(1)
        while (self.sisfin_ler(screen, 2, 1,7) == 'Cardapio'):
            time.sleep(1)

        self.sisfin_escrever(screen, 4, 17, 'CEF')
        time.sleep(0.5)
        self.sisfin_escrever(screen, 5, 17, periodo)
        time.sleep(1)

        self.sisfin_escrever(screen, 21, 8, '10')
        time.sleep(1.5)

    def vmcd_acessar_sem_detalhar_menu_opcao(self, screen, periodo):
        time.sleep(0.5)
        self.sisfin_escrever(screen, 21,8,'VMCD')
        time.sleep(1)
        while (self.sisfin_ler(screen, 2, 1,7) == 'Cardapio'):
            time.sleep(1)

        self.sisfin_escrever(screen, 4, 17, 'CEF')
        time.sleep(0.5)
        self.sisfin_escrever(screen, 5, 17, periodo)
        time.sleep(1.5)

    def vmcd_tratamento_sem_movimentacao(self, screen, celula_planilha):
        print("Não há movimentações nesta data...")
        celula_planilha = 0 
        print("")
        self.sisfin_escrever(screen, 21, 8, '.')

    def veca_acessar(self, screen, periodo):
        time.sleep(0.5)
        self.sisfin_escrever(screen, 21,8,'VECA')
        time.sleep(1)
        while (self.sisfin_ler(screen, 2, 1,7) == 'Cardapio'):
            time.sleep(1)

        self.sisfin_escrever(screen, 4, 17, 'CEF')
        time.sleep(0.5)
        self.sisfin_escrever(screen, 5, 17, periodo)
        time.sleep(1.5)

    def veca_tratamento_sem_movimentacao(self, screen, celula_planilha):
        print("Sem movimentação nesta data")
        print("")
        celula_planilha = 0
        self.sisfin_escrever(screen, 21, 8, '.')
        time.sleep(0.5)

    def tirar_printscreen(self, tipo_relatorio, nome_arquivo):
        if tipo_relatorio != 'parcial':
            pasta_prints = r'prints' 
            if not os.path.exists(pasta_prints):
                os.makedirs(pasta_prints)
            screenshot = pyautogui.screenshot()
            screenshot.save(rf"prints\\{nome_arquivo}.png")
            time.sleep(2)
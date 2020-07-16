import time
import pyodbc
import pyautogui
from openpyxl import Workbook
import time
import win32com.client
from credenciais import USUARIO, SENHA
# from classes.sisfin import Sisfin
from datetime import date, timedelta, datetime
import getpass
import os 
import sys 

def capturar_dados_sifin():

    print("")
    escrever_cabecalho('CAPTURA DADOS SAQUES BACEN POR BASE')

    # usuario = input('Matrícula: ')
    # senha = getpass.getpass('Senha: ')

    usuario = USUARIO
    senha = SENHA

    if getattr(sys, 'frozen', False):
        nome_root = os.path.dirname(sys.executable)
    else:
        nome_root = os.path.dirname(__file__)
    caminho_sisfin = os.path.join(nome_root, 'SISFIN.edp')
    os.startfile(caminho_sisfin)
    time.sleep(3)

    system = win32com.client.Dispatch("EXTRA.System")
    sess0 = system.ActiveSession
    screen = sess0.Screen

    pyautogui.keyDown('win')
    pyautogui.keyDown('up')
    pyautogui.keyUp('win')
    pyautogui.keyUp('up')

    time.sleep(1)
    write(screen, 22,11,usuario)
    time.sleep(1)
    screen.SendKeys(senha)
    enter(screen)
    time.sleep(1)
    while (read(screen, 2, 1, 7) != 'Cardapio'):
        time.sleep(1)

    lista_tesourarias = ["TESOURARIA AL", "TESOURARIA AM", "TESOURARIA BA", "TESOURARIA BR", "TESOURARIA CE", "TESOURARIA ES", "TESOURARIA GO", "TESOURARIA JF", "TESOURARIA MA", "TESOURARIA MG", "TESOURARIA MS", "TESOURARIA MT", "TESOURARIA PA", "TESOURARIA PB", "TESOURARIA PE", "TESOURARIA PI", "TESOURARIA PR", "TESOURARIA RJ", "TESOURARIA RN", "TESOURARIA RS", "TESOURARIA SC", "TESOURARIA SE", "TESOURARIA SP", "TESOURARIA UB"]
    for tesouraria in lista_tesourarias:
        for dia in range(0, 11):
            if dia == -1:
                periodo = f"H-1"
                data_periodo = date.today() - timedelta(days=1)
            elif dia == 0:
                periodo = f"H"
                data_periodo = date.today()
            else:
                periodo = f"H+{dia}"
                data_periodo = date.today() + timedelta(days=dia)

            vmrd_acessar(screen, tesouraria, periodo) 
            time.sleep(1)
            print(read(screen, 5, 19, 10).strip())
            if read(screen, 5, 19, 10).strip() != "":
                while read(screen, 6, 19, 4).strip() != "":          
                    data_atendimento = data_periodo
                    estado_tesouraria = tesouraria[-2:]
                    nome_base = read(screen, 9, 21, 30).strip()
                    operacao = read(screen, 6, 19, 4).strip()
                    tipo_trasporte = read(screen, 10, 21, 20).strip()
                    tarifa = read(screen, 10, 61, 17).strip().replace('.', '').replace(',', '.') if read(screen, 10, 61, 20).strip().replace('.', '').replace(',', '.') != "" else '0.00'
                    preposto_informado = read(screen, 11, 21, 20).strip()
                    valor_circulante = read(screen, 19, 24, 17).strip().replace('.', '').replace(',', '.') if read(screen, 19, 24, 17).strip().replace('.', '').replace(',', '.') != "" else '0.00'
                    valor_dilacerado = read(screen, 19, 43, 17).strip().replace('.', '').replace(',', '.') if read(screen, 19, 43, 17).strip().replace('.', '').replace(',', '.') != "" else '0.00'
                    situacao = read(screen, 22, 11, 50).strip()
                    cir = read(screen, 23, 23, 50).strip()
                    time.sleep(0.5)

                    print(f"Data atendimento: {data_atendimento} - UF: {estado_tesouraria} - Nome Base: {nome_base} - Operação: {operacao} - Tipo Transporte: {tipo_trasporte} - Tarifa: {tarifa} - Preposto Informado: {preposto_informado} - Valor Circulante: R$ {valor_circulante} - Valor Dilacerado: R$ {valor_dilacerado} - Situação: {situacao} - CIR: {cir}")

                    write(screen, 21, 8, 'N')
                write(screen, 6, 19, '.')
            else:
                print('entrou no false')
                # write(screen, 6, 19, '.')
                pyautogui.keyDown('esc')
        

def read(screen, row, col, length):
    return screen.Area(row, col, row, col+length).value  

def write(screen, row, col, text):
    screen.row = row
    screen.col = col
    screen.SendKeys(text)
    time.sleep(0.5)
    enter(screen)
    time.sleep(0.2)

def voltar_pagina_inicial_cardapio(screen, linha):
    print("")
    # VOLTAR PARA A PÁGINA INICIAL DO CARDÁPIO
    write(screen, linha, 8, '.')
    time.sleep(0.5)

def escrever_cabecalho(mensagem):
    print("{:*^100}" .format("*"))
    print("{:*^100}" .format(f" {mensagem} "))
    print("{:*^100}" .format("*"))
    print("")

def enter(screen):
    return screen.SendKeys("<Enter>")

def vmrd_acessar(screen, tesouraria, periodo):
    time.sleep(0.5)
    write(screen, 21,8,'VMRD')
    time.sleep(1)
    while (read(screen, 2, 1,7) == 'Cardapio'):
        time.sleep(1)

    write(screen, 4, 19, tesouraria)
    time.sleep(0.5)
    write(screen, 5, 19, periodo)
    time.sleep(0.5)
    if read(screen, 6, 19, 4) != 'Saque':
        write(screen, 6, 17, 'S')
        time.sleep(0.5)


if __name__ == "__main__":
    capturar_dados_sifin()
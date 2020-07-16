import time
import pyodbc
import pyautogui
from openpyxl import Workbook
import time
import win32com.client
from credenciais import *
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

    conexao_banco = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USER};PWD={PASSWORD}")
    cursor = conexao_banco.cursor()

    if not cursor.tables(table='TP_90003_04_TEMP_SAQUE_BB_BACEN_POR_BASE', tableType='TABLE').fetchone():
        consulta = "CREATE TABLE [dbo].[TP_90003_04_TEMP_SAQUE_BB_BACEN_POR_BASE]([dtDataAtendimento] [date], [chSiglaTesouraria] [char](2), [vcNomeBase] [varchar](50), [vcTipoOperacao] [varchar](10), [vcTipoTransporte] [varchar](50), [dcValorTarifa] [decimal](17, 2), [chPrepostoInformado] [char](3), [dcValorCirculante] [decimal](17, 2), [dcValorDilacerado] [decimal](17, 2), [vcSituacao] [varchar](150), [vcCir] [varchar](50), [dtDataCaptura] [datetime])"
        cursor.execute(consulta)
        conexao_banco.commit()

    lista_tesourarias = ["TESOURARIA AL", "TESOURARIA AM", "TESOURARIA BA", "TESOURARIA BR", "TESOURARIA CE", "TESOURARIA ES", "TESOURARIA GO", "TESOURARIA JF", "TESOURARIA MA", "TESOURARIA MG", "TESOURARIA MS", "TESOURARIA MT", "TESOURARIA PA", "TESOURARIA PB", "TESOURARIA PE", "TESOURARIA PI", "TESOURARIA PR", "TESOURARIA RJ", "TESOURARIA RN", "TESOURARIA RS", "TESOURARIA SC", "TESOURARIA SE", "TESOURARIA SP", "TESOURARIA UB"]
    for dia in range(0, 11):
        for tesouraria in lista_tesourarias:
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
            if read(screen, 5, 19, 10).strip() != "":
                while read(screen, 6, 19, 4).strip() != "": 
                    sigla_tesouraria = ''
                    nome_base = ''
                    operacao = ''
                    tipo_trasporte = ''
                    tarifa = ''
                    preposto_informado = ''
                    valor_circulante = ''
                    valor_dilacerado = ''
                    situacao = ''
                    cir = ''  
                    while sigla_tesouraria == "" and nome_base == "" and operacao == "" and tipo_trasporte == "" and tarifa == "" and preposto_informado == "" and valor_circulante == "" and valor_dilacerado == "" and situacao == "" and cir == "":
                        time.sleep(0.5)   
                        data_hora_consulta = str(f"{datetime.now():%Y-%m-%d %H:%M:%S}")
                        data_atendimento = data_periodo
                        sigla_tesouraria = tesouraria[-2:]
                        nome_base = read(screen, 9, 21, 30).strip()
                        operacao = read(screen, 6, 19, 4).strip()
                        tipo_trasporte = read(screen, 10, 21, 20).strip()
                        tarifa = read(screen, 10, 62, 14).strip().replace('.', '').replace(',', '.') if read(screen, 10, 62, 14).strip() != '' else '0.00'
                        preposto_informado = read(screen, 11, 21, 20).strip()
                        valor_circulante = read(screen, 19, 24, 17).strip().replace('.', '').replace(',', '.') if read(screen, 19, 24, 17).strip() != '' else '0.00'
                        valor_dilacerado = read(screen, 19, 43, 17).strip().replace('.', '').replace(',', '.') if read(screen, 19, 43, 17).strip() != '' else '0.00'
                        situacao = read(screen, 22, 11, 60).strip()
                        cir = read(screen, 23, 23, 50).strip()
                        time.sleep(0.5)
                        print(read(screen, 21, 16, 7).strip())
                        if read(screen, 21, 16, 7).strip() == "Alterado":
                            enter(screen)
                            time.sleep(0.5)
                            situacao = read(screen, 22, 11, 50).strip()
                            cir = read(screen, 23, 23, 50).strip()
                            write(screen, 21, 8, 'N')
                    print(f"Data atendimento: {data_atendimento} - Sigla: {sigla_tesouraria} - Nome Base: {nome_base} - Operação: {operacao} - Tipo Transporte: {tipo_trasporte} - Tarifa: {tarifa} - Preposto Informado: {preposto_informado} - Valor Circulante: R$ {valor_circulante} - Valor Dilacerado: R$ {valor_dilacerado} - Situação: {situacao} - CIR: {cir}")
                    while sigla_tesouraria != " " and nome_base != " " and operacao != " " and tipo_trasporte != " " and tarifa != " " and preposto_informado != " " and valor_circulante != " " and valor_dilacerado != " " and situacao != " " and cir != " ":
                        realiza_insert_banco(conexao_banco, cursor, data_atendimento, sigla_tesouraria, nome_base, operacao, tipo_trasporte, tarifa, preposto_informado, valor_circulante, valor_dilacerado, situacao, cir, data_hora_consulta)
                    write(screen, 21, 8, 'N')
                write(screen, 6, 19, '.')
            else:
                write(screen, 6, 19, '.')

    os.system("taskkill /f /im EXTRA.EXE")
    print("")

    escrever_cabecalho('FIM')

def realiza_insert_banco(conexao_banco, cursor, data_atendimento, sigla_tesouraria, nome_base, operacao, tipo_trasporte, tarifa, preposto_informado, valor_circulante, valor_dilacerado, situacao, cir, data_hora_consulta):
    consulta = """
        SET DATEFORMAT YMD;
        INSERT INTO [dbo].[TP_90003_04_TEMP_SAQUE_BB_BACEN_POR_BASE]
           ([dtDataAtendimento]
           ,[chSiglaTesouraria]
           ,[vcNomeBase]
           ,[vcTipoOperacao]
           ,[vcTipoTransporte]
           ,[dcValorTarifa]
           ,[chPrepostoInformado]
           ,[dcValorCirculante]
           ,[dcValorDilacerado]
           ,[vcSituacao]
           ,[vcCir]
           ,[dtDataCaptura])
        VALUES
            (CAST('{}' AS DATE)
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,'{}'
            ,CAST('{}' AS DATETIME))
        """
    cursor.execute(consulta.format(data_atendimento, sigla_tesouraria, nome_base, operacao, tipo_trasporte, tarifa, preposto_informado, valor_circulante, valor_dilacerado, situacao, cir, data_hora_consulta))
    conexao_banco.commit()      

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
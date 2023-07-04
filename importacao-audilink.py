import configparser
from datetime import datetime
import json
from pathlib import Path
import sys
import os
import pandas as pd
from tabulate import tabulate
import builtins
import logging
from robotpy.Robot import Robot

config = configparser.ConfigParser()
log = """
    <style> 
    table, th, td { border: 1px solid black; border-collapse: collapse; }
    th, td { padding: 5px; }
    </style>
"""

'''
    # find on all cells first date of the sheet
'''

def print(message):
    global log
    log += str(message) + "<br>"
    builtins.print(message)
    
def stringIsInDateFormat(string, dateFormat):
    try:
        date = pd.to_datetime(string, format=dateFormat)
        return date is not None
    except:
        return None


def Importacao_audilink(robotParameters):
    config.read('importacao-audilink.ini', encoding='utf-8')

    print('Importacao Audilink ' + str(robotParameters['month']) + '/' + str(robotParameters['year']))

    config['paths']['month'] = str(robotParameters['month']).zfill(2)
    config['paths']['year'] = str(robotParameters['year'])

    month = int(robotParameters['month'])
    year = int(robotParameters['year'])

    # 1 - Pegar pasta do arquivo e filtro do arquivo de lançamentos no arquivo de configuração
    # 2 - Pegar pasta do arquivo e filtro do arquivo de contas no arquivo de configuração
    # 3 - Pegar arquivo de lançamentos e arquivo de contas
    # 4 - Extrair lançamentos do arquivo execel
    # 5 - Extrair contas do arquivo csv
    # 6 - Para cada lançamento, susbtituir a conta pelo arquivo de contas
    # 7 - Se não existir conta no arquivo de contas - criar linha no arquivo de contas, mencionar no log e marcar que tem uma conta desconhecida
    # 8 - Caso não exista nenhuma conta desconhecida, criar arquivo de importacao e printar/log que foi criado na mesma pasta do arquivo de lançamentos
    # 9 - Caso exista conta desconhecida, mostrar link para editar o arquivo de conta.  


#Protection against running the script twice
try:
    mes_teste = 5
    ano_teste = 2023

    #Se existir argumentos, define o call_id como o primeiro parametro, se não, define como None
    call_id = sys.argv[1] if len(sys.argv) > 1 else None

    #start robot with first argument
    robot = Robot(call_id)

    try:
        #from robot.parameters get 'mes' and 'ano' as int
        mes = int(robot.parameters['mes']) if call_id is not None else mes_teste
        ano = int(robot.parameters['ano']) if call_id is not None else ano_teste

        try:
            # call the main function passing the argument list
            Importacao_audilink({'month': mes, 'year': ano})
            robot.setReturn(log)
        except Exception as e:
            logging.exception(e)
            robot.setReturn("Erro desconhecido: " + str(e))
    except Exception as e:
        print(e)
        robot.setReturn("Parametros passados invalidos: " + json.dumps(robot.parameters))
except Exception as e:
    print(e)
    pass

sys.exit(0)
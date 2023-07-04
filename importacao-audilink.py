import configparser
import json
import os
from pathlib import Path
import re
import sys
import pandas as pd
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
    
def stringIsFilterOf(string, filter):
    if filter is None or filter == '':
        return True

    return re.match(filter, string) is not None

def getFileOnFolder(folder, filter):
    for file in folder.iterdir():
        if file.is_file() and stringIsFilterOf(file.name, filter):
            return file
    return None

def Importacao_audilink(robotParameters):
    config.read('importacao-audilink.ini', encoding='utf-8')
    codigo_empresa = config['paths']['codigo_empresa']

    print('Importacao Audilink ' + str(robotParameters['month']) + '/' + str(robotParameters['year']))

    # Atualiza o mes e ano no arquivo de configuração
    config['paths']['month'] = str(robotParameters['month']).zfill(2)
    config['paths']['year'] = str(robotParameters['year'])
    config.write(open('importacao-audilink.ini', 'w'))

    month = int(robotParameters['month'])
    year = int(robotParameters['year'])

    # 1 - Pegar pasta do arquivo e filtro do arquivo de lançamentos no arquivo de configuração
    localArquivoLancamentos = config['paths']['arquivo_lancamentos']
    pastaArquivoLancamentos = Path(localArquivoLancamentos)
    if not pastaArquivoLancamentos.exists(): raise Exception('Pasta do arquivo de lançamentos não existe: ' + localArquivoLancamentos)
    
    # 2 - Pegar pasta do arquivo e filtro do arquivo de contas no arquivo de configuração
    localArquivoContas = config['paths']['arquivo_contas']
    pastaArquivoContas = Path(localArquivoContas)
    if not pastaArquivoContas.exists(): raise Exception('Pasta do arquivo de contas não existe: ' + localArquivoContas)
    
    # 3 - Procura arquivo de lançamentos pelo paths.filtro_arquivo_lancamentos
    arquivoLancamentos = getFileOnFolder(pastaArquivoLancamentos, config['paths']['filtro_arquivo_lancamentos'])
    if arquivoLancamentos is None: raise Exception('Arquivo de lançamentos não encontrado: ' + config['paths']['filtro_arquivo_lancamentos'])
    
    # 3 - Procura arquivo de contas pelo paths.filtro_arquivo_contas
    arquivoContas = getFileOnFolder(pastaArquivoContas, config['paths']['filtro_arquivo_contas'])
    if arquivoContas is None: raise Exception('Arquivo de contas não encontrado: ' + config['paths']['filtro_arquivo_contas'])

    # 4 - Extrair lançamentos do arquivo excel pelos nomes das colunas no arquivo de configuração identificados por: [arquivo_lancamentos] .coluna_conta, .coluna_data, .coluna_tipo_conta, .coluna_valor, .coluna_historico, .coluna_descricao
    colunasArquivo = [
        config['arquivo_lancamentos']['coluna_conta'],   
        config['arquivo_lancamentos']['coluna_data'],             
        config['arquivo_lancamentos']['coluna_tipo_conta'],
        config['arquivo_lancamentos']['coluna_valor'],
        config['arquivo_lancamentos']['coluna_historico'],
        config['arquivo_lancamentos']['coluna_descricao']
    ]
    def isColunaValida(coluna):
        for colunaValida in colunasArquivo:
            if stringIsFilterOf(coluna, colunaValida):
                return True
        
        return False
    
    def getColunaIndexByFilter(colunas, filter):
        for index, coluna in enumerate(colunas):
            if stringIsFilterOf(coluna, filter):
                return index
        
        return None

    lancamentos = pd.read_excel(arquivoLancamentos, usecols=isColunaValida)
    colunas = lancamentos.columns
    colunaContaIndex = getColunaIndexByFilter(colunas, config['arquivo_lancamentos']['coluna_conta'])
    colunaDataIndex = getColunaIndexByFilter(colunas, config['arquivo_lancamentos']['coluna_data'])
    colunaTipoContaIndex = getColunaIndexByFilter(colunas, config['arquivo_lancamentos']['coluna_tipo_conta'])
    colunaValorIndex = getColunaIndexByFilter(colunas, config['arquivo_lancamentos']['coluna_valor'])
    colunaHistoricoIndex = getColunaIndexByFilter(colunas, config['arquivo_lancamentos']['coluna_historico'])
    colunaDescricaoIndex = getColunaIndexByFilter(colunas, config['arquivo_lancamentos']['coluna_descricao']) 

    # 5 - Extrair contas do arquivo csv 0 - conta original, 1 - conta nova
    contas = pd.read_csv(arquivoContas, sep=';', encoding='utf-8', usecols=['CONTA', 'DE_PARA'], dtype={'DE_PARA': 'Int64'}).dropna()

    # 6 - Para cada lançamento, susbtituir a conta pelo arquivo de contas
    def getContaDePara(conta):
        contaTrim = str(conta).strip()
        for index, contaDePara in contas.iterrows():
            if str(contaDePara['CONTA']).strip() == contaTrim:
                return contaDePara['DE_PARA']
        
        return None
    
    contasInvalidas = []
    for index, lancamento in lancamentos.iterrows():
        conta = lancamento[colunaContaIndex]
        contaDePara = getContaDePara(conta)
        if contaDePara is None:            
            contasInvalidas.append(str(conta).strip())
            continue

        lancamentos.at[index, colunaContaIndex] = contaDePara

    # 7 - Se não existir conta no arquivo de contas - criar linha no arquivo de contas, mencionar no log e marcar que tem uma conta desconhecida
    if len(contasInvalidas) > 0:
        contasInvalidas = list(dict.fromkeys(contasInvalidas))

        print('Contas não encontradas no arquivo de contas:')
        for contaInvalida in contasInvalidas:
            print(contaInvalida)

            #criar linha no arquivo de contas usando o concat do pandas
            contas = pd.concat([contas, pd.DataFrame([[contaInvalida, None]], columns=['CONTA', 'DE_PARA'])], ignore_index=True)
        
        #salvar arquivo de contas
        contas.to_csv(arquivoContas, sep=';', encoding='utf-8', index=False)
        print('')
        print('Preencha o arquivo de contas com as contas corretas e rode o robô novamente.')
        print('Arquivo de contas em: ' + arquivoContas)

    # 8 - Caso não exista nenhuma conta desconhecida, criar arquivo de importacao e printar/log que foi criado na mesma pasta do arquivo de lançamentos
    if len(contasInvalidas) == 0:

        importacao = []
        for index, lancamento in lancamentos.iterrows():

            row = []
            row.append(str(codigo_empresa))
            row.append("")
            row.append("")
            row.append(lancamento.iloc[colunaDataIndex].strftime('%d/%m/%Y'))
            row.append(lancamento.iloc[colunaContaIndex] if lancamento.iloc[colunaTipoContaIndex] == 'D' else "")
            row.append(lancamento.iloc[colunaContaIndex] if lancamento.iloc[colunaTipoContaIndex] == 'C' else "")
            row.append("")
            row.append("80")

            historico = str(lancamento.iloc[colunaHistoricoIndex]) + ' - ' if str(lancamento.iloc[colunaHistoricoIndex]) != 'nan' else ''
            historico = historico + str(lancamento.iloc[colunaDescricaoIndex]) if str(lancamento.iloc[colunaDescricaoIndex]) != 'nan' else historico
            historico = historico.replace(';', ',')
            row.append(historico)
            row.append(str(lancamento.iloc[colunaValorIndex]).format('0.2f').replace('.', ','))

            importacao.append(';'.join(row))
        
        importacaoTexto = '\n'.join(importacao)
        arquivoLancamentosFolder = os.path.dirname(arquivoLancamentos)
        arquivoImportacao = os.path.join(arquivoLancamentosFolder, 'importacao_audilink_' + str(month) + '_' + str(year) + '.txt')
        with open(arquivoImportacao, 'w', encoding='utf-8') as f:
            f.write(importacaoTexto)

        print('Arquivo de importação criado em: ' + arquivoImportacao)        


#Protection against running the script twice
try:
    mes_teste = 6
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
            builtins.print('\n\n\nLOG SALVO NO ROBÔ:\n')
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
# -*- coding: utf-8 -*-
import traceback, re, pyperclip, shutil, time, os, pyautogui as p, PySimpleGUI as sg
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from threading import Thread
from pathlib import Path
from functools import wraps
from io import BytesIO
from PIL import Image

from sys import path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img, _click_position_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, _open_lista_dados, _where_to_start, _barra_de_status

e_dir = Path('dados')


def cria_dados(pasta_arquivos):
    planilha_de_dados = p.confirm(buttons=['Criar planilha com os arquivos na pasta selecionada', 'Selecionar planilha já existente'])
    if planilha_de_dados == 'Criar planilha com os arquivos na pasta selecionada':
        try:
            for dado in os.listdir(e_dir):
                os.remove(os.path.join(e_dir, dado))
        except:
            pass
        
        arq_name = []
        for arq in os.listdir(pasta_arquivos):
            if arq.endswith('.RFB'):
                with open(os.path.join(pasta_arquivos, arq), 'r') as arquivo:
                    # Leia o conteúdo do arquivo
                    conteudo = arquivo.read()
                    cnpj = re.compile(r'DCTFM\s+\d\d\d\d\d\d\d\d\d(\d\d\d\d\d\d\d\d\d\d\d\d\d\d)\d\d\d\d').search(conteudo).group(1)
                
                names = arq.split('_')
                codigo = str(names[1]).replace('.RFB', '')
                arq_name.append((cnpj, codigo))
            
        arq_name = sorted(arq_name)
        for name in arq_name:
            _escreve_relatorio_csv(f'{name[0]};DCTF_{name[1]}.RFB')
            
        dados = _open_lista_dados(os.path.join(e_dir, 'dados.csv'))
    else:
        dados = _open_lista_dados()

    return dados
    
    
def abrir_dctf_mensal(event, pasta_arquivos):
    pasta = pasta_arquivos.split('/')
    unidade = pasta[0].replace(':', '')
    
    while not _find_img('tela_inicial.png', conf=0.9):
        time.sleep(1)

    _click_img('tela_inicial.png', conf=0.9)

    time.sleep(1)
    while not _find_img('origem_dos_dados.png', conf=0.9):
        p.hotkey('ctrl', 'm')
        if event == '-encerrar-':
            return
        time.sleep(1)

    for imagen in ['unidade_nao_selecionada.png', 'unidade_nao_selecionada_2.png', 'unidade_nao_selecionada_3.png']:
        if _find_img(imagen, conf=0.9):
            _click_img(imagen, conf=0.9, clicks=2)

    if event == '-encerrar-':
        return

    p.press(unidade)
    time.sleep(0.5)

    if event == '-encerrar-':
        return

    p.press('tab')
    time.sleep(0.5)

    if event == '-encerrar-':
        return

    for count, subpasta in enumerate(pasta):
        if event == '-encerrar-':
            return
        if subpasta != pasta[0]:
            p.write(subpasta)
            time.sleep(2)
            _click_position_img('pasta_selecionada_2.png', '+', pixels_x=15, conf=0.9, clicks=2)
            time.sleep(1)

            
def abrir_arquivo(event, arquivo, empresa, pasta_arquivos):
    while not _find_img('nome_arquivo.png', conf=0.9):
        if event == '-encerrar-':
            return ''
        time.sleep(1)
    
    # clica na barra para digitar o nome do arquivo
    _click_position_img('nome_arquivo.png', '+', pixels_y=17, conf=0.9, clicks=2)
    time.sleep(0.5)
    
    # escreve o nome do arquivo
    p.write(empresa[1])
    time.sleep(0.5)
    
    # aperta ok
    p.hotkey('alt', 'o')
    
    auxiliar = ''
    while not _find_img('declaracao_importada.png', conf=0.9):
        if _find_img('arquivo_invalido.png', conf=0.9):
            p.press('enter')
            return 'Arquivo inválido, ou não encontrado'
        if _find_img('declaracao_ja_importada.png', conf=0.9):
            p.press('enter')
            auxiliar = ', Declaração já importada'
        if _find_img('declaracao_nao_importada.png', conf=0.9):
            p.press('enter')
            _wait_img('imprimir_relatorio.png', conf=0.9)
            p.hotkey('alt', 'i')
            if not imprimir_relatorio(event, empresa, pasta_arquivos):
                return ''
            p.hotkey('alt', 'f')
            move_arquivo_usado(pasta_arquivos, 'Arquivos Com Erros', arquivo)
            return 'Erro ao importar a Declaração, relatório de erros salvo'
            
    p.press('enter')
    move_arquivo_usado(pasta_arquivos, 'Arquivos Importados', arquivo)
    return f'Arquivo importado com sucesso{auxiliar}'


def imprimir_relatorio(event, empresa, pasta_arquivos):
    pasta_relatorio = os.path.join(pasta_arquivos, 'Relatórios de erro').replace('//', '/')
    # cria uma pasta com o nome do relatório para salvar os arquivos
    os.makedirs(pasta_relatorio, exist_ok=True)

    cnpj, arquivo = empresa
    arquivo = arquivo.replace('.RFB', '')
    
    while not _find_img('salvar_como.png', conf=0.9):
        if event == '-encerrar-':
            return False
    # exemplo: cnpj;DAS;01;2021;22-02-2021;Guia do MEI 01-2021
    p.write(f'{cnpj} - {arquivo} - Relatórios de erro ')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    # Selecionar local
    p.press('tab', presses=6)
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    p.press('enter')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    pyperclip.copy(pasta_relatorio)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    p.press('enter')
    time.sleep(0.5)
    
    if event == '-encerrar-':
        return False
    
    p.hotkey('alt', 'l')
    time.sleep(1)
    
    if event == '-encerrar-':
        return False
    
    if _find_img('substituir.png', conf=0.9):
        p.press('s')
    time.sleep(1)
    
    return True
    

def cria_copia_de_seguranca():
    os.makedirs('C:\Cópias DCTF', exist_ok=True)

    # fecha a janela de importação
    p.hotkey('alt', 'c')
    time.sleep(1)
    
    # abre o dropdown de ferramentas
    p.hotkey('alt', 'f')
    time.sleep(1)
    
    # abre a tela para salvar cópia de segurança
    p.press('s')
    time.sleep(1)

    # espera a tela abrir
    while not _find_img('tela_salvar_copia.png', conf=0.9):
        time.sleep(1)

    # clica para selecionar a pasta de backup
    _click_position_img('unidade_c_backup.png', '+', pixels_y=17, conf=0.99, clicks=2)
    time.sleep(1)

    # seleciona a unidade c
    p.press('c')

    # clica para selecionar a pasta de backup
    _click_position_img('pasta_c_backup.png', '+', pixels_y=17, conf=0.99)
    time.sleep(1)

    # digita o nome da pasta
    p.write('Cópias DCTF')
    time.sleep(1)

    _click_img('copias_dctf.png', conf=0.99, clicks=2)
    time.sleep(1)

    # seleciona todos os arquivos
    p.hotkey('alt', 't')
    time.sleep(1)

    # confirma
    p.hotkey('alt', 'o')
    time.sleep(1)

    # espera a tela abrir
    while not _find_img('backup_criado.png', conf=0.9):
        if _find_img('substituir_backup.png', conf=0.9):
            p.press('enter')

    time.sleep(1)

    p.press('enter')
    time.sleep(2)
    # confirma
    p.hotkey('alt', 'c')
    time.sleep(2)


def move_arquivo_usado(pasta_arquivos, local, arquivo):
    pasta_arquivos_importados = os.path.join(pasta_arquivos, local)
    os.makedirs(pasta_arquivos_importados, exist_ok=True)

    shutil.move(os.path.join(pasta_arquivos, arquivo), os.path.join(pasta_arquivos_importados, arquivo))


@_barra_de_status
def run(window, event):
    andamentos = 'Importar Arquivos DCTF Mensal'
    
    total_empresas = empresas[:]
    
    abrir_dctf_mensal(event, pasta_arquivos)
    
    window['-titulo-'].update('Rotina automática, não interfira.')
    for count, empresa in enumerate(empresas[:], start=1):
        # printa o indice da empresa que está sendo executada
        
        while True:
            try:
                window['-Mensagens-'].update(f'{str(count - 1)} de {str(len(total_empresas))} | {str((len(total_empresas)) - count + 1)} Restantes')
                break
            except:
                pass
            
        cnpj, arquivo = empresa
        resultado = abrir_arquivo(event, arquivo, empresa, pasta_arquivos)
        if resultado != '':
            _escreve_relatorio_csv(f'{cnpj};{arquivo};{resultado}', nome=andamentos, local=pasta_arquivos)

        if count % 200 == 0:
            # a cada 200 arquivos, faz um backup
            cria_copia_de_seguranca()
            abrir_dctf_mensal(event, pasta_arquivos)
            
        if event == '-encerrar-':
            return

    cria_copia_de_seguranca()


if __name__ == '__main__':
    empresas = None
    while True:
        pasta_arquivos = _ask_for_dir()
        if not pasta_arquivos:
            break

        empresas = cria_dados(pasta_arquivos)

        if empresas:
            break

        p.alert(text='Essa pasta não contem arquivos ".RFB"')

    if empresas:
        run()

    p.alert(text='Execução finalizada.')

# -*- coding: utf-8 -*-
import time, fitz, sys, shutil, re, os, traceback, requests.exceptions, pdfkit, concurrent.futures
from xhtml2pdf import pisa
from requests import Session
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from anticaptchaofficial.hcaptchaproxyless import *
from datetime import datetime

import comum
from comum import controle_rotinas, id_execucao, update_window_elements


# captura o caminho absoluto do arquivo atual para criar uma pasta onde será feito o download dos arquivos para depois movê-los para a pasta definitiva
if getattr(sys, 'frozen', False):
    # Se estiver em um executável criado por PyInstaller
    current_file_path = sys.executable
else:
    # Se estiver em um script Python normal
    current_file_path = os.path.abspath(__file__)
current_file_path_list = current_file_path.split('\\')[:-1]
new_current_file_path = ''
for caminho in current_file_path_list:
    new_current_file_path = new_current_file_path + caminho + '\\'
pasta_final_download = new_current_file_path + f'Downloads {id_execucao}'


def run_cndtni(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def abre_pagina_consulta(window_principal, driver):
        return driver, 'resultado'
    
    def consulta_cndni(window_principal, driver, nome, cnpj, pasta_final_download_cndtni, pasta_final):
        return driver, 'resultado'
    
    def renomeia_cndni(window_principal, nome, cnpj, pasta_final_download_cndtni, pasta_final):
        return 'resultado'
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade', 'PostoFiscalUsuario', 'PostoFiscalSenha', 'PostoFiscalContador'],
                                                         filtrar_celulas_em_branco=['CNPJ', 'Razao', 'Cidade'])
    if not total_empresas:
        return False, pasta_final
    
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    driver = ''
    pasta_final_download_cndtni = pasta_final_download + f' {andamentos}'
    
    # configura o navegador
    options = comum.configura_navegador(window_principal, pasta=pasta_final_download_cndtni, retorna_options=True)
    
    tempos = [datetime.now()]
    tempo_execucao = []
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade, usuario, senha, perfil = empresa
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        nome = nome.replace('/', '')
        print(cnpj)
        
        resultado = 'ok'
        while True:
            # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
            if usuario_anterior != usuario:
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'ENCERRAR':
                    try:
                        driver.close()
                    except:
                        pass
                    return False, pasta_final
                
                # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
                try:
                    driver.close()
                except:
                    pass
                
                contador = 0
                while True:
                    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                    if cr == 'ENCERRAR':
                        try:
                            driver.close()
                        except:
                            pass
                        return False, pasta_final
                    try:
                        # abre uma nova sessão no site da fazenda
                        driver, sid = comum.new_session_fazenda_driver(window_principal, usuario, senha, perfil, retorna_driver=True, options=options)
                        if sid == 'erro_captcha':
                            return False, pasta_final
                        if sid == 'Falha na resolução do CAPTCHA.' or sid == 'Erro ao logar no perfil':
                            print('❗ Erro ao logar na empresa, tentando novamente')
                            continue
                            
                        break
                    except TimeoutException:
                        print('O site demorou muito pra responder...')
                    except:
                        print('❗ Erro ao logar na empresa, tentando novamente')
                    contador += 1
                    
                    if contador >= 5:
                        print('❌ Impossível de logar com esse usuário')
                        sid = 'Impossível de logar com esse usuário'
                        driver = 'erro'
                        break
                
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'ENCERRAR':
                    try:
                        driver.close()
                    except:
                        pass
                    return False, pasta_final
                
                if driver == 'erro':
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': sid},
                                                 nome=andamentos, local=pasta_final)
                    usuario_anterior = usuario
                    break
                
                else:
                    print('Driver:', driver, 'sID:', sid)
                    driver, resultado = abre_pagina_consulta(window_principal, driver)
            
            # se o resultado da abertura da página de consulta for 'ok', consulta a empresa
            if resultado == 'ok':
                driver, resultado = consulta_cndni(window_principal, driver, nome, cnpj, pasta_final_download_cndtni, pasta_final)
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'ENCERRAR':
                    try:
                        driver.close()
                    except:
                        pass
                    return False, pasta_final
                
                if resultado != 'erro':
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': resultado},
                                                 nome=andamentos, local=pasta_final)
                    usuario_anterior = usuario
                    break
                
                try:
                    driver.close()
                except:
                    pass
                usuario_anterior = 'padrão'
                continue
            
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'ENCERRAR':
                try:
                    driver.close()
                except:
                    pass
                return False, pasta_final
    
    try:
        driver.close()
    except:
        pass
    os.rmdir(pasta_final_download_cndtni)
    return True, pasta_final


def run_debitos_municipais_jundiai(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def login_jundiai(window_principal, cnpj, insc_muni, pasta_final):
        return driver, 'ok', 'situacao_print'
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'InsMunicipal', 'Razao', 'Cidade'], colunas_filtro=['Cidade'], palavras_filtro=['Jundiaí'])
    if not total_empresas:
        return False, pasta_final
    
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        contador = 0
        # iteração para logar na empresa e consultar, tenta 5 vezes
        while True:
            update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Entrando no perfil da empresa...'}})
            driver, situacao, situacao_print = login_jundiai(window_principal, cnpj, insc_muni, pasta_final)
            print(situacao_print)
            
            if situacao == 'recomeçar':
                try:
                    driver.close()
                except:
                    pass
                contador += 1
                if contador >= 5:
                    break
                continue
            
            else:
                break
        
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            return False, pasta_final
        
        comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': situacao_print[2:]},
                                     nome=andamentos, local=pasta_final)
    
    return True, pasta_final


def run_debitos_municipais_valinhos(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def reset_table(table):
        return table
    
    def format_data(content):
        return str(new_soup)
    
    def login_valinhos(window_principal, driver, cnpj, insc_muni):
            return False, 'ok'
    
    def consulta_valinhos(driver, cnpj, pasta_final):
            return driver, 'resultado'
        
    def consulta_certidao_valinhos(window_principal, driver, cnpj, insc_muni, nome, pasta_final_download_certidao_valinhos, pasta_final):
        return driver, f'Certidão Negativa salva - {situacao}'
        
    def analisa_certidao(arquivo_baixado):
            return 'resultado'
        
        
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'InsMunicipal', 'Razao', 'Cidade'], colunas_filtro=['Cidade'], palavras_filtro=['Valinhos'])
    if not total_empresas:
        return False, pasta_final
    
    pasta_final_download_certidao_valinhos = pasta_final_download + f' {andamentos}'
    
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        try:
            insc_muni = int(insc_muni)
        except:
            continue
        
        while True:
            status, driver = comum.configura_navegador(window_principal, pasta=pasta_final_download_certidao_valinhos, desabilita_seguranca=True)
            update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Entrando no perfil da empresa...'}})
            driver, resultado = login_valinhos(window_principal, driver, cnpj, insc_muni)
            
            if resultado == 'encerrar':
                break
            if resultado != 'recomeçar':
                if resultado == 'ok':
                    driver, resultado = consulta_valinhos(driver, cnpj, pasta_final)
                    # volta para a tela inicial
                    button = driver.find_element(by=By.ID, value='_CBoTdImgCancel')
                    button.click()
                    while not comum.find_by_id('span9Menu', driver):
                        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                        if cr == 'ENCERRAR':
                            return driver, 'encerrar'
                        time.sleep(1)
                    
                    update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Consultando a Certidão Negativa'}})
                    while True:
                        driver, resultado_cert = consulta_certidao_valinhos(window_principal, driver, cnpj, insc_muni, nome, pasta_final_download_certidao_valinhos, pasta_final)
                        if resultado_cert == 'recomeçar':
                            driver.refresh()
                        else:
                            break
                    
                    if resultado_cert != 'encerrar':
                        comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': resultado, 'CERTIDÃO NEGATIVA': resultado_cert},
                                                     nome=andamentos, local=pasta_final)
                else:
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': resultado, 'CERTIDÃO NEGATIVA': ''},
                                                 nome=andamentos, local=pasta_final)
                break
            
            update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Erro ao logar na empresa, tentando novamente...'}})
        
        try:
            driver.close()
        except:
            pass
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            return False, pasta_final
    
    return True, pasta_final


def run_debitos_municipais_vinhedo(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def pesquisar_vinhedo(window_principal, cnpj, insc_muni, pasta_final_certidao):
            return False, 'ok'
    
    def salvar_guia_vinhedo(window_principal, driver, cnpj, pasta_final_certidao):
        return 'Certidão negativa salva'
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'InsMunicipal', 'Razao', 'Cidade'], colunas_filtro=['Cidade'], palavras_filtro=['Vinhedo'])
    if not total_empresas:
        return False, pasta_final
        
    # configura o navegador
    pasta_final_certidao = os.path.join(pasta_final, 'Certidões')
    os.makedirs(pasta_final_certidao, exist_ok=True)
    
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, insc_muni, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        insc_muni = insc_muni.replace('/', '').replace('.', '').replace('-', '')
        
        while True:
            update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Entrando no perfil da empresa...'}})
            
            situacao, situacao_print = pesquisar_vinhedo(window_principal, cnpj, insc_muni, pasta_final_certidao)
            if situacao_print == 'encerrar':
                break
                
            if situacao_print != 'recomeçar':
                print(situacao_print)
                if situacao == 'Desculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.':
                    window_principal.write_event_value('-run_debitos_municipais_vinhedo_alerta_1-', 'alerta')
                    #alert(f'❗ Rotina "{andamentos}":\n\nDesculpe, mas ocorreram problemas de rede. Por favor, tente novamente mais tarde.')
                    return False, pasta_final
                
                comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'INSCRIÇÃO MUNICIPAL': insc_muni, 'NOME': nome, 'RESULTADO': situacao},
                                             nome=andamentos, local=pasta_final)
                break
            
            update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Erro ao logar na empresa, tentando novamente...'}})
            
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            return False, pasta_final
        
    return True, pasta_final


def run_debitos_estaduais(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    situacoes_debitos_estaduais = {
        'C': '✔ Certidão sem pendencias',
        'S': '❗ Nao apresentou STDA',
        'G': '❗ Nao apresentou GIA',
        'E': '❌ Nao baixou arquivo',
        'T': '❗ Transporte de Saldo Credor Incorreto',
        'P': '❗ Pendencias',
        'I': '❗ Pendencias GIA',
    }
    
    def confere_pendencias(pagina):
        return ' e '.join(situacao)
    
    def consulta_deb_estaduais(window_principal, pasta_final, cnpj, cidade, s, s_id):
        return situacao
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade', 'PostoFiscalUsuario', 'PostoFiscalSenha', 'PostoFiscalContador'],
                                                         filtrar_celulas_em_branco=['CNPJ', 'Razao', 'Cidade'])
    if not total_empresas:
        return False, pasta_final
    
    # inicia a variável que verifica se o usuário da execução anterior é igual ao atual
    usuario_anterior = 'padrão'
    s = False
    
    tempos = [datetime.now()]
    tempo_execucao = []
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade, usuario, senha, perfil = empresa
        cnpj = comum.concatena(cnpj, 14, 'antes', 0)
        cnpj = str(cnpj)
        print(cnpj)
        
        while True:
            # verifica se o usuario anterior é o mesmo para não fazer login de novo com o mesmo usuário
            if usuario_anterior != usuario:
                # se o usuario anterior for diferente e existir uma sessão aberta, a sessão é fechada
                if s: s.close()
                
                # abre uma nova sessão
                s = Session()
                
                contador = 0
                # loga no site da secretaria da fazenda com web driver e salva os cookies do site e a id da sessão
                while True:
                    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                    if cr == 'ENCERRAR':
                        if s: s.close()
                        return False, pasta_final
                    
                    if contador >= 3:
                        cookies = 'erro'
                        sid = 'Erro ao logar na empresa'
                        break
                    try:
                        cookies, sid = comum.new_session_fazenda_driver(window_principal, usuario, senha, perfil)
                        if sid == 'erro_captcha':
                            return False, pasta_final
                        if sid == 'Falha na resolução do CAPTCHA.':
                            print('❗ Erro ao resolver o captcha, tentando novamente')
                            continue
                        break
                    except TimeoutException:
                        print('O site demorou muito para responder...')
                    except:
                        print('❗ Erro ao logar na empresa, tentando novamente')
                        contador += 1
                
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'ENCERRAR':
                    if s: s.close()
                    return False, pasta_final
                
                time.sleep(1)
                # se não salvar os cookies vai para o próximo dado
                if cookies == 'erro' or not cookies:
                    texto = sid
                    usuario_anterior = 'padrão'
                    
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': texto},
                                                 nome=andamentos, local=pasta_final)
                    continue
                
                # adiciona os cookies do login da sessão por request no web driver
                for cookie in cookies:
                    s.cookies.set(cookie['name'], cookie['value'])
            
            # se não retornar a id da sessão do web driver fecha a sessão por request
            if not sid:
                situacao = '❌ Erro no login'
                usuario_anterior = 'padrão'
                s.close()
                break
                
            # se retornar a id da sessão do web driver executa a consulta
            else:
                # retorna o resultado da consulta
                situacao = consulta_deb_estaduais(window_principal, pasta_final, cnpj, cidade, s, sid)
                if situacao == 'erro':
                    update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Erro ao coletar as informações, tentando novamente...'}})
                    continue
                # guarda o usuario da execução atual
                usuario_anterior = usuario
                
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'ENCERRAR':
                    if s: s.close()
                    return False, pasta_final
                break
                
        # escreve na planilha de andamentos o resultado da execução atual
        texto = f'{cnpj};{str(situacao[2:])}'
        try:
            texto = texto.replace('❗ ', '').replace('❌ ', '').replace('✔ ', '')
            comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'NOME': nome, 'CIDADE': cidade, 'RESULTADO': texto},
                                        nome=andamentos, local=pasta_final)
        except:
            raise Exception(f"Erro ao escrever esse texto: {texto}")
        print(situacao)
        
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            if s: s.close()
            return False, pasta_final
    
    if s: s.close()
    return True, pasta_final


def run_debitos_divida_ativa_uniao(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    ultima_divida = 'Log/ultima_divida.txt'
    
    def click(driver, divida):
        # função para clicar em elementos via ID
        print('>>> Baixando arquivo')
        contador_2 = 0
        while True:
            try:
                driver.find_element(by=By.ID, value=divida).click()
                break
            except:
                contador_2 += 1
            
            if contador_2 > 5:
                return False
        
        return True
    
    def download_divida(driver, divida, descricao, tipo, processo, inscricao, andamentos, pasta_final_download_dau, pasta_final_dividas):
        return driver, 'ok', False
    
    def converte_html_pdf(download_folder, pasta_final_dividas, arquivo, descricao, tipo, processo, inscricao, andamentos):
        return True
    
    def pega_info_arquivo(html_path, descricao):
        return nome_pdf, descricao, nome_inteiro, cpf_cnpj
    
    def login_sieg(driver):
        return driver
    
    def sieg_iris(driver):
        return driver
    
    def consulta_lista(window_principal, driver, andamentos, pasta_final_download_dau, pasta_final_dividas):
        return driver
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                                            colunas_usadas=False, colunas_filtro=False, palavras_filtro=False)
    if continuar_rotina == '-reiniciar_rotina-':
        with open(ultima_divida, 'w', encoding='utf-8') as f:
            f.write('1')
    
    pasta_final_dividas = os.path.join(pasta_final, 'Dividas')
    pasta_final_download_dau = pasta_final_download + f' {andamentos}'
    
    # iniciar o driver do chome
    while True:
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            break
        
        status, driver = comum.configura_navegador(window_principal, pasta=pasta_final_download_dau, timeout=60)
        driver = login_sieg(driver)
        if driver:
            driver = sieg_iris(driver)
            if driver:
                driver = consulta_lista(window_principal, driver, andamentos, pasta_final_download_dau, pasta_final_dividas)
                if driver:
                    driver.close()
                    break
        
        update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'SIEG IRIS demorou muito pra responder, tentando novamente'}})
        print('SIEG IRIS demorou muito pra responder, tentando novamente')
    
    cr = open(controle_rotinas, 'r', encoding='utf-8').read()
    if cr == 'ENCERRAR':
        return False, pasta_final
    
    if continuar_rotina == '-reiniciar_rotina-':
        with open(ultima_divida, 'w', encoding='utf-8') as f:
            f.write('1')
    
    os.rmdir(pasta_final_download_dau)
    return True, pasta_final


def run_debitos_divida_ativa(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def verifica_debitos(pagina):
        return 'ok'
            
    def limpa_registros(html):
        return html
    
    def salva_pagina(pagina, cnpj, nome, andamentos, pasta_final, compl=''):
        return True
    
    def inicia_sessao(window_principal):
        return True, 'ok'
    
    def consulta_debito(window_principal, s, nome_botao_consulta, cnpj_formatado, nome, andamentos, pasta_final):
       return s
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade'], colunas_filtro=False, palavras_filtro=False)
    if not total_empresas:
        return False, pasta_final
    
    cookies = ''
    nova_sessao = True
    nome_botao_consulta = None
    for count, [index_atual, empresa] in enumerate(df_empresas.iloc[index:].iterrows(), start=1):
        # printa o índice da empresa que está sendo executada
        tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
        
        cnpj, nome, cidade = empresa
        cnpj = str(cnpj)
        cnpj_formatado = comum.concatena(cnpj, 14, 'antes', 0)
        
        # enquanto a consulta não conseguir ser realizada e tenta de novo
        tentativas = 0
        while True:
            if nova_sessao:
                # inicia uma sessão com webdriver para gerar cookies no site e garantir que as requisições funcionem depois
                contador = 0
                while True:
                    driver, nome_botao_consulta = inicia_sessao(window_principal)
                    if nome_botao_consulta == 'captcha_inativo':
                        return False, pasta_final
                    if nome_botao_consulta == 'erro_captcha':
                        update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Erro ao iniciar a sessão, tentando novamente'}})
                        print('Erro ao iniciar a sessão, tentando novamente')
                    
                    elif nome_botao_consulta == 'erro_site':
                        update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'O site demorou muito pra responder, tentando novamente'}})
                        print('O site demorou muito pra responder, tentando novamente')
                        contador += 1
                    if contador > 5:
                        window_principal.write_event_value('-run_debitos_divida_ativa_alerta_1-', 'alerta')
                        # alert(f'❌ O site "https://www.dividaativa.pge.sp.gov.br/sc/pages/consultas/consultarDebito.jsf" não respondeu conforme o esperado, consulta encerrada.')
                        return False, pasta_final
                    
                    else:
                        break
                
                # armazena os cookies gerados pelo webdriver
                cookies = driver.get_cookies()
                driver.quit()
                nova_sessao = False
                
            # inicia uma sessão para as requisições
            with Session() as s:
                # pega a lista de cookies armazenados e os configura na sessão
                for cookie in cookies:
                    s.cookies.set(cookie['name'], cookie['value'])
                cr = open(controle_rotinas, 'r', encoding='utf-8').read()
                if cr == 'ENCERRAR':
                    break
                    
                if tentativas >= 5:
                    comum.escreve_relatorio_xlsx({'CNPJ': cnpj_formatado, 'NOME': nome, 'RESULTADO': 'Empresa com débitos', 'COMPLEMENTO': 'Não foi possível capturar os dados da dívida. O número de tentativas foi excedido.'},
                                                 nome=andamentos, local=pasta_final)
                    break
                
                try:
                    s = consulta_debito(window_principal, s, nome_botao_consulta, cnpj_formatado, nome, andamentos, pasta_final)
                    if s == 'erro':
                        continue
                    if s == 'erro_captcha':
                        return False, pasta_final
                    break
                except requests.exceptions.ConnectionError:
                    tentativas += 1
                    print(f'>>> Erro de conexão com o site, tentando novamente... {tentativas}/5')
                    update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'Erro de conexão com o site, tentando novamente...'}})
                    nova_sessao = True
                    pass
                except requests.exceptions.ReadTimeout:
                    tentativas += 1
                    print(f'>>> O site demorou muito para responder, tentando novamente... {tentativas}/5')
                    update_window_elements(window_principal, {'-Mensagens_2-': {'value': 'O site demorou muito para responder, tentando novamente...'}})
                    nova_sessao = True
                    pass
            
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            return False, pasta_final
    
    return True, pasta_final


def run_pendencias_sigiss(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos, tempos, tempo_execucao, continuar_rotina):
    def divide_list(lst, n):
        # Divide a lista em n partes
        k, m = divmod(len(lst), n)
        return [lst[i * k + min(i, m):(i + 1) * k + min(i + 1, m)] for i in range(n)]
    
    def login_sigiss(cnpj, senha):
        return True, '', s
    
    def consulta(s, cnpj, nome, pasta_final):
        return True, documento, s
    
    # filtrar e criar a nova planilha de dados
    pasta_final, index, df_empresas, total_empresas = comum.configura_dados(window_principal, continuar_rotina, codigo_20000, situacao_dados, planilha_dados, pasta_final_, andamentos,
                                                         colunas_usadas=['CNPJ', 'Razao', 'Cidade', 'Senha Prefeitura'], colunas_filtro=['Cidade'], palavras_filtro=['Valinhos'])
    if not total_empresas:
        return False, pasta_final
        
    # Divida a lista principal em 3 partes
    sublists = divide_list(df_empresas[index:].values.tolist(), 6)
    
    update_window_elements(window_principal, {'-Mensagens_2-': {'value': ''}})

    def process_sublist(sublists):
        tempos = [datetime.now()]
        tempo_execucao = []
        for count, empresa in enumerate(sublists, start=1):
            # printa o índice da empresa que está sendo executada
            tempos, tempo_execucao = comum.indice(count, empresa, total_empresas, index, window_principal, tempos, tempo_execucao)
            
            cnpj, nome, cidade, senha = empresa
            cnpj = str(cnpj)
            cnpj = comum.concatena(cnpj, 14, 'antes', 0)
            print(cnpj)
            
            nome = nome.replace('/', ' ')
            
            # faz login no SIGISSWEB
            execucao, documento, s = login_sigiss(cnpj, senha)
            if execucao:
                # se fizer login, consulta a situação da empresa
                execucao, documento, s = consulta(s, cnpj, nome, pasta_final)
            
            # escreve os resultados da consulta
            comum.escreve_relatorio_xlsx({'CNPJ': cnpj, 'SENHA': senha, 'NOME':nome, 'RESULTADO':documento}, nome=andamentos, local=pasta_final)
            s.close()
            
            cr = open(controle_rotinas, 'r', encoding='utf-8').read()
            if cr == 'ENCERRAR':
                return False, pasta_final
        
    # Use ThreadPoolExecutor para processamento paralelo
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # Mapeie a função para cada sublista
        list(executor.map(process_sublist, sublists))
    
    return True, pasta_final


def executa_rotina(window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final, rotina, continuar_rotina):
    rotinas = {'Consulta Certidão Negativa de Débitos Tributários Não Inscritos': run_cndtni,                     # ok
               'Consulta Débitos Municipais Jundiaí': run_debitos_municipais_jundiai,                             # ok
               'Consulta Débitos Municipais Valinhos': run_debitos_municipais_valinhos,                           # ok
               'Consulta Débitos Municipais Vinhedo': run_debitos_municipais_vinhedo,                             # ok
               'Consulta Débitos Estaduais - Situação do Contribuinte': run_debitos_estaduais,                    # ok
               'Consulta Dívida Ativa da União': run_debitos_divida_ativa_uniao,                                  # ok
               'Consulta Dívida Ativa Procuradoria Geral do Estado de São Paulo': run_debitos_divida_ativa,       # ok
               'Consulta Pendências SIGISSWEB Valinhos': run_pendencias_sigiss                                    # ok
               }
    
    tempos = [datetime.now()]
    tempo_execucao = []

    #try:
    resultado, pasta_final_oficial = rotinas[str(rotina)](window_principal, codigo_20000, situacao_dados, planilha_dados, pasta_final, rotina, tempos, tempo_execucao, continuar_rotina)
    """except ValueError:
        window_principal.write_event_value('-erro_planilha_de_dados-', 'alerta')
        return"""
    
    updates = {'-Mensagens_2-': {'value': ''},
               '-Mensagens-': {'value': ''}}
    update_window_elements(window_principal, updates)
    if resultado:
        updates = {'-progressbar-': {'value': 100, 'max': 100},
                   '-Progresso_texto-': {'value': '100%'}}
        update_window_elements(window_principal, updates)
        window_principal.write_event_value('-FINALIZADO-', 'alerta')
        # alert(f'✔ Rotina "{rotina}" finalizada')
    else:
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            window_principal.write_event_value('-ENCERRADO-', 'alerta')
        try:
            os.remove(os.path.join(pasta_final_oficial, 'Dados.xlsx'))
            os.rmdir(pasta_final_oficial)
        except:
            pass
    
    update_window_elements(window_principal, {'-progressbar-': {'value': 0, 'max': 100}})
    
import re, os, shutil, pandas as pd
from datetime import datetime

import comum

controle_rotinas = 'Log/Controle.txt'


def executa_analisador(window_principal, pasta_arquivos, pasta_final):
    index = 0
    tempos = [datetime.now()]
    tempo_execucao = []
    nova_pasta = 'Nova pasta'
    total_arquivos = len(os.listdir(pasta_arquivos))
    for count, arquivo in enumerate(os.listdir(pasta_arquivos), start=1):
        cr = open(controle_rotinas, 'r', encoding='utf-8').read()
        if cr == 'ENCERRAR':
            window_principal.write_event_value('-ENCERRADO-', 'alerta')
            return
        
        caminho_do_arquivo = os.path.join(pasta_arquivos, arquivo)
        arquivo_selecionado = f'Arquivo selecionado: {caminho_do_arquivo}'
        
        tempos, tempo_execucao = comum.indice(count, arquivo_selecionado, total_arquivos, index, window_principal, tempos, tempo_execucao)
        
        if not arquivo.endswith('.DEC'):
            continue
            
        conteudo_total = ''
        codigos = 'Possuí lançamento do código '
        with open(caminho_do_arquivo, 'r') as arquivo_dec:
            # Leia o conteúdo do arquivo
            conteudo = arquivo_dec.readlines()
            for linha in conteudo:
                conteudo_total += linha
            #print('Conteúdo: ', conteudo_total)
            
        dados = re.compile(r'IRPF\s+\d+(\d\d\d\d\d\d\d\d\d\d\d)\s+\d\d\d\d([A-Z\s]+)\s+').search(conteudo_total)
        cpf = dados.group(1)
        nome = dados.group(2).strip()
        
        lancamento = re.compile(r'23(\d\d\d\d\d\d\d\d\d\d\d)0014\d{23}').search(conteudo_total)
        if lancamento:
            codigos += '14 - Herança'
            
        lancamento = re.compile(r'23(\d\d\d\d\d\d\d\d\d\d\d)0019\d{23}').search(conteudo_total)
        if lancamento:
            codigos = codigos.replace('14 - Herança', '14 - Herança e ')
            codigos += '19 - Meação'
        
        if codigos == 'Possuí lançamento do código ':
            codigos = 'Não possuí os códigos de lançamento buscados'
        else:
            if codigos == 'Possuí lançamento do código 14 - Herança':
                nova_pasta = 'Arquivos com Herança'
                
            elif codigos == 'Possuí lançamento do código 19 - Meação':
                nova_pasta = 'Arquivos com Meação'
                
            elif codigos == 'Possuí lançamento do código 14 - Herança e 19 - Meação':
                nova_pasta = 'Arquivos com Herança e Meação'
                
            pasta_final_arquivos = os.path.join(pasta_final, nova_pasta)
            caminho_final_arquivo = os.path.join(pasta_final, nova_pasta, arquivo)
            os.makedirs(pasta_final_arquivos, exist_ok=True)
            shutil.move(caminho_do_arquivo, caminho_final_arquivo)
                
        comum.escreve_relatorio_xlsx({'CPF': cpf, 'NOME': nome, 'ARQUIVO': arquivo, 'SITUAÇÂO': codigos}, pasta_final, 'Analisa IRPF')
            
    window_principal.write_event_value('-FINALIZADO-', 'alerta')
    
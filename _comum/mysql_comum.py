from datetime import datetime

import mysql.connector

dados = "Dados DB.txt"
f = open(dados, 'r', encoding='utf-8')
senha = f.readline()


def conect_db_task_watch():
    """ Conecta oa banco de dados """
    # Conectar ao banco de dados MySQL
    conn = mysql.connector.connect(
        host="host",  # Seu host MySQL
        port="port",  # Porta
        user="user",  # Seu usuário MySQL
        password=senha,  # Sua senha MySQL
        database="banco_dados"  # Nome do banco de dados
    )
    
    return conn


def _create_tables_task_watch():
    """ Cria as tabelas das rotinas e do histórico delas """
    conn = conect_db_task_watch()
    cursor = conn.cursor()
    
    # Criar tabela para armazenar o status dos scripts
    cursor.execute('''CREATE TABLE IF NOT EXISTS status_scripts (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        script_name VARCHAR(255),
                        status VARCHAR(20),
                        data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        andamentos VARCHAR(255),
                        tempo_medio VARCHAR(255),
                        tempo_estimado VARCHAR(255),
                        previsao_termino VARCHAR(255),
                        porcentagem VARCHAR(8))''')
    conn.commit()
    
    # Criar tabela para armazenar o status dos scripts
    cursor.execute('''CREATE TABLE IF NOT EXISTS historico_rotinas (
                            id INT AUTO_INCREMENT PRIMARY KEY,
                            data_hora VARCHAR(18),
                            script_name VARCHAR(255),
                            andamentos VARCHAR(255))''')
    
    conn.commit()
    conn.close()
    print('Conexão com o banco de dados bem sucedida')


def _update_script_status(script_name, status, andamentos='', tempo_medio='', tempo_estimado='', previsao_termino='', porcentagem='', type=False, ):
    """ Atualiza ou cria uma nova linha referente a rotina em execução """
    conn = conect_db_task_watch()
    cursor = conn.cursor()
    
    # Verificar se o script já existe no banco de dados
    query = '''SELECT id FROM status_scripts WHERE script_name = %s'''
    cursor.execute(query, (script_name,))
    result_id = cursor.fetchone()

    # se encontrou a linha referente ao cartão
    if result_id:
        # apenas atualiza o timestemp
        if type == 'tempo':
            query = '''UPDATE status_scripts
                                   SET status = %s, data_hora = CURRENT_TIMESTAMP
                                   WHERE id = %s'''
            values = (status, result_id[0])
        
        # atualiza o cartão com alguma mensagem referente a etapa que o robô está
        elif type == 'ocorrencia':
            # Atualizar o status do script no banco de dados
            query = '''UPDATE status_scripts
                                   SET status = %s, data_hora = CURRENT_TIMESTAMP, tempo_medio = %s, tempo_estimado = %s, previsao_termino = %s
                                   WHERE id = %s'''
            values = (status, tempo_medio, tempo_estimado, previsao_termino, result_id[0])
        
        # Atualizar os andamentos do script no banco de dados
        else:
            query = '''UPDATE status_scripts
                           SET status = %s, data_hora = CURRENT_TIMESTAMP, andamentos = %s, tempo_medio = %s, tempo_estimado = %s, previsao_termino = %s, porcentagem = %s
                           WHERE id = %s'''
            values = (status, andamentos, tempo_medio, tempo_estimado, previsao_termino, porcentagem, result_id[0])
    
    # se não encontrar a linha correspondente, cria uma nova
    else:
        # Inserir o status do script no banco de dados
        query = '''INSERT INTO status_scripts (script_name, status, andamentos, tempo_medio, tempo_estimado, previsao_termino, porcentagem)
                   VALUES (%s, %s, %s, %s, %s, %s, %s)'''
        values = (script_name, status, andamentos, tempo_medio, tempo_estimado, previsao_termino, porcentagem)
    
    cursor.execute(query, values)
    conn.commit()
    conn.close()


def _update_historico_status(script_name, andamentos=''):
    """ Cria uma linha no histórico das ocorrências das rotinas """
    conn = conect_db_task_watch()
    cursor = conn.cursor()
    
    data_hora = datetime.now().strftime("%d/%m/%Y - %H:%M")
    
    # Inserir o status do script no banco de dados
    query = '''INSERT INTO historico_rotinas (data_hora, script_name, andamentos)
               VALUES (%s, %s, %s)'''
    values = (data_hora, script_name, andamentos)
    
    cursor.execute(query, values)
    conn.commit()
    conn.close()

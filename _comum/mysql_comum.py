import mysql.connector
dados = "Dados DB.txt"
f = open(dados, 'r', encoding='utf-8')
senha = f.readline()


def conect_db():
    # Conectar ao banco de dados MySQL
    conn = mysql.connector.connect(
            host="host", # Seu host MySQL
            port="1234", # Porta
            user="user", # Seu usuário MySQL
            password=senha, # Sua senha MySQL
            database="database" # Nome do banco de dados
        )
    
    return conn


def _conect_db():
    conn = conect_db()
    cursor = conn.cursor()
    
    # Criar tabela para armazenar o status dos scripts
    cursor.execute('''CREATE TABLE IF NOT EXISTS status_scripts (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        script_name VARCHAR(255),
                        status VARCHAR(255),
                        data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        andamentos VARCHAR(255),
                        tempo_estimado VARCHAR(255),
                        previsao_termino VARCHAR(255),
                        porcentagem VARCHAR(8))''')
    conn.commit()
    
    # Criar tabela para armazenar o status dos scripts
    cursor.execute('''CREATE TABLE IF NOT EXISTS historico_rotinas (
                            id INT AUTO_INCREMENT PRIMARY KEY,
                            data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                            script_name VARCHAR(255),
                            andamentos VARCHAR(255))''')
    
    conn.commit()
    conn.close()
    print('Conexão com o banco de dados bem sucedida')
    
    
def _update_script_status(script_name, status, andamentos='', tempo_estimado='', previsao_termino='', porcentagem='', type=False,):
    conn = conect_db()
    cursor = conn.cursor()
    
    # Verificar se o script já existe no banco de dados
    query = '''SELECT id FROM status_scripts WHERE script_name = %s'''
    cursor.execute(query, (script_name,))
    result_id = cursor.fetchone()
    #print('resultado da consulta',result_id)
    if result_id:
        if type == 'tempo':
            query = '''UPDATE status_scripts
                                   SET status = %s, data_hora = CURRENT_TIMESTAMP
                                   WHERE id = %s'''
            values = (status, result_id[0])
        
        elif type == 'ocorrencia':
            # Atualizar o status do script no banco de dados
            query = '''UPDATE status_scripts
                                   SET status = %s, data_hora = CURRENT_TIMESTAMP, tempo_estimado = %s, previsao_termino = %s
                                   WHERE id = %s'''
            values = (status, tempo_estimado, previsao_termino, result_id[0])
        else:
            # Atualizar o status do script no banco de dados
            query = '''UPDATE status_scripts
                           SET status = %s, data_hora = CURRENT_TIMESTAMP, andamentos = %s, tempo_estimado = %s, previsao_termino = %s, porcentagem = %s
                           WHERE id = %s'''
            values = (status, andamentos, tempo_estimado, previsao_termino, porcentagem, result_id[0])
    else:
        # Inserir o status do script no banco de dados
        query = '''INSERT INTO status_scripts (script_name, status, andamentos, tempo_estimado, previsao_termino, porcentagem)
                   VALUES (%s, %s, %s, %s, %s, %s)'''
        values = (script_name, status, andamentos, tempo_estimado, previsao_termino, porcentagem)
        
    cursor.execute(query, values)
    conn.commit()
    conn.close()


def _update_historico_status(script_name, andamentos=''):
    conn = conect_db()
    cursor = conn.cursor()

    # Inserir o status do script no banco de dados
    query = '''INSERT INTO historico_rotinas (script_name, andamentos)
               VALUES (%s, %s)'''
    values = (script_name, andamentos)
    
    cursor.execute(query, values)
    conn.commit()
    conn.close()

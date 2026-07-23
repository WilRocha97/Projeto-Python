import pyautogui as p
import pandas as pd
import math, os
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '_comum')))
from comum_comum import _ask_for_file


def dividir_xlsx_em_csvs(arquivo_xlsx, num_partes):
    """
    Divide uma planilha xlsx em múltiplos arquivos CSV.
    
    Args:
        arquivo_xlsx: caminho do arquivo .xlsx
        num_partes: quantidade de arquivos CSV desejados
    """
    # Lê a planilha
    df = pd.read_excel(arquivo_xlsx, dtype=str)
    
    total_linhas = len(df)
    
    if num_partes > total_linhas:
        p.alert(f'Número de planilhas "{str(num_partes)}" é maior que o número de linhas da planilha "{str(total_linhas)}"')
        return
    
    linhas_por_parte = math.ceil(total_linhas / num_partes)
    
    print(f"Total de linhas: {total_linhas}")
    print(f"Linhas por arquivo: ~{linhas_por_parte}")
    
    # Divide e salva cada parte
    for i in range(num_partes):
        inicio = i * linhas_por_parte
        fim = min((i + 1) * linhas_por_parte, total_linhas)
        
        if inicio >= total_linhas:
            break
            
        parte = df.iloc[inicio:fim]
        nome_saida = f"parte_{i + 1}.csv"
        parte.to_csv(nome_saida, sep=';', index=False, header=False)
        
        print(f"Criado: {nome_saida} ({len(parte)} linhas)")

# Exemplo de uso
if __name__ == "__main__":
    ftypes = [('Plain text files', '*.xlsx')]
    
    file = _ask_for_file(filetypes=ftypes)
    if file:
        quantidade = int(p.prompt(title='Script incrível', text='Quantas planilhas de dados deseja criar?'))                  # Quantidade de CSVs desejados
        
        
        dividir_xlsx_em_csvs(file, quantidade)
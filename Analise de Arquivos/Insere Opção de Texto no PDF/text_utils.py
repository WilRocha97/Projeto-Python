"""
Utilitários de texto: extração de CNPJ e substituição de placeholders
de data/hora.
"""

import re
from datetime import datetime


# Regex para CNPJ: aceita formatado (12.345.678/0001-99) e só dígitos
# (12345678000199), além de variações com -, _ ou nada como separador.
# Lookbehind/lookahead negativos evitam capturar parte de um número maior.
CNPJ_REGEX = re.compile(
    r'(?<!\d)(\d{2})[.\-_]?(\d{3})[.\-_]?(\d{3})[/\-_]?(\d{4})[\-_]?(\d{2})(?!\d)'
)


def extrair_cnpj(nome_arquivo):
    """
    Procura um CNPJ no nome do arquivo e devolve formatado XX.XXX.XXX/XXXX-XX.
    Retorna None se não encontrar.
    """
    m = CNPJ_REGEX.search(nome_arquivo)
    if not m:
        return None
    return f"{m.group(1)}.{m.group(2)}.{m.group(3)}/{m.group(4)}-{m.group(5)}"


def substituir_placeholders(texto):
    """Substitui {data}, {hora}, {data_hora} pelos valores atuais."""
    if not texto:
        return texto
    agora = datetime.now()
    return (
        texto
        .replace("{data_hora}", agora.strftime("%d/%m/%Y %H:%M"))
        .replace("{data}", agora.strftime("%d/%m/%Y"))
        .replace("{hora}", agora.strftime("%H:%M"))
    )

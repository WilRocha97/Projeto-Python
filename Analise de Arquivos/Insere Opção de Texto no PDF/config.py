"""
Configurações personalizáveis do "Inserir Texto em PDFs - V&P".

Toda customização que NÃO envolve código (textos, cores, tamanhos)
deve ficar centralizada aqui.
"""

import sys
from pathlib import Path

# =========================================================================
# OPÇÕES DE TEXTO
# =========================================================================

# Lista de opções de texto disponíveis para o usuário.
# Para adicionar uma nova opção, basta acrescentar uma linha:
#     "Rótulo exibido na lista": "Texto que será inserido no PDF",
# A data, hora e CNPJ são adicionados automaticamente pelos checkboxes da UI.
# Se quiser injetar a data/hora dentro do texto (em vez de no fim), use os
# placeholders {data}, {hora}, {data_hora}.
OPCOES_TEXTO = {
    "-": " ",
    "Guia ISS Prestado": "Guia ISS Prestado",
    "Guia ISS Tomado": "Guia ISS Tomado",
    "Nota Fiscal de Serviço Tomado": "Nota Fiscal de Serviço Tomado",
    "Nota Fiscal de Serviço Prestado": "Nota Fiscal de Serviço Prestado",
    "Nota Fiscal de Venda": "Nota Fiscal de Venda",
}


# Lista de textos EXTRAS opcionais, apresentados em um dropdown separado
# que concatena ao texto principal.
# Para adicionar/remover opções, basta editar este dicionário.
# A entrada "-" mapeada para string vazia significa "não anexar nada extra".
OPCOES_EXTRAS = {
    "-": "",
    "Primeira parcela": "Primeira parcela",
    "Segunda parcela": "Segunda parcela",
    "Terceira parcela": "Terceira parcela",
}


# =========================================================================
# SAÍDA
# =========================================================================

# Nome da subpasta de saída criada automaticamente quando o usuário
# não escolhe uma pasta de destino.
PASTA_SAIDA_PADRAO = "PDFs com texto"


# =========================================================================
# FONTE
# =========================================================================

# Se diferente de None, força esse tamanho em TODOS os PDFs (em pontos).
# Se None, detecta automaticamente o tamanho mais comum do PDF.
TAMANHO_FONTE_OVERRIDE = None     # Ex.: 9 para forçar 9pt sempre
TAMANHO_FONTE_PADRAO = 9          # Fallback se a detecção falhar


# =========================================================================
# MARGENS DO RODAPÉ
# =========================================================================

# Em pontos (1 pt ≈ 0,353 mm; 28,35 pt ≈ 10 mm)
MARGEM_RODAPE_PTS = 25
MARGEM_LATERAL_PTS = 28


# =========================================================================
# CORES DA UI (padrão V&P)
# =========================================================================

COR_FUNDO = "#0a0a0a"
COR_DESTAQUE = "#e8a817"
COR_TEXTO = "#ffffff"
COR_FUNDO_INPUT = "#1a1a1a"
COR_BORDA = "#333333"

# =========================================================================
# RECURSOS (ícones, imagens)
# =========================================================================
 
# Caminho do arquivo .ico, RELATIVO à pasta do main.py.
# Trocar o ícone: substituir o arquivo neste caminho.
ICONE_RELATIVO = "Assets/IP_icon.ico"
 
 
def resource_path(relativo):
    """
    Retorna o caminho absoluto de um recurso (assets, ícones, etc.),
    funcionando tanto em modo desenvolvimento quanto empacotado pelo
    PyInstaller (--onedir e --onefile).
 
    PyInstaller extrai os recursos em _MEIPASS em runtime; em dev,
    a base é a pasta deste arquivo.
    """
    if hasattr(sys, "_MEIPASS"):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).parent
    return str(base / relativo)
 
 
def caminho_icone():
    """
    Retorna o caminho absoluto do ícone, ou None se o arquivo não existir.
    Usar None permite que o app rode sem ícone caso o arquivo seja removido.
    """
    p = Path(resource_path(ICONE_RELATIVO))
    return str(p) if p.exists() else None
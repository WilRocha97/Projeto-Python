"""
=========================================================================
INSERIR TEXTO EM PDFs - V&P
=========================================================================
Entry point do aplicativo.

Para executar diretamente:
    python main.py

Para gerar o executável sem instalador (Não recomendado) (padrão V&P --onedir):
    pyinstaller --onedir --windowed --name "Inserir_Texto_PDF" main.py

Dependências:
    pip install PySide6 pypdf pdfplumber reportlab

Estrutura do projeto:
    main.py            — entry point
    config.py          — constantes (textos, cores, fontes, margens)
    text_utils.py      — extração de CNPJ e placeholders de data
    pdf_processor.py   — manipulação de PDF (overlay, merge, fonte)
    worker.py          — QThread que orquestra o lote
    styles.py          — folha de estilos QSS
    ui.py              — JanelaPrincipal (PySide6)
=========================================================================
"""

import sys, os, psutil, atexit
from pyautogui import confirm
from os import remove, getpid

def create_lock_file(lock_file_path):
    try:
        # Tente criar o arquivo de trava
        with open(lock_file_path, 'x') as lock_file:
            lock_file.write(str(getpid()))
        return True
    except FileExistsError:
        # O arquivo de trava já existe, indicando que outra instância está em execução
        return False


def remove_lock_file(lock_file_path):
    try:
        remove(lock_file_path)
    except FileNotFoundError:
        pass


 # Especifique o caminho para o arquivo de trava
lock_file_path = 'insere_info_pdf.lock'
while True:
    # Verifique se outra instância está em execução
    if not create_lock_file(lock_file_path):
        opcao = confirm(text="O programa já está em execução. Deseja encerrar o programa?", buttons=['Sim', 'Não'])
        if opcao == 'Não':
            sys.exit(1)
        else:
            os.remove(lock_file_path)
            """Encerra todos os processos com o nome especificado."""
            for proc in psutil.process_iter(['name']):
                """try:"""
                # print(proc.info['name'])
                if proc.info['name'].lower() == 'insere_info_pdf.exe'.lower():
                    proc.kill()
                    print(f"Processo 'insere_info_pdf' encerrado.")
                """except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass"""
            continue
    else:
        break

# Defina uma função para remover o arquivo de trava ao final da execução
atexit.register(remove_lock_file, lock_file_path)
    

versao_atual = '1.0.2'
caminho_arquivo = 'T:\ROBÔ\_Executáveis\Insere Info PDF\Insere Info PDF.exe'

cr = open('T:\\ROBÔ\\_Executáveis\\Insere Info PDF\\Versão.txt', 'r', encoding='utf-8').read()
if str(cr) != versao_atual:
    while True:
        atualizar = confirm(text='Existe uma nova versão do programa, deseja atualizar agora?', buttons=('Sim', 'Não'))
        if atualizar == 'Sim':
            break
        elif atualizar == 'Não':
            break
    
    if atualizar == 'Sim':
        os.startfile(caminho_arquivo)
        sys.exit()
        
# -------------------------------------------------------------------------
# DPI awareness (Windows)
# -------------------------------------------------------------------------
# O Qt 6 tenta configurar Per-Monitor V2 por padrão. Quando o Windows ou o
# manifesto injetado pelo PyInstaller já definiu outro contexto, o Qt loga:
#
#   qt.qpa.window: SetProcessDpiAwarenessContext() failed: Acesso negado.
#
# Setar QT_QPA_PLATFORM com dpiawareness=1 (System DPI Aware) ANTES de
# qualquer import do Qt evita o conflito. É compatível com a maioria das
# máquinas V&P (1366x768) e também funciona em monitores 4K.
# -------------------------------------------------------------------------
if sys.platform.startswith("win"):
    os.environ.setdefault("QT_QPA_PLATFORM", "windows:dpiawareness=1")
 
 
# -------------------------------------------------------------------------
# AppUserModelID (Windows)
# -------------------------------------------------------------------------
# Sem isso, o Windows usa a identidade do processo pai (python.exe ou
# bootloader do PyInstaller) para decidir o ícone da taskbar — o que faz
# o ícone do Python aparecer no lugar do nosso. Definir um AUMID próprio
# faz o Windows tratar o app como uma identidade separada e usar o ícone
# que setamos via setWindowIcon().
#
# Formato recomendado: Empresa.Produto.SubProduto.Versao
# -------------------------------------------------------------------------
if sys.platform.startswith("win"):
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "VeigaPostal.InserirTextoPDF.1.0"
        )
    except (AttributeError, OSError):
        pass
 
# Validação de dependências antes de carregar a UI.
try:
    import pdfplumber  # noqa: F401
    import pypdf       # noqa: F401
    import reportlab   # noqa: F401
except ImportError as e:
    print(f"Biblioteca faltando: {e}")
    print("Instale com: pip install PySide6 pypdf pdfplumber reportlab")
    sys.exit(1)

if sys.platform.startswith("win"):
    os.environ.setdefault("QT_QPA_PLATFORM", "windows:dpiawareness=1")
    
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from ui import JanelaPrincipal
from config import caminho_icone


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # Ícone do aplicativo (taskbar / atalho).
    icone = caminho_icone()
    if icone:
        app.setWindowIcon(QIcon(icone))
    
    janela = JanelaPrincipal()
    janela.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

"""
Worker thread para o processamento dos PDFs em background, mantendo
a UI responsiva.
"""

from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QThread, Signal

from text_utils import extrair_cnpj
from pdf_processor import processar_pdf


class WorkerProcessamento(QThread):
    """
    Thread que processa uma lista de PDFs, anexando ao texto base:
    - Texto extra opcional (parcela, etc.)
    - Sufixo de data ou data+hora (se modo_data != 'nenhum')
    - CNPJ extraído do nome do arquivo (se incluir_cnpj=True)

    Emite:
        progresso(atual, total, mensagem, sucesso) — após cada PDF
        concluido(sucessos, erros) — ao final do lote
    """

    progresso = Signal(int, int, str, bool)
    concluido = Signal(int, int)

    def __init__(self, pdfs, pasta_saida, texto, incluir_cnpj, modo_data,
                 texto_extra=""):
        super().__init__()
        self.pdfs = pdfs                  # lista de Path (PDFs a processar)
        self.pasta_saida = pasta_saida
        self.texto = texto
        self.incluir_cnpj = incluir_cnpj
        self.modo_data = modo_data        # 'nenhum' | 'data' | 'data_hora'
        self.texto_extra = texto_extra    # string a concatenar após o texto base
        self._cancelar = False

    def cancelar(self):
        self._cancelar = True

    def run(self):
        pdfs = self.pdfs
        total = len(pdfs)

        if total == 0:
            self.concluido.emit(0, 0)
            return

        Path(self.pasta_saida).mkdir(parents=True, exist_ok=True)

        # Sufixo do texto extra (parcela, etc.) — igual para todo o lote.
        sufixo_extra = f" - {self.texto_extra}" if self.texto_extra else ""

        # Sufixo de data calculado UMA vez para que todo o lote fique consistente.
        agora = datetime.now()
        sufixo_data = ""
        if self.modo_data == "data":
            sufixo_data = f" - {agora.strftime('%d/%m/%Y')}"
        elif self.modo_data == "data_hora":
            sufixo_data = f" - {agora.strftime('%d/%m/%Y %H:%M')}"

        sucessos, erros = 0, 0
        for i, pdf in enumerate(pdfs, 1):
            if self._cancelar:
                break

            # Texto base + extra + sufixo de data (igual para todos os PDFs do lote).
            texto_final = f"{self.texto}{sufixo_extra}{sufixo_data}"

            # Anexa o CNPJ se habilitado e encontrado no nome do arquivo.
            cnpj_info = ""
            if self.incluir_cnpj:
                cnpj = extrair_cnpj(pdf.name)
                if cnpj:
                    texto_final = f"{texto_final} - CNPJ: {cnpj}"
                    cnpj_info = f" | CNPJ: {cnpj}"
                else:
                    cnpj_info = " | CNPJ não encontrado no nome"

            saida = Path(self.pasta_saida) / pdf.name
            ok, msg = processar_pdf(str(pdf), str(saida), texto_final)
            if ok:
                sucessos += 1
                self.progresso.emit(i, total, f"OK   {pdf.name}  -  {msg}{cnpj_info}", True)
            else:
                erros += 1
                self.progresso.emit(i, total, f"ERRO {pdf.name}  -  {msg}", False)

        self.concluido.emit(sucessos, erros)
"""
Janela principal do "Inserir Texto em PDFs - V&P".

Toda a lógica de UI fica isolada aqui. As classes/funções que ela usa
vêm dos outros módulos:
- config:        constantes
- text_utils:    substituir_placeholders
- pdf_processor: (não direto — é usado pelo worker)
- worker:        WorkerProcessamento
- styles:        obter_estilo
"""

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit, QProgressBar,
    QFileDialog, QMessageBox, QFrame, QCheckBox, QStyledItemDelegate
)
from PySide6.QtGui import QIcon

from config import OPCOES_TEXTO, OPCOES_EXTRAS, PASTA_SAIDA_PADRAO, caminho_icone
from text_utils import substituir_placeholders
from styles import obter_estilo
from worker import WorkerProcessamento


class JanelaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.setWindowTitle("Inserir Texto em PDFs - V&P")
        self.setMinimumSize(720, 700)

        # Ícone da janela (barra de título).
        icone = caminho_icone()
        if icone:
            self.setWindowIcon(QIcon(icone))

        self._criar_ui()
        self.setStyleSheet(obter_estilo())
        self._atualizar_previa()

    # ----- UI -----
    def _criar_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        titulo = QLabel("Inserir Texto em PDFs")
        titulo.setObjectName("titulo")
        layout.addWidget(titulo)

        subt = QLabel("Adiciona um texto no rodapé da primeira página dos PDFs")
        subt.setObjectName("subtitulo")
        layout.addWidget(subt)

        layout.addWidget(self._linha())

        # Entrada: pasta com vários PDFs OU arquivo PDF único
        layout.addWidget(QLabel("PDF ou pasta com PDFs:"))
        h1 = QHBoxLayout()
        self.input_entrada = QLineEdit()
        self.input_entrada.setPlaceholderText("Selecione uma pasta ou um arquivo PDF...")
        btn_pasta = QPushButton("Pasta...")
        btn_pasta.clicked.connect(self._selecionar_pasta_entrada)
        btn_arquivo = QPushButton("PDF...")
        btn_arquivo.clicked.connect(self._selecionar_arquivo_entrada)
        h1.addWidget(self.input_entrada)
        h1.addWidget(btn_pasta)
        h1.addWidget(btn_arquivo)
        layout.addLayout(h1)

        # Pasta de saída
        layout.addWidget(
            QLabel("Pasta de saída (deixe vazio para criar uma subpasta automaticamente):")
        )
        h2 = QHBoxLayout()
        self.input_pasta_saida = QLineEdit()
        self.input_pasta_saida.setPlaceholderText(
            f"Padrão: <local_da_entrada>\\{PASTA_SAIDA_PADRAO}"
        )
        btn_saida = QPushButton("Procurar...")
        btn_saida.clicked.connect(self._selecionar_pasta_saida)
        h2.addWidget(self.input_pasta_saida)
        h2.addWidget(btn_saida)
        layout.addLayout(h2)

        # Combo de opções de texto
        layout.addWidget(QLabel("Texto a inserir:"))
        self.combo_texto = QComboBox()
        for rotulo in OPCOES_TEXTO.keys():
            self.combo_texto.addItem(rotulo)
        self.combo_texto.currentTextChanged.connect(self._atualizar_previa)
        self._forcar_estilo_no_popup(self.combo_texto)
        layout.addWidget(self.combo_texto)

        # Combo de texto extra opcional (parcela, etc.)
        layout.addWidget(QLabel("Texto referente a parcela (opcional):"))
        self.combo_extra = QComboBox()
        for rotulo in OPCOES_EXTRAS.keys():
            self.combo_extra.addItem(rotulo)
        self.combo_extra.currentTextChanged.connect(self._atualizar_previa)
        self._forcar_estilo_no_popup(self.combo_extra)
        layout.addWidget(self.combo_extra)

        # Opção: incluir CNPJ extraído do nome do arquivo
        self.check_cnpj = QCheckBox(
            "Adicionar CNPJ da empresa (extraído do nome do arquivo, se presente)"
        )
        self.check_cnpj.toggled.connect(self._atualizar_previa)
        layout.addWidget(self.check_cnpj)

        # Opções: incluir data ou data e hora atuais (mutuamente exclusivas)
        self.check_data = QCheckBox("Adicionar data atual")
        self.check_data.toggled.connect(self._on_check_data)
        layout.addWidget(self.check_data)

        self.check_data_hora = QCheckBox("Adicionar data e hora atuais")
        self.check_data_hora.toggled.connect(self._on_check_data_hora)
        layout.addWidget(self.check_data_hora)

        # Prévia
        layout.addWidget(QLabel("Prévia:"))
        self.label_previa = QLabel()
        self.label_previa.setObjectName("previa")
        self.label_previa.setWordWrap(True)
        layout.addWidget(self.label_previa)

        # Botão executar
        self.btn_executar = QPushButton("INSERIR TEXTO NOS PDFs")
        self.btn_executar.setObjectName("btnPrincipal")
        self.btn_executar.clicked.connect(self._executar)
        layout.addWidget(self.btn_executar)

        # Progresso
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # Log
        layout.addWidget(QLabel("Log:"))
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, 1)

    def _linha(self):
        f = QFrame()
        f.setFrameShape(QFrame.HLine)
        f.setObjectName("linha")
        return f

    def _forcar_estilo_no_popup(self, combo):
        """
        Trabalha em conjunto com 'combobox-popup: 0' no QSS.

        Trocar o delegate padrão do QComboBox por um QStyledItemDelegate
        força o Qt a renderizar cada item do popup usando o estilo QSS
        atual — inclusive as regras ::item:hover e ::item:selected.

        Sem isso, o Windows usa o delegate nativo do sistema que ignora
        cores customizadas e o hover fica "invisível".
        """
        combo.setItemDelegate(QStyledItemDelegate(combo))

    # ----- Eventos -----
    def _on_check_data(self, checked):
        # Data e Data+Hora são mutuamente exclusivas
        if checked:
            self.check_data_hora.setChecked(False)
        self._atualizar_previa()

    def _on_check_data_hora(self, checked):
        if checked:
            self.check_data.setChecked(False)
        self._atualizar_previa()

    def _atualizar_previa(self):
        texto = self._obter_texto_atual()
        if not texto:
            self.label_previa.setText("(nenhum texto selecionado)")
            return
        previa = substituir_placeholders(texto)
        extra = self._obter_texto_extra()
        if extra:
            previa = f"{previa} - {extra}"
        agora = datetime.now()
        if self.check_data.isChecked():
            previa = f"{previa} - {agora.strftime('%d/%m/%Y')}"
        elif self.check_data_hora.isChecked():
            previa = f"{previa} - {agora.strftime('%d/%m/%Y %H:%M')}"
        if self.check_cnpj.isChecked():
            previa = f"{previa} - CNPJ: <extraído do nome do arquivo>"
        self.label_previa.setText(f'"{previa}"')

    def _obter_texto_atual(self):
        return OPCOES_TEXTO.get(self.combo_texto.currentText(), "")

    def _obter_texto_extra(self):
        return OPCOES_EXTRAS.get(self.combo_extra.currentText(), "")

    def _selecionar_pasta_entrada(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecione a pasta com os PDFs")
        if pasta:
            self.input_entrada.setText(pasta)

    def _selecionar_arquivo_entrada(self):
        arquivo, _ = QFileDialog.getOpenFileName(
            self, "Selecione o PDF", "", "Arquivos PDF (*.pdf)"
        )
        if arquivo:
            self.input_entrada.setText(arquivo)

    def _selecionar_pasta_saida(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecione a pasta de saída")
        if pasta:
            self.input_pasta_saida.setText(pasta)

    def _executar(self):
        # Resolve a entrada — pode ser pasta ou arquivo PDF.
        entrada_str = self.input_entrada.text().strip()
        if not entrada_str:
            QMessageBox.warning(self, "Atenção",
                                 "Selecione uma pasta ou um arquivo PDF.")
            return

        entrada = Path(entrada_str)
        if not entrada.exists():
            QMessageBox.warning(self, "Atenção",
                                 "Caminho selecionado não existe.")
            return

        if entrada.is_file():
            if entrada.suffix.lower() != ".pdf":
                QMessageBox.warning(self, "Atenção",
                                     "O arquivo selecionado não é um PDF.")
                return
            pdfs = [entrada]
            pasta_base = entrada.parent
            modo_entrada = "arquivo"
        elif entrada.is_dir():
            pdfs = sorted(
                p for p in entrada.iterdir()
                if p.is_file() and p.suffix.lower() == ".pdf"
            )
            pasta_base = entrada
            modo_entrada = "pasta"
        else:
            QMessageBox.warning(self, "Atenção",
                                 "Selecione uma pasta ou um arquivo PDF válido.")
            return

        if not pdfs:
            QMessageBox.information(self, "Aviso",
                                     "Nenhum PDF encontrado na pasta selecionada.")
            return

        # Pasta de saída — padrão = subpasta no local da entrada.
        pasta_saida = self.input_pasta_saida.text().strip()
        if not pasta_saida:
            pasta_saida = str(pasta_base / PASTA_SAIDA_PADRAO)

        # Validação contra sobrescrita do(s) original(is).
        saida_resolved = Path(pasta_saida).resolve()
        if modo_entrada == "pasta" and saida_resolved == entrada.resolve():
            QMessageBox.warning(
                self, "Atenção",
                "A pasta de saída não pode ser igual à pasta de entrada.\n"
                "Os arquivos originais seriam sobrescritos."
            )
            return
        if modo_entrada == "arquivo" and saida_resolved == entrada.parent.resolve():
            QMessageBox.warning(
                self, "Atenção",
                "A pasta de saída não pode ser a mesma do PDF original.\n"
                "O arquivo original seria sobrescrito."
            )
            return

        texto = self._obter_texto_atual()
        if not texto:
            QMessageBox.warning(self, "Atenção", "Informe o texto a inserir.")
            return

        texto_final = substituir_placeholders(texto)
        texto_extra = self._obter_texto_extra()
        incluir_cnpj = self.check_cnpj.isChecked()

        # Modo de data (mutuamente exclusivo)
        if self.check_data.isChecked():
            modo_data = "data"
        elif self.check_data_hora.isChecked():
            modo_data = "data_hora"
        else:
            modo_data = "nenhum"

        # Monta string informativa dos extras para confirmação/log.
        extras = []
        if texto_extra:
            extras.append(texto_extra)
        if modo_data == "data":
            extras.append("data atual")
        elif modo_data == "data_hora":
            extras.append("data e hora atuais")
        if incluir_cnpj:
            extras.append("CNPJ do nome do arquivo")
        info_extras = " (+ " + ", ".join(extras) + ")" if extras else ""

        # Texto resumindo o que será processado.
        if modo_entrada == "arquivo":
            titulo_qtd = f"1 PDF selecionado:\n{entrada.name}"
        else:
            titulo_qtd = f"{len(pdfs)} PDF(s) encontrados na pasta"

        resp = QMessageBox.question(
            self,
            "Confirmar",
            f"{titulo_qtd}\n\n"
            f"Texto a inserir:\n\"{texto_final}\"{info_extras}\n\n"
            f"Pasta de saída:\n{pasta_saida}\n\n"
            f"Deseja continuar?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return

        self.log.clear()
        self.log.append(f"Iniciando processamento de {len(pdfs)} PDF(s)...")
        self.log.append(f'Texto: "{texto_final}"{info_extras}')
        self.log.append(f"Saída: {pasta_saida}")
        self.log.append("")

        self.progress.setVisible(True)
        self.progress.setMaximum(len(pdfs))
        self.progress.setValue(0)
        self.btn_executar.setEnabled(False)

        self.worker = WorkerProcessamento(
            pdfs, pasta_saida, texto_final, incluir_cnpj, modo_data, texto_extra
        )
        self.worker.progresso.connect(self._on_progresso)
        self.worker.concluido.connect(self._on_concluido)
        self.worker.start()

    def _on_progresso(self, atual, total, msg, sucesso):
        self.progress.setValue(atual)
        cor = "#7CFC00" if sucesso else "#FF6B6B"
        self.log.append(f'<span style="color:{cor}">{msg}</span>')

    def _on_concluido(self, sucessos, erros):
        self.btn_executar.setEnabled(True)
        self.log.append("")
        self.log.append(f"Concluído! Sucessos: {sucessos} | Erros: {erros}")
        QMessageBox.information(
            self, "Concluído",
            f"Processamento finalizado.\n\nSucessos: {sucessos}\nErros: {erros}"
        )
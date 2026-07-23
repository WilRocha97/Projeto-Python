"""
Folha de estilos QSS (tema escuro V&P).

Centralizar o estilo aqui facilita futuras alterações visuais sem
mexer na lógica da janela.
"""

from config import (
    COR_FUNDO,
    COR_DESTAQUE,
    COR_TEXTO,
    COR_FUNDO_INPUT,
    COR_BORDA,
)


def obter_estilo():
    """Retorna a folha de estilos QSS completa, com cores do config.py."""
    return f"""
        QMainWindow, QWidget {{
            background-color: {COR_FUNDO};
            color: {COR_TEXTO};
        }}
        QLabel {{
            color: {COR_TEXTO};
            font-size: 10pt;
        }}
        QLabel#titulo {{
            color: {COR_DESTAQUE};
            font-size: 18pt;
            font-weight: bold;
        }}
        QLabel#subtitulo {{
            color: #999999;
            font-size: 10pt;
        }}
        QLabel#previa {{
            background-color: {COR_FUNDO_INPUT};
            border: 1px solid {COR_BORDA};
            border-radius: 4px;
            padding: 8px;
            color: {COR_DESTAQUE};
            font-style: italic;
        }}
        QLineEdit, QComboBox, QTextEdit {{
            background-color: {COR_FUNDO_INPUT};
            color: {COR_TEXTO};
            border: 1px solid {COR_BORDA};
            border-radius: 4px;
            padding: 6px;
            font-size: 10pt;
        }}
        QLineEdit:focus, QComboBox:focus {{
            border: 1px solid {COR_DESTAQUE};
        }}
        /*
         * combobox-popup: 0 força o Qt a usar o popup interno estilizável
         * em vez do popup NATIVO do Windows (que ignora QSS). Sem essa
         * linha, as regras ::item:hover NÃO funcionam no Windows.
         */
        QComboBox {{
            combobox-popup: 0;
        }}
        QComboBox::drop-down {{
            border: none;
            width: 22px;
        }}
        /* Lista suspensa do combo (o popup) */
        QComboBox QAbstractItemView {{
            background-color: {COR_FUNDO_INPUT};
            color: {COR_TEXTO};
            border: 1px solid {COR_BORDA};
            /* outline:0 remove o retângulo pontilhado de foco em cima do item */
            outline: 0;
        }}
        /* Cada item da lista (aparência base) */
        QComboBox QAbstractItemView::item {{
            background-color: {COR_FUNDO_INPUT};
            color: {COR_TEXTO};
            padding: 6px 10px;
            min-height: 22px;
            border: none;
        }}
        /* Hover: quando o mouse passa por cima de um item */
        QComboBox QAbstractItemView::item:hover {{
            background-color: {COR_DESTAQUE};
            color: {COR_FUNDO};
        }}
        /* Selected: item destacado pelas setas do teclado ou item ativo atual */
        QComboBox QAbstractItemView::item:selected {{
            background-color: {COR_DESTAQUE};
            color: {COR_FUNDO};
        }}
        QPushButton {{
            background-color: {COR_FUNDO_INPUT};
            color: {COR_TEXTO};
            border: 1px solid {COR_BORDA};
            border-radius: 4px;
            padding: 8px 16px;
            font-size: 10pt;
        }}
        QPushButton:hover {{
            border: 1px solid {COR_DESTAQUE};
            color: {COR_DESTAQUE};
        }}
        QPushButton#btnPrincipal {{
            background-color: {COR_DESTAQUE};
            color: {COR_FUNDO};
            font-weight: bold;
            font-size: 11pt;
            padding: 12px;
        }}
        QPushButton#btnPrincipal:hover {{
            background-color: #ffc943;
        }}
        QPushButton:disabled {{
            background-color: #2a2a2a;
            color: #666666;
            border: 1px solid #2a2a2a;
        }}
        QProgressBar {{
            background-color: {COR_FUNDO_INPUT};
            border: 1px solid {COR_BORDA};
            border-radius: 4px;
            text-align: center;
            color: {COR_TEXTO};
        }}
        QProgressBar::chunk {{
            background-color: {COR_DESTAQUE};
            border-radius: 3px;
        }}
        QFrame#linha {{
            color: {COR_BORDA};
            background-color: {COR_BORDA};
            max-height: 1px;
        }}
        QCheckBox {{
            color: {COR_TEXTO};
            font-size: 10pt;
            padding: 4px 0;
            spacing: 8px;
        }}
        QCheckBox::indicator {{
            width: 16px;
            height: 16px;
            border: 1px solid {COR_BORDA};
            background-color: {COR_FUNDO_INPUT};
            border-radius: 3px;
        }}
        QCheckBox::indicator:hover {{
            border: 1px solid {COR_DESTAQUE};
        }}
        QCheckBox::indicator:checked {{
            background-color: {COR_DESTAQUE};
            border: 1px solid {COR_DESTAQUE};
        }}
    """
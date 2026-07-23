"""
Funções de processamento e manipulação de PDFs.

Este módulo é completamente independente da UI — pode ser reutilizado
em scripts de linha de comando ou em outros robôs.
"""

from collections import Counter
from io import BytesIO

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

from config import (
    TAMANHO_FONTE_OVERRIDE,
    TAMANHO_FONTE_PADRAO,
    MARGEM_RODAPE_PTS,
    MARGEM_LATERAL_PTS,
)


def detectar_tamanho_fonte(pdf_path):
    """
    Detecta o tamanho de fonte mais usado na primeira página do PDF.
    Retorna o tamanho em pontos (int). Em caso de falha, retorna o padrão.
    """
    if TAMANHO_FONTE_OVERRIDE is not None:
        return TAMANHO_FONTE_OVERRIDE
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return TAMANHO_FONTE_PADRAO
            chars = pdf.pages[0].chars or []
            tamanhos = [round(c["size"]) for c in chars if c.get("size")]
            if not tamanhos:
                return TAMANHO_FONTE_PADRAO
            return Counter(tamanhos).most_common(1)[0][0]
    except Exception:
        return TAMANHO_FONTE_PADRAO


def criar_overlay(texto, page_width, page_height, fonte_tamanho, rotacao):
    """
    Cria uma página PDF com o texto posicionado no RODAPÉ VISUAL,
    considerando a rotação da página original.

    O canvas usa o tamanho NATURAL da página (mediabox). Quando o viewer
    aplica /Rotate, o texto fica posicionado e legível corretamente.
    """
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(page_width, page_height))
    c.setFont("Helvetica", fonte_tamanho)

    mx = (MARGEM_LATERAL_PTS - MARGEM_LATERAL_PTS) + 2
    my = (MARGEM_RODAPE_PTS - MARGEM_RODAPE_PTS) + 2

    print(mx, my)
    if rotacao == 90:
        # Viewer rotaciona 90° clockwise → rodapé visual = lado esquerdo natural.
        # Texto rotacionado +90° para ficar legível após a rotação do viewer.
        c.saveState()
        c.translate(my, mx)
        c.rotate(90)
        c.drawString(0, 0, texto)
        c.restoreState()
    elif rotacao == 180:
        # Viewer rotaciona 180° → rodapé visual = topo natural.
        c.saveState()
        c.translate(page_width - mx, page_height - my)
        c.rotate(180)
        c.drawString(0, 0, texto)
        c.restoreState()
    elif rotacao == 270:
        # Viewer rotaciona 270° cw → rodapé visual = lado direito natural.
        c.saveState()
        c.translate(page_width - my, page_height - mx)
        c.rotate(-90)
        c.drawString(0, 0, texto)
        c.restoreState()
    else:
        # Sem rotação: rodapé natural (canto inferior-esquerdo).
        c.drawString(mx, my, texto)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


def processar_pdf(pdf_entrada, pdf_saida, texto):
    """
    Insere o texto no rodapé da primeira página do PDF.
    Retorna (sucesso: bool, mensagem: str).
    """
    try:
        reader = PdfReader(pdf_entrada)

        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                return False, "PDF protegido por senha"

        if not reader.pages:
            return False, "PDF sem páginas"

        fonte_tamanho = detectar_tamanho_fonte(pdf_entrada)

        primeira = reader.pages[0]

        # Detecta a rotação da página
        try:
            rotacao = int(primeira.rotation) % 360
        except Exception:
            rot_obj = primeira.get("/Rotate", 0)
            rotacao = int(rot_obj or 0) % 360

        page_width = float(primeira.mediabox.width)
        page_height = float(primeira.mediabox.height)

        overlay_buffer = criar_overlay(
            texto, page_width, page_height, fonte_tamanho, rotacao
        )
        overlay_reader = PdfReader(overlay_buffer)
        overlay_page = overlay_reader.pages[0]

        primeira.merge_page(overlay_page)

        writer = PdfWriter()
        writer.add_page(primeira)
        for pagina in reader.pages[1:]:
            writer.add_page(pagina)

        with open(pdf_saida, "wb") as f:
            writer.write(f)

        return True, f"OK (fonte: {fonte_tamanho}pt | rotação: {rotacao}°)"
    except Exception as e:
        return False, f"Erro: {e}"

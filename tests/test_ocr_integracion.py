"""Integración LENTA: el OCR real (Tesseract portable + capa invisible) lee un PDF escaneado y el resaltador encuentra
el nombre. Se salta sola si no está el Tesseract empaquetado (ciclo rápido / CI sin el portable).

    python -m pytest tests/test_ocr_integracion.py -m integracion -q
"""
import io
import os
from pathlib import Path

import pymupdf as fitz
import pytest

RAIZ = Path(__file__).resolve().parent.parent
TESSERACT = RAIZ / "tesseract" / "tesseract.exe"
TESSDATA = RAIZ / "tesseract" / "tessdata"


def _pdf_escaneado(tmp_path, texto: str) -> str:
    """Un PDF 'escaneado': se escribe el texto, se rasteriza la página y se arma un PDF de solo imagen (sin capa de
    texto), como una planilla escaneada real."""
    con_texto = fitz.open()
    pagina = con_texto.new_page(width=595, height=842)
    pagina.insert_text((60, 120), texto, fontsize=28)
    pix = pagina.get_pixmap(dpi=300)
    con_texto.close()

    solo_imagen = fitz.open()
    p = solo_imagen.new_page(width=pix.width, height=pix.height)
    p.insert_image(p.rect, pixmap=pix)
    ruta = tmp_path / "escaneado.pdf"
    solo_imagen.save(str(ruta))
    solo_imagen.close()
    return str(ruta)


@pytest.mark.integracion
def test_ocr_resalta_nombre_en_pdf_escaneado(tmp_path):
    if not TESSERACT.is_file() or not (TESSDATA / "spa.traineddata").is_file():
        pytest.skip(f"falta el Tesseract empaquetado ({TESSERACT}); no se puede probar el OCR real")
    os.environ["TESSERACT_EXE"] = str(TESSERACT)
    os.environ["TESSDATA_PREFIX"] = str(TESSDATA)
    from resaltado_pdf import resaltar_nombres_en_pdf

    # Confirmamos que de verdad es un escaneo (sin texto), si no, no probaría el OCR.
    doc = fitz.open(_pdf_escaneado(tmp_path, "JUAN PEREZ MORA"))
    assert doc[0].get_text().strip() == ""
    doc.close()

    ruta = _pdf_escaneado(tmp_path, "JUAN PEREZ MORA")
    res = resaltar_nombres_en_pdf(ruta, ["JUAN PEREZ MORA"], str(tmp_path / "resaltado.pdf"))

    assert res.uso_ocr is True                                   # se leyó con OCR
    assert res.coincidencias_por_nombre["JUAN PEREZ MORA"] >= 1  # y encontró el nombre

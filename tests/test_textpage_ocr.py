"""La capa OCR: escribir las palabras de Tesseract como texto invisible deja la página buscable. Sin Tesseract, no se
toca nada. (El OCR real de punta a punta lo cubre test_ocr_integracion.py.)"""
import pymupdf as fitz

from resaltado_pdf import _asegurar_buscable, _escribir_capa_invisible


def test_capa_invisible_deja_la_pagina_buscable():
    doc = fitz.open()
    p = doc.new_page()
    # (texto, left, top, ancho, alto) en píxeles, como los da Tesseract sobre un render a 300 dpi
    palabras = [("ARGUELLO", 100, 300, 180, 40), ("JIMENEZ", 300, 300, 160, 40)]
    _escribir_capa_invisible(p, palabras, 72 / 300)

    tp = p.get_textpage()
    assert p.search_for("ARGUELLO", textpage=tp)
    assert p.search_for("JIMENEZ", textpage=tp)
    texto = p.get_text(textpage=tp)
    assert "ARGUELLO" in texto and "JIMENEZ" in texto


def test_dedup_no_repite_la_misma_palabra():
    doc = fitz.open()
    p = doc.new_page()
    palabras = [("HOLA", 100, 100, 80, 30), ("HOLA", 101, 101, 80, 30)]  # misma palabra, casi misma posición
    _escribir_capa_invisible(p, palabras, 72 / 300)
    assert p.get_text(textpage=p.get_textpage()).upper().count("HOLA") == 1


def test_sin_tesseract_devuelve_el_original(tmp_path, monkeypatch):
    monkeypatch.delenv("TESSERACT_EXE", raising=False)
    monkeypatch.setattr("resaltado_pdf.shutil.which", lambda _nombre: None)   # tampoco hay uno en el sistema
    doc = fitz.open()
    doc.new_page()
    ruta = tmp_path / "x.pdf"
    doc.save(str(ruta))
    doc.close()

    salida, uso_ocr = _asegurar_buscable(str(ruta), tmp_path / "buscables")
    assert salida == str(ruta)
    assert uso_ocr is False

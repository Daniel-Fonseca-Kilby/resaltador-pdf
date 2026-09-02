"""Prueba directa de _buscar_por_fila (agrupar palabras de una misma fila)."""
import pymupdf as fitz

from resaltado_pdf import _buscar_por_fila


def test_agrupa_palabras_de_la_misma_fila(ruta_pdf_ejemplo):
    documento = fitz.open(ruta_pdf_ejemplo)
    pagina = documento[0]

    filas = _buscar_por_fila(pagina, ["JUAN", "PEREZ"])

    assert len(filas) == 1
    assert len(filas[0]) == 2  # un rect por cada palabra encontrada
    documento.close()


def test_no_agrupa_si_una_palabra_no_existe(ruta_pdf_ejemplo):
    documento = fitz.open(ruta_pdf_ejemplo)
    pagina = documento[0]

    filas = _buscar_por_fila(pagina, ["JUAN", "PALABRA_INEXISTENTE"])

    assert filas == []
    documento.close()

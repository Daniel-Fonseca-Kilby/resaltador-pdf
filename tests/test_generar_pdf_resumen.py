"""Pruebas de generar_pdf_resumen (el PDF que va dentro del zip del modo cliente)."""
import pymupdf as fitz

from resaltado_pdf import generar_pdf_resumen


def test_genera_pdf_valido_con_texto():
    datos = generar_pdf_resumen(
        "Resumen del procesamiento",
        [
            ("Archivos que no se pudieron procesar", ["archivo1.pdf: error de prueba"]),
            ("Cédulas no encontradas", ["Juan Perez (cedula 111111111, cliente Cliente Uno)"]),
        ],
    )
    assert isinstance(datos, bytes)
    assert datos.startswith(b"%PDF")

    documento = fitz.open(stream=datos, filetype="pdf")
    texto = documento[0].get_text()
    assert "Resumen del procesamiento" in texto
    assert "archivo1.pdf" in texto
    assert "111111111" in texto
    documento.close()


def test_ignora_secciones_vacias():
    datos = generar_pdf_resumen("Titulo", [("Seccion vacia", [])])
    documento = fitz.open(stream=datos, filetype="pdf")
    texto = documento[0].get_text()
    assert "Seccion vacia" not in texto
    documento.close()


def test_pagina_nueva_si_hay_muchas_lineas():
    lineas = [f"Linea numero {i}" for i in range(80)]
    datos = generar_pdf_resumen("Titulo", [("Seccion larga", lineas)])
    documento = fitz.open(stream=datos, filetype="pdf")
    assert documento.page_count > 1
    documento.close()

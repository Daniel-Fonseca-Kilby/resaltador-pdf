"""
conftest.py

Fixtures de resaltado_pdf.py. El PDF de ejemplo se genera en memoria con
nombres y cédulas inventados, no hace falta ningún dato real.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pymupdf as fitz
import pytest


def _escribir_fila(pagina, y, celdas, fontsize=10):
    """Escribe una fila de texto en columnas separadas horizontalmente,
    imitando cómo una planilla real reparte el texto en columnas anchas."""
    x = 36
    for texto, ancho in celdas:
        pagina.insert_text((x, y), texto, fontsize=fontsize)
        x += ancho


# separación amplia (40pt) para no depender de los márgenes exactos de la fuente
Y_TITULOS = 100
Y_PRIMERA_FILA = 140
ALTO_FILA = 30

FILAS_EJEMPLO = [
    ("111111111", "JUAN", "PEREZ MORA", "Ninguna"),
    ("222222222", "MARIA JOSE", "SOLANO MU OZ", "Ninguna"),
    ("333333333", "LUIS", "ZUNIGA RAMIREZ", "Ninguna"),
]


@pytest.fixture
def ruta_pdf_ejemplo(tmp_path):
    """Tabla inventada estilo 'mnk': encabezado + títulos + 3 filas.
    La fila 2 trae "MU OZ" (espacio en vez de Ñ) para imitar el bug de MNK."""
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)

    pagina.insert_text((36, 40), "EMPRESA DE PRUEBA S.A.", fontsize=13)
    pagina.insert_text((36, 60), "Planilla de prueba - Junio", fontsize=10)

    _escribir_fila(pagina, Y_TITULOS, [
        ("IDENTIFICACION", 90), ("NOMBRE", 80), ("APELLIDOS", 100), ("OBSERVACION", 90),
    ])

    y = Y_PRIMERA_FILA
    for cedula, nombre, apellidos, obs in FILAS_EJEMPLO:
        _escribir_fila(pagina, y, [
            (cedula, 90), (nombre, 80), (apellidos, 100), (obs, 90),
        ])
        y += ALTO_FILA

    ruta = tmp_path / "planilla_prueba.pdf"
    documento.save(str(ruta))
    documento.close()
    return str(ruta)


@pytest.fixture
def registros_ejemplo():
    """Registros ficticios equivalentes a lo que vendría de un Excel con
    columnas Identificación y Cliente."""
    return [
        {"cedula": "111111111", "cliente": "Cliente Prueba Uno", "nombre": "Juan Perez Mora"},
        {"cedula": "222222222", "cliente": "Cliente Prueba Uno", "nombre": "Maria Jose Solano Munoz"},
        {"cedula": "333333333", "cliente": "Cliente Prueba Dos", "nombre": "Luis Zuniga Ramirez"},
        {"cedula": "999999999", "cliente": "Cliente Prueba Dos", "nombre": "No Existe En El Pdf"},
    ]

"""Pruebas de integración de las rutas HTTP de api/index.py (Flask
test_client, sin arrancar un servidor real). Usa los mismos PDF/registros
inventados de conftest.py -nada de datos reales de empleados."""
import base64
import io
import json

import openpyxl
import pytest

from api.index import app


@pytest.fixture
def cliente_flask():
    return app.test_client()


def _abrir_pdf(ruta):
    with open(ruta, "rb") as f:
        return io.BytesIO(f.read())


def _crear_excel_cliente(registros):
    """Excel con columnas Identificación/Cliente/Nombre, como lo subiría
    un usuario para el Modo Cliente."""
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(["Identificacion", "Cliente", "Nombre"])
    for registro in registros:
        hoja.append([registro["cedula"], registro["cliente"], registro["nombre"]])
    buffer = io.BytesIO()
    libro.save(buffer)
    buffer.seek(0)
    return buffer


def test_index_devuelve_200(cliente_flask):
    respuesta = cliente_flask.get("/")
    assert respuesta.status_code == 200


def test_procesar_sin_pdfs_devuelve_400(cliente_flask):
    respuesta = cliente_flask.post("/api/procesar", data={"nombres": "Juan Perez"})

    assert respuesta.status_code == 400
    assert respuesta.is_json
    assert "PDF" in respuesta.get_json()["error"]


def test_procesar_excel_no_valido_devuelve_400(cliente_flask, ruta_pdf_ejemplo):
    datos = {
        "nombres": "",
        "excel": (io.BytesIO(b"esto no es un excel"), "lista.txt"),
        "pdfs": (_abrir_pdf(ruta_pdf_ejemplo), "planilla.pdf"),
    }

    respuesta = cliente_flask.post("/api/procesar", data=datos, content_type="multipart/form-data")

    assert respuesta.status_code == 400
    assert "Excel" in respuesta.get_json()["error"]


def test_procesar_modo_simple_devuelve_zip(cliente_flask, ruta_pdf_ejemplo):
    datos = {
        "nombres": "Juan Perez",
        "pdfs": (_abrir_pdf(ruta_pdf_ejemplo), "planilla.pdf"),
    }

    respuesta = cliente_flask.post("/api/procesar", data=datos, content_type="multipart/form-data")

    assert respuesta.status_code == 200
    assert respuesta.headers["Content-Type"] == "application/zip"
    assert respuesta.headers["X-Modo"] == "simple"
    assert respuesta.headers["X-Total-Archivos"] == "1"


def test_procesar_modo_cliente_devuelve_zip(cliente_flask, ruta_pdf_ejemplo, registros_ejemplo):
    datos = {
        "excel": (_crear_excel_cliente(registros_ejemplo), "planilla.xlsx"),
        "pdfs": (_abrir_pdf(ruta_pdf_ejemplo), "planilla.pdf"),
        "formato": "mnk",
    }

    respuesta = cliente_flask.post("/api/procesar", data=datos, content_type="multipart/form-data")

    assert respuesta.status_code == 200
    assert respuesta.headers["Content-Type"] == "application/zip"
    assert respuesta.headers["X-Modo"] == "cliente"
    assert respuesta.headers["X-Total-Clientes"] == "2"


def test_procesar_modo_cliente_manda_no_encontrados_en_cabecera(cliente_flask, ruta_pdf_ejemplo, registros_ejemplo):
    # registros_ejemplo trae una cédula (999999999) que no está en el PDF
    datos = {
        "excel": (_crear_excel_cliente(registros_ejemplo), "planilla.xlsx"),
        "pdfs": (_abrir_pdf(ruta_pdf_ejemplo), "planilla.pdf"),
        "formato": "mnk",
    }

    respuesta = cliente_flask.post("/api/procesar", data=datos, content_type="multipart/form-data")

    assert respuesta.headers["X-Total-No-Encontrados"] == "1"
    lista = json.loads(base64.b64decode(respuesta.headers["X-No-Encontrados-B64"]).decode("utf-8"))
    assert len(lista) == 1
    assert "999999999" in lista[0]


def test_detectar_modo_excel_reconoce_columnas_de_cliente(cliente_flask, registros_ejemplo):
    datos = {"excel": (_crear_excel_cliente(registros_ejemplo), "planilla.xlsx")}

    respuesta = cliente_flask.post("/api/detectar-modo-excel", data=datos, content_type="multipart/form-data")

    assert respuesta.status_code == 200
    cuerpo = respuesta.get_json()
    assert cuerpo["modo"] == "cliente"
    assert cuerpo["total_registros"] == 4


def test_detectar_modo_excel_una_sola_columna_es_modo_simple(cliente_flask):
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(["Nombre"])
    hoja.append(["Juan Perez"])
    buffer = io.BytesIO()
    libro.save(buffer)
    buffer.seek(0)

    datos = {"excel": (buffer, "nombres.xlsx")}
    respuesta = cliente_flask.post("/api/detectar-modo-excel", data=datos, content_type="multipart/form-data")

    assert respuesta.status_code == 200
    assert respuesta.get_json()["modo"] == "simple"


def test_detectar_modo_excel_columnas_no_reconocidas_da_error(cliente_flask):
    libro = openpyxl.Workbook()
    hoja = libro.active
    hoja.append(["Fecha", "Comentario"])
    hoja.append(["2026-01-01", "Nota"])
    buffer = io.BytesIO()
    libro.save(buffer)
    buffer.seek(0)

    datos = {"excel": (buffer, "otro.xlsx")}
    respuesta = cliente_flask.post("/api/detectar-modo-excel", data=datos, content_type="multipart/form-data")

    assert respuesta.status_code == 400
    assert "no contiene una columna" in respuesta.get_json()["error"]

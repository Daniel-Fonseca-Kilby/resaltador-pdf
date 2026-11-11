"""
test_seguridad.py

Encabezados de seguridad, CSP y caché de Planillas (los comunes de las apps de Grupo VMA, api/seguridad_web.py), y
JSZip servido desde la propia app en vez de un CDN de internet.
"""
import pytest

from api import seguridad_web
from api.index import _historial_solicitudes_por_ip, app


@pytest.fixture
def cliente(monkeypatch):
    _historial_solicitudes_por_ip.clear()
    monkeypatch.setitem(app.config, "SESION_VMA_SECRETO", None)  # abierta: aquí solo importan los encabezados
    return app.test_client()


def test_la_pagina_lleva_la_csp_con_sus_scripts(cliente):
    respuesta = cliente.get("/")
    html, politica = respuesta.get_data(as_text=True), respuesta.headers["Content-Security-Policy"]
    assert respuesta.status_code == 200 and "<script>" in html
    assert seguridad_web.scripts_cubiertos(html, politica)
    assert "script-src 'self' 'sha256-" in politica and "frame-ancestors 'none'" in politica
    assert respuesta.headers["Cache-Control"] == "no-store"
    for nombre, valor in seguridad_web.ENCABEZADOS:
        assert respuesta.headers[nombre] == valor


def test_jszip_se_sirve_desde_aqui(cliente):
    html = cliente.get("/").get_data(as_text=True)
    assert "cdnjs" not in html and "https://" not in html.split("<footer")[0].split("</head>")[0]
    assert '<script src="/static/vendor/jszip.min.js"></script>' in html
    archivo = cliente.get("/static/vendor/jszip.min.js")
    assert archivo.status_code == 200 and b"JSZip v3.10.1" in archivo.data[:200]
    archivo.close()


def test_sin_manejadores_en_atributos(cliente):
    """La CSP bloquea onclick/onchange escritos en el HTML: todo va por addEventListener."""
    import re
    html = cliente.get("/").get_data(as_text=True)
    assert not re.search(r"\son[a-z]+\s*=", html)


def test_la_api_va_sin_cache(cliente):
    respuesta = cliente.post("/api/procesar")
    assert respuesta.headers["Cache-Control"] == "no-store"
    assert "Content-Security-Policy" in respuesta.headers


# --- Archivos subidos ---------------------------------------------------------------------------------------------
@pytest.mark.parametrize("subido, esperado", [
    ("planilla enero (2).pdf", "planilla enero (2).pdf"),
    ("Planilla Ñandú.pdf", "Planilla Ñandú.pdf"),
    ("../../secreto.pdf", "secreto.pdf"),
    ("..\\..\\windows\\secreto.pdf", "secreto.pdf"),
    ("C:\\temp\\x.pdf", "x.pdf"),
    ("x.pdf:oculto.pdf", "x.pdfoculto.pdf"),     # flujo alternativo de NTFS
    ("a\r\nb<c>|d?.pdf", "abcd.pdf"),
    ("NUL.pdf", "_NUL.pdf"),                     # nombres reservados de Windows
    ("con", "_con"),
    ("puntos.pdf...  ", "puntos.pdf"),
    ("", "archivo"),
    ("...", "archivo"),
])
def test_nombre_seguro(subido, esperado):
    from api.index import _nombre_seguro
    assert _nombre_seguro(subido) == esperado


def test_nombre_seguro_no_es_eterno():
    from api.index import _nombre_seguro
    nombre = _nombre_seguro("x" * 500 + ".pdf")
    assert len(nombre) <= 150 and nombre.endswith(".pdf")


def test_un_archivo_que_no_es_pdf_aunque_se_llame_asi():
    import io
    from api.index import _es_pdf
    from werkzeug.datastructures import FileStorage
    falso = FileStorage(io.BytesIO(b"MZ esto es un programa"), "planilla.pdf")
    verdadero = FileStorage(io.BytesIO(b"%PDF-1.7 contenido"), "planilla.PDF")
    assert not _es_pdf(falso) and _es_pdf(verdadero)
    assert verdadero.stream.read(4) == b"%PDF"  # lo deja listo para guardarlo completo
    assert not _es_pdf(FileStorage(io.BytesIO(b"%PDF-1.7"), "planilla.exe"))


def _excel_con(partes: dict) -> bytes:
    import io
    import zipfile
    salida = io.BytesIO()
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as archivo:
        for nombre, contenido in partes.items():
            archivo.writestr(nombre, contenido)
    return salida.getvalue()


def _subir_excel(cliente, contenido: bytes):
    import io
    return cliente.post("/api/detectar-modo-excel", content_type="multipart/form-data",
                        data={"excel": (io.BytesIO(contenido), "planilla.xlsx")})


def test_excel_que_descomprimido_es_enorme_se_rechaza(cliente, monkeypatch):
    monkeypatch.setattr("api.index._MAXIMO_EXCEL_DESCOMPRIMIDO", 1000)
    respuesta = _subir_excel(cliente, _excel_con({"xl/worksheets/sheet1.xml": b"<a>" + b" " * 5000 + b"</a>"}))
    assert respuesta.status_code == 400 and "demasiado grande" in respuesta.get_json()["error"]


def test_excel_con_entidades_xml_se_rechaza(cliente):
    import io
    import zipfile
    import openpyxl
    salida = io.BytesIO()
    openpyxl.Workbook().save(salida)
    with zipfile.ZipFile(salida) as bueno:  # un Excel de verdad, con un DOCTYPE metido en su hoja
        partes = {nombre: bueno.read(nombre) for nombre in bueno.namelist()}
    hoja = "xl/worksheets/sheet1.xml"
    partes[hoja] = partes[hoja].replace(b"<worksheet", b'<!DOCTYPE a [<!ENTITY x "y">]><worksheet', 1)
    respuesta = _subir_excel(cliente, _excel_con(partes))
    assert respuesta.status_code == 400 and "no se pudo leer" in respuesta.get_json()["error"]


def test_excel_de_verdad_se_sigue_leyendo(cliente):
    import io
    import openpyxl
    libro = openpyxl.Workbook()
    libro.active.append(["Nombre"])
    libro.active.append(["Juan Perez"])
    salida = io.BytesIO()
    libro.save(salida)
    assert _subir_excel(cliente, salida.getvalue()).status_code == 200


def test_el_resumen_excel_no_ejecuta_formulas_de_los_datos():
    """Un nombre o cliente que empiece con = no puede volverse una fórmula al abrir el resumen en Excel."""
    import io
    import openpyxl
    from resaltado_pdf import generar_excel_resumen
    malicioso = '=HYPERLINK("http://otro-sitio.com","clic")'
    datos = generar_excel_resumen([{"cedula": "=1+1", "numero_asegurado": "", "nombre": malicioso, "cliente": malicioso,
                                    "polizas": ["=2+2"], "encontrado": True}])
    libro = openpyxl.load_workbook(io.BytesIO(datos))
    for hoja in libro:
        for fila in hoja.iter_rows():
            for celda in fila:
                assert celda.data_type != "f", (hoja.title, celda.coordinate, celda.value)
    assert libro["Detalle de Oficiales"]["C2"].value == malicioso  # el texto queda igual, como texto


def test_el_depurador_de_flask_no_se_puede_encender_desde_el_codigo():
    """debug=True abre una consola que ejecuta código: quien la alcanza, controla el servidor."""
    from pathlib import Path
    codigo = (Path(__file__).resolve().parent.parent / "api" / "index.py").read_text(encoding="utf-8")
    assert "debug=True" not in codigo


def test_el_doctype_no_se_esconde_detras_de_un_comentario_largo(cliente):
    import io
    import zipfile
    import openpyxl
    salida = io.BytesIO()
    openpyxl.Workbook().save(salida)
    with zipfile.ZipFile(salida) as bueno:
        partes = {nombre: bueno.read(nombre) for nombre in bueno.namelist()}
    hoja = "xl/worksheets/sheet1.xml"
    relleno = b"<!--" + b"x" * 5000 + b"-->"
    partes[hoja] = partes[hoja].replace(b"<worksheet", relleno + b'<!DOCTYPE a [<!ENTITY x "y">]><worksheet', 1)
    respuesta = _subir_excel(cliente, _excel_con(partes))
    assert respuesta.status_code == 400 and "no se pudo leer" in respuesta.get_json()["error"]


def test_la_csp_de_planillas_no_admite_estilos_inyectados(cliente):
    respuesta = cliente.get("/")
    politica = respuesta.headers["Content-Security-Policy"]
    assert "unsafe-inline" not in politica
    assert seguridad_web.estilos_cubiertos(respuesta.get_data(as_text=True), politica)


# --- PDF que vienen de afuera -------------------------------------------------------------------------------------
def test_un_pdf_con_demasiadas_paginas_no_se_procesa(ruta_pdf_ejemplo, tmp_path, monkeypatch):
    import resaltado_pdf
    monkeypatch.setattr(resaltado_pdf, "MAXIMO_PAGINAS_POR_PDF", 0)
    with pytest.raises(ValueError, match="demasiadas páginas"):
        resaltado_pdf.resaltar_nombres_en_pdf(ruta_pdf_ejemplo, ["JUAN"], str(tmp_path / "salida.pdf"))


def test_en_un_lote_el_pdf_enorme_queda_reportado(ruta_pdf_ejemplo, registros_ejemplo, tmp_path, monkeypatch):
    from pathlib import Path
    import resaltado_pdf
    monkeypatch.setattr(resaltado_pdf, "MAXIMO_PAGINAS_POR_PDF", 0)
    resultado = resaltado_pdf.resaltar_por_cedula_sin_recortar([ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "s"))
    assert "demasiadas páginas" in resultado["errores_por_archivo"][Path(ruta_pdf_ejemplo).name]


def test_un_lote_que_tarda_demasiado_se_corta(ruta_pdf_ejemplo, registros_ejemplo, tmp_path, monkeypatch):
    from pathlib import Path
    import resaltado_pdf
    monkeypatch.setattr(resaltado_pdf, "MAXIMO_SEGUNDOS_POR_LOTE", -1)
    resultado = resaltado_pdf.resaltar_por_cedula_sin_recortar([ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "s"))
    assert "tiempo" in resultado["errores_por_archivo"][Path(ruta_pdf_ejemplo).name]


def test_host_falso_no_desvia_al_portal(cliente, monkeypatch):
    monkeypatch.setenv("VMA_HOSTS", "10.0.0.5,jupiter")
    monkeypatch.setitem(app.config, "SESION_VMA_SECRETO", "p" * 40)
    loc = cliente.get("/", base_url="http://evil.example:8001").headers["Location"]
    assert "evil.example" not in loc and loc.startswith("http://10.0.0.5:8080/")


def test_planillas_no_anuncia_version_del_servidor(cliente):
    assert cliente.get("/").headers.get("Server") == "Planillas"

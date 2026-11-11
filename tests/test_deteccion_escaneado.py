"""Detección de páginas escaneadas por cobertura de imagen (no solo por conteo de caracteres) y limpieza de la capa de
texto basura del escáner antes de hacer OCR.

El bug original: una planilla escaneada que traía la mala OCR del propio escáner se colaba como 'digital' (tenía texto,
aunque mal ubicado) y terminaba resaltando filas equivocadas. Ahora una imagen a página casi completa se trata como
escaneada aunque traiga texto, se le quita esa capa basura y se re-lee con el Tesseract empaquetado.
"""
import pymupdf as fitz
import pytest

import resaltado_pdf
from resaltado_pdf import _asegurar_buscable, _pagina_parece_escaneada

AN, AL = 595, 842


def _pagina_digital(doc, con_logo=False):
    p = doc.new_page(width=AN, height=AL)
    if con_logo:
        logo = fitz.open()
        lp = logo.new_page(width=120, height=60)
        lp.insert_text((5, 30), "LOGO", fontsize=20)
        pix = lp.get_pixmap()
        logo.close()
        p.insert_image(fitz.Rect(40, 20, 160, 80), pixmap=pix)
    p.insert_text((40, 100), "112340567 JUAN PEREZ MORA", fontsize=9)
    return p


def _doc_escaneado(con_basura=False):
    """Devuelve (doc, pagina) de un PDF de solo imagen; con_basura agrega una capa de texto invisible como la mala OCR
    del escáner (digitos que NO corresponden a las filas, para provocar el falso resaltado)."""
    tmp = fitz.open()
    p = tmp.new_page(width=AN, height=AL)
    p.insert_text((40, 100), "112340567 JUAN PEREZ MORA", fontsize=9)
    pix = p.get_pixmap(dpi=150)
    tmp.close()
    doc = fitz.open()
    ep = doc.new_page(width=pix.width, height=pix.height)
    ep.insert_image(ep.rect, pixmap=pix)
    if con_basura:
        ep.insert_text((50, 40), "Caja Costarricense Seguro Social 999888777 planilla varios",
                       fontsize=8, render_mode=3)
    return doc, ep


def test_digital_no_es_escaneada():
    doc = fitz.open()
    assert _pagina_parece_escaneada(_pagina_digital(doc)) is False
    doc.close()


def test_digital_con_logo_no_es_escaneada():
    doc = fitz.open()
    assert _pagina_parece_escaneada(_pagina_digital(doc, con_logo=True)) is False
    doc.close()


def test_imagen_sola_es_escaneada():
    doc, ep = _doc_escaneado()
    assert _pagina_parece_escaneada(ep) is True
    doc.close()


def test_imagen_con_capa_basura_es_escaneada():
    """El caso del bug: antes se colaba como digital por tener texto > umbral."""
    doc, ep = _doc_escaneado(con_basura=True)
    assert len(ep.get_text().strip()) > resaltado_pdf.UMBRAL_TEXTO_OCR  # de verdad tiene texto basura
    assert _pagina_parece_escaneada(ep) is True
    doc.close()


def test_ocr_quita_la_capa_basura_del_escaner(tmp_path, monkeypatch):
    """Con un Tesseract simulado (lee una palabra limpia), el OCR debe DEJAR la página con lo que leyó y SIN la basura
    del escáner -si no, esos dígitos falsos seguirían provocando resaltados equivocados."""
    doc, _ep = _doc_escaneado(con_basura=True)
    ruta = tmp_path / "escan.pdf"
    doc.save(str(ruta))
    doc.close()

    monkeypatch.setattr(resaltado_pdf, "_tesseract_exe", lambda: "tesseract-falso")
    monkeypatch.setattr(resaltado_pdf, "_palabras_ocr",
                        lambda _exe, _png: [("LIMPIO", 100.0, 200.0, 120.0, 40.0),
                                            ("CEDULA", 100.0, 240.0, 120.0, 40.0),
                                            ("112340567", 100.0, 280.0, 120.0, 40.0)])

    salida, uso = _asegurar_buscable(str(ruta), tmp_path / "buscables")
    assert uso is True
    d = fitz.open(salida)
    texto = d[0].get_text()
    d.close()
    assert "LIMPIO" in texto                 # quedó lo que leyó el OCR
    assert "999888777" not in texto          # y se borró la basura del escáner


def test_escaneada_sin_lectura_cuenta_para_el_aviso(tmp_path, monkeypatch):
    """Si la página es escaneada y el OCR no leyó nada útil, se cuenta para avisar al usuario (fallar de forma clara)."""
    doc, _ep = _doc_escaneado()
    ruta = tmp_path / "escan.pdf"
    doc.save(str(ruta))
    doc.close()

    monkeypatch.setattr(resaltado_pdf, "_tesseract_exe", lambda: "tesseract-falso")
    monkeypatch.setattr(resaltado_pdf, "_palabras_ocr", lambda _exe, _png: [])  # no leyó nada

    resaltado_pdf.reiniciar_contador_ocr()
    salida, uso = _asegurar_buscable(str(ruta), tmp_path / "buscables")
    assert uso is False
    assert resaltado_pdf.paginas_escaneadas_sin_leer() >= 1


def test_digital_no_dispara_ocr(tmp_path, monkeypatch):
    """Un PDF digital normal no debe entrar al OCR (ni siquiera intentar)."""
    doc = fitz.open()
    _pagina_digital(doc, con_logo=True)
    ruta = tmp_path / "digital.pdf"
    doc.save(str(ruta))
    doc.close()

    llamado = {"ocr": False}
    monkeypatch.setattr(resaltado_pdf, "_tesseract_exe", lambda: "tesseract-falso")

    def _no_deberia(_exe, _png):
        llamado["ocr"] = True
        return []

    monkeypatch.setattr(resaltado_pdf, "_palabras_ocr", _no_deberia)
    salida, uso = _asegurar_buscable(str(ruta), tmp_path / "buscables")
    assert salida == str(ruta)
    assert uso is False
    assert llamado["ocr"] is False


@pytest.mark.integracion
def test_integracion_escaneado_con_basura_resalta_solo_lo_correcto(tmp_path):
    """Punta a punta con el Tesseract real: un escaneo con capa basura del escáner se re-lee y el nombre buscado aparece
    (prueba que quitar la capa + OCR deja la imagen legible)."""
    from pathlib import Path
    raiz = Path(__file__).resolve().parent.parent
    tess = raiz / "tesseract" / "tesseract.exe"
    tessdata = raiz / "tesseract" / "tessdata"
    if not tess.is_file() or not (tessdata / "spa.traineddata").is_file():
        pytest.skip("falta el Tesseract empaquetado; no se puede probar el OCR real")
    import os
    os.environ["TESSERACT_EXE"] = str(tess)
    os.environ["TESSDATA_PREFIX"] = str(tessdata)
    from resaltado_pdf import resaltar_nombres_en_pdf

    doc, _ep = _doc_escaneado(con_basura=True)
    ruta = tmp_path / "escan_basura.pdf"
    doc.save(str(ruta))
    doc.close()

    res = resaltar_nombres_en_pdf(str(ruta), ["JUAN PEREZ MORA"], str(tmp_path / "out.pdf"))
    assert res.uso_ocr is True
    assert res.coincidencias_por_nombre["JUAN PEREZ MORA"] >= 1

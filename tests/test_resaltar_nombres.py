from resaltado_pdf import resaltar_nombres_en_pdf


def test_encuentra_un_nombre_que_si_esta_en_el_pdf(ruta_pdf_ejemplo, tmp_path):
    salida = tmp_path / "salida.pdf"
    resultado = resaltar_nombres_en_pdf(ruta_pdf_ejemplo, ["JUAN"], str(salida))
    assert resultado.coincidencias_por_nombre["JUAN"] > 0
    assert salida.exists()


def test_nombre_que_no_existe_da_cero_coincidencias(ruta_pdf_ejemplo, tmp_path):
    salida = tmp_path / "salida.pdf"
    resultado = resaltar_nombres_en_pdf(ruta_pdf_ejemplo, ["NOMBRE INEXISTENTE"], str(salida))
    assert resultado.coincidencias_por_nombre["NOMBRE INEXISTENTE"] == 0


def test_encuentra_apellido_con_ene_aunque_el_pdf_tenga_un_espacio(ruta_pdf_ejemplo, tmp_path):
    # el PDF tiene "MU OZ" en vez de "MUÑOZ" (bug de MNK), se busca con Ñ normal
    salida = tmp_path / "salida.pdf"
    resultado = resaltar_nombres_en_pdf(ruta_pdf_ejemplo, ["MUÑOZ"], str(salida))
    assert resultado.coincidencias_por_nombre["MUÑOZ"] > 0


def test_varios_nombres_a_la_vez(ruta_pdf_ejemplo, tmp_path):
    salida = tmp_path / "salida.pdf"
    resultado = resaltar_nombres_en_pdf(
        ruta_pdf_ejemplo, ["JUAN", "ZUNIGA", "NO EXISTE"], str(salida)
    )
    assert resultado.coincidencias_por_nombre["JUAN"] > 0
    assert resultado.coincidencias_por_nombre["ZUNIGA"] > 0
    assert resultado.coincidencias_por_nombre["NO EXISTE"] == 0


def test_reporta_tiempos_por_fase(ruta_pdf_ejemplo, tmp_path):
    salida = tmp_path / "salida.pdf"
    resultado = resaltar_nombres_en_pdf(ruta_pdf_ejemplo, ["JUAN"], str(salida))
    assert resultado.tiempos["conteos"]["paginas"] == 1
    assert "extraer_texto" in resultado.tiempos["segundos"]
    assert "guardar_disco" in resultado.tiempos["segundos"]


# --- Modo por nombre con recorte (resaltar_nombres_y_recortar) -------------
# La planilla de ejemplo (conftest.py) tiene 3 filas:
#   111111111 JUAN       PEREZ MORA
#   222222222 MARIA JOSE SOLANO MU OZ   (Ñ perdida, como en MNK)
#   333333333 LUIS       ZUNIGA RAMIREZ
# El PDF recortado lleva un resaltado por cada renglón copiado, así que
# contar resaltados = contar renglones que quedaron en el recorte.

import pymupdf as fitz

from resaltado_pdf import resaltar_nombres_y_recortar


def _resaltados(ruta_pdf):
    documento = fitz.open(ruta_pdf)
    total = sum(1 for pagina in documento for a in pagina.annots() or [] if a.type[1] == "Highlight")
    documento.close()
    return total


def test_recorte_por_nombre_deja_solo_los_renglones_buscados(ruta_pdf_ejemplo, tmp_path):
    resultado = resaltar_nombres_y_recortar(
        [ruta_pdf_ejemplo], ["JUAN", "MUÑOZ"], str(tmp_path / "salida"), formato="mnk",
    )

    assert resultado["ruta_pdf"] is not None
    assert _resaltados(resultado["ruta_pdf"]) == 2  # JUAN y MUÑOZ, LUIS queda fuera
    conteo = resultado["coincidencias_por_archivo"]["planilla_prueba.pdf"]
    assert conteo == {"JUAN": 1, "MUÑOZ": 1}
    assert resultado["errores_por_archivo"] == {}


def test_recorte_por_nombre_dos_nombres_en_el_mismo_renglon_salen_una_vez(ruta_pdf_ejemplo, tmp_path):
    resultado = resaltar_nombres_y_recortar(
        [ruta_pdf_ejemplo], ["JUAN", "PEREZ MORA"], str(tmp_path / "salida"), formato="mnk",
    )

    assert _resaltados(resultado["ruta_pdf"]) == 1
    assert resultado["coincidencias_por_archivo"]["planilla_prueba.pdf"] == {"JUAN": 1, "PEREZ MORA": 1}


def test_recorte_por_nombre_junta_todas_las_planillas_en_un_solo_pdf(ruta_pdf_ejemplo, tmp_path):
    import shutil

    segunda = tmp_path / "otra_planilla.pdf"
    shutil.copy(ruta_pdf_ejemplo, segunda)

    resultado = resaltar_nombres_y_recortar(
        [ruta_pdf_ejemplo, str(segunda)], ["LUIS"], str(tmp_path / "salida"), formato="mnk",
    )

    assert _resaltados(resultado["ruta_pdf"]) == 2  # un renglón de cada planilla
    assert list((tmp_path / "salida").glob("*.pdf")) == [tmp_path / "salida" / "Nombres encontrados.pdf"]
    assert resultado["coincidencias_por_archivo"]["otra_planilla.pdf"] == {"LUIS": 1}


def test_recorte_por_nombre_sin_coincidencias_no_genera_pdf(ruta_pdf_ejemplo, tmp_path):
    resultado = resaltar_nombres_y_recortar(
        [ruta_pdf_ejemplo], ["NOMBRE INEXISTENTE"], str(tmp_path / "salida"), formato="mnk",
    )

    assert resultado["ruta_pdf"] is None
    assert resultado["coincidencias_por_archivo"]["planilla_prueba.pdf"] == {"NOMBRE INEXISTENTE": 0}


def test_recorte_por_nombre_sin_resaltar_no_marca_amarillo(ruta_pdf_ejemplo, tmp_path):
    resultado = resaltar_nombres_y_recortar(
        [ruta_pdf_ejemplo], ["JUAN"], str(tmp_path / "salida"), formato="mnk", resaltar_filas=False,
    )

    assert resultado["ruta_pdf"] is not None
    assert _resaltados(resultado["ruta_pdf"]) == 0

"""Pruebas de resaltar_nombres_en_pdf con el PDF de ejemplo (datos
inventados, ver conftest.py)."""
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

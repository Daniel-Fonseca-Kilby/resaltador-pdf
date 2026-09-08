"""Pruebas de las utilidades de normalización de texto (no usan PDFs)."""
from api.index import _extraer_cedula_limpia, _indice_por_sinonimos
from resaltado_pdf import (
    _coincide_cliente,
    _nombre_archivo_seguro,
    _normalizar,
    _normalizar_cedula,
    _puede_aparecer_en_pagina,
    _todas_las_palabras_en_texto,
    _variantes_ene,
)


def test_normalizar_quita_tildes_y_pone_mayusculas():
    assert _normalizar("María López") == "MARIA LOPEZ"


def test_normalizar_ya_esta_normalizado():
    assert _normalizar("JUAN PEREZ") == "JUAN PEREZ"


def test_variantes_ene_vacio_si_no_hay_ene():
    assert _variantes_ene("PEREZ MORA") == []


def test_variantes_ene_reemplaza_por_espacio_y_por_nada():
    variantes = _variantes_ene("MUÑOZ")
    assert "MU OZ" in variantes
    assert "MUOZ" in variantes


def test_coincide_cliente_exacto():
    mapa = {"111111111": ["Cliente Uno"]}
    assert _coincide_cliente("111111111", mapa) == ["Cliente Uno"]


def test_coincide_cliente_con_digito_de_tipo_identificacion_antepuesto():
    # Algunas planillas (ej. CCSS) anteponen un dígito de tipo de
    # identificación al número real de cédula (ej. "0-303370238").
    mapa = {"303370238": ["Cliente Uno"]}
    assert _coincide_cliente("0303370238", mapa) == ["Cliente Uno"]


def test_coincide_cliente_no_encontrado():
    mapa = {"111111111": ["Cliente Uno"]}
    assert _coincide_cliente("999999999", mapa) == []


def test_coincide_cliente_devuelve_varios_clientes_para_la_misma_cedula():
    # Un oficial puede cubrir turnos en más de un cliente (relación 1 a N).
    mapa = {"111111111": ["Cliente Walmart", "Cliente BAC"]}
    assert _coincide_cliente("111111111", mapa) == ["Cliente Walmart", "Cliente BAC"]


def test_normalizar_cedula_quita_ceros_a_la_izquierda():
    assert _normalizar_cedula("010234056") == "10234056"


def test_normalizar_cedula_sin_ceros_no_cambia():
    assert _normalizar_cedula("10234056") == "10234056"


def test_coincide_cliente_tolera_cedula_sin_ceros_a_la_izquierda():
    # El Excel puede traer la cédula sin el cero inicial si la columna
    # quedó como celda numérica en vez de texto (ej. 8 dígitos en vez de 9).
    mapa = {"10234056": ["Cliente Uno"]}
    assert _coincide_cliente("010234056", mapa) == ["Cliente Uno"]


def test_extraer_cedula_limpia_elimina_decimal_cero():
    # BUSCARV/ERP suele devolver la cédula como float (303370238.0) cuando
    # la columna del Excel quedó como celda numérica en vez de texto.
    assert _extraer_cedula_limpia(303370238.0) == "303370238"
    assert _extraer_cedula_limpia("303370238.0") == "303370238"


def test_extraer_cedula_limpia_con_guiones():
    assert _extraer_cedula_limpia("0-30337-0238") == "0303370238"


def test_extraer_cedula_limpia_valor_vacio():
    assert _extraer_cedula_limpia(None) == ""
    assert _extraer_cedula_limpia("") == ""


def test_todas_las_palabras_en_texto_todas_presentes():
    texto = _normalizar("JUAN PEREZ MORA")
    assert _todas_las_palabras_en_texto(texto, ["PEREZ", "MORA"]) is True


def test_todas_las_palabras_en_texto_falta_una():
    texto = _normalizar("JUAN PEREZ MORA")
    assert _todas_las_palabras_en_texto(texto, ["PEREZ", "GONZALEZ"]) is False


def test_todas_las_palabras_en_texto_lista_vacia():
    assert _todas_las_palabras_en_texto(_normalizar("JUAN PEREZ"), []) is False


def test_todas_las_palabras_en_texto_tolera_ene_como_espacio():
    # el bug de MNK: "MUÑOZ" puede salir como "MU OZ" en el texto del PDF
    texto = _normalizar("SOLANO MU OZ")
    assert _todas_las_palabras_en_texto(texto, ["MUÑOZ"]) is True


def test_puede_aparecer_en_pagina_por_frase_completa():
    texto = _normalizar("JUAN PEREZ MORA aparece aqui")
    assert _puede_aparecer_en_pagina(texto, "Juan Perez Mora", ["PEREZ", "MORA"]) is True


def test_puede_aparecer_en_pagina_por_palabras_sueltas():
    # nombre y apellido en columnas separadas -la frase completa no está
    # junta, pero las palabras sí están todas en la página
    texto = _normalizar("JUAN   PEREZ MORA")
    assert _puede_aparecer_en_pagina(texto, "Juan Perez Mora", ["JUAN", "PEREZ", "MORA"]) is True


def test_puede_aparecer_en_pagina_ausente():
    texto = _normalizar("MARIA LOPEZ SOLANO")
    assert _puede_aparecer_en_pagina(texto, "Juan Perez Mora", ["JUAN", "PEREZ", "MORA"]) is False


def test_indice_por_sinonimos_prefiere_termino_mas_especifico():
    # Las planillas reales de VMA traen "Empresa" (unidad interna) Y
    # "Cliente" (a quién se factura) en el mismo Excel -"Empresa" está
    # antes en las columnas, pero no debe ganar solo por eso.
    encabezado = ["Empresa", "Identificacion", "Cliente"]
    assert _indice_por_sinonimos(encabezado, ["CLIENTE", "EMPRESA"]) == 2


def test_indice_por_sinonimos_usa_sinonimo_si_no_hay_termino_exacto():
    encabezado = ["Identificacion", "Empresa"]
    assert _indice_por_sinonimos(encabezado, ["CLIENTE", "EMPRESA"]) == 1


def test_indice_por_sinonimos_sin_coincidencias():
    assert _indice_por_sinonimos(["Fecha", "Monto"], ["CLIENTE", "EMPRESA"]) is None


def test_nombre_archivo_seguro_limpia_caracteres_invalidos():
    assert _nombre_archivo_seguro("Cliente/Prueba") == "Cliente_Prueba"


def test_nombre_archivo_seguro_vacio_usa_valor_por_defecto():
    assert _nombre_archivo_seguro("   ") == "SIN_CLIENTE"

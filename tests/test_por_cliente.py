"""Pruebas de resaltar_por_cedula_y_exportar_por_cliente (datos inventados, ver conftest.py)."""
from pathlib import Path

import pymupdf as fitz

from resaltado_pdf import (
    _extender_leyenda_para_incluir_total,
    _franja_leyenda_en_pagina,
    _franja_total_en_pagina,
    _franjas_pie_de_pagina,
    _recortar_total_antes_de_leyenda,
    _techo_de_datos,
    resaltar_por_cedula_y_exportar_por_cliente,
)


def _escribir_fila(pagina, y, celdas, fontsize=10):
    """Escribe una fila de texto en columnas separadas horizontalmente,
    imitando cómo una planilla real reparte el texto en columnas anchas."""
    x = 36
    for texto, ancho in celdas:
        pagina.insert_text((x, y), texto, fontsize=fontsize)
        x += ancho


_TITULOS_COLUMNAS = [
    ("IDENTIFICACION", 90), ("NOMBRE", 80), ("APELLIDOS", 100), ("OBSERVACION", 90),
]


def _crear_pdf_planilla(ruta, empresa, filas_por_pagina, espacio_filas=30):
    """Arma una planilla con una página por cada lista de filas de
    'filas_por_pagina', repitiendo encabezado (empresa + títulos de
    columna) en cada página -como pasa en una planilla real de varias
    hojas para la misma póliza."""
    documento = fitz.open()
    for filas in filas_por_pagina:
        pagina = documento.new_page(width=595, height=842)
        pagina.insert_text((36, 40), empresa, fontsize=13)
        _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
        y = 140
        for cedula, nombre, apellidos, obs in filas:
            _escribir_fila(pagina, y, [(cedula, 90), (nombre, 80), (apellidos, 100), (obs, 90)])
            y += espacio_filas
    documento.save(str(ruta))
    documento.close()


def test_exporta_un_pdf_por_cliente(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    carpeta_salida = tmp_path / "salida_por_cliente"

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_ejemplo, str(carpeta_salida), formato="mnk"
    )

    assert set(resultado["archivos_por_cliente"]) == {"Cliente Prueba Uno", "Cliente Prueba Dos"}
    for ruta in resultado["archivos_por_cliente"].values():
        assert Path(ruta).exists()


def test_reporta_cedulas_que_no_aparecen_en_el_pdf(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    carpeta_salida = tmp_path / "salida_por_cliente"

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_ejemplo, str(carpeta_salida), formato="mnk"
    )

    cedulas_no_encontradas = {r["cedula"] for r in resultado["no_encontrados"]}
    assert cedulas_no_encontradas == {"999999999"}


def test_detalle_registros_marca_encontrado_y_poliza_de_origen(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    carpeta_salida = tmp_path / "salida_por_cliente"

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_ejemplo, str(carpeta_salida), formato="mnk"
    )

    detalle_por_cedula = {d["cedula"]: d for d in resultado["detalle_registros"]}

    encontrado = detalle_por_cedula["111111111"]
    assert encontrado["encontrado"] is True
    assert encontrado["polizas"] == [Path(ruta_pdf_ejemplo).name]

    no_encontrado = detalle_por_cedula["999999999"]
    assert no_encontrado["encontrado"] is False
    assert no_encontrado["polizas"] == []


def test_no_reporta_errores_con_un_pdf_valido(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    carpeta_salida = tmp_path / "salida_por_cliente"

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_ejemplo, str(carpeta_salida), formato="mnk"
    )

    assert resultado["errores_por_archivo"] == {}


def test_extranjero_se_encuentra_por_numero_de_asegurado_si_el_dimex_no_aparece(tmp_path):
    """La CCSS imprime a un extranjero bajo su número de asegurado de la
    Caja, no bajo el DIMEX que trae el Excel. Si el registro trae
    'numero_asegurado' y ese número sí aparece en el PDF (aunque el DIMEX
    no aparezca en ningún lado), debe encontrarse en la misma pasada -sin
    necesitar el rescate por nombre- y quedar reportado bajo su cédula
    real (el DIMEX), no bajo el número de asegurado."""
    ruta_poliza = tmp_path / "poliza_ccss.pdf"
    _crear_pdf_planilla(
        ruta_poliza,
        "EMPRESA CCSS EXTRANJERO",
        filas_por_pagina=[[("905550003", "MARIA", "GOMEZ TORRES", "Ninguna")]],
    )
    registros = [
        {
            "cedula": "155812345678",  # DIMEX -no aparece en el PDF
            "cliente": "Cliente Extranjero",
            "nombre": "Maria Gomez Torres",
            "numero_asegurado": "905550003",  # el número que sí imprime la CCSS
        },
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="ccss"
    )

    assert resultado["no_encontrados"] == []
    detalle = resultado["detalle_registros"][0]
    assert detalle["encontrado"] is True
    assert detalle["encontrado_por"] == "numero_asegurado"
    # se reporta bajo la cédula real (DIMEX), no bajo el número de asegurado
    assert detalle["cedula"] == "155812345678"

    documento = fitz.open(resultado["archivos_por_cliente"]["Cliente Extranjero"])
    try:
        assert "GOMEZ" in documento[0].get_text()
    finally:
        documento.close()


def test_numero_de_asegurado_vacio_no_afecta_la_busqueda_normal_por_cedula(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    """Si el registro no trae número de asegurado (columna vacía o
    ausente, el caso normal para un nacional), todo debe seguir
    funcionando exactamente igual que antes -por cédula."""
    carpeta_salida = tmp_path / "salida_por_cliente"

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_ejemplo, str(carpeta_salida), formato="mnk"
    )

    detalle_por_cedula = {d["cedula"]: d for d in resultado["detalle_registros"]}
    assert detalle_por_cedula["111111111"]["encontrado_por"] == "cedula"


def test_pdf_con_contrasena_se_reporta_como_error_sin_tumbar_el_proceso(tmp_path):
    documento = fitz.open()
    documento.new_page()
    ruta = tmp_path / "protegido.pdf"
    documento.save(str(ruta), encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="clave123")
    documento.close()

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta)],
        [{"cedula": "111111111", "cliente": "Cliente X"}],
        str(tmp_path / "salida"),
    )

    assert "protegido.pdf" in resultado["errores_por_archivo"]
    assert resultado["archivos_por_cliente"] == {}


def test_techo_de_datos_queda_entre_el_encabezado_y_la_primera_fila(ruta_pdf_ejemplo):
    # Deben coincidir con Y_TITULOS y Y_PRIMERA_FILA en conftest.py.
    y_titulos = 100
    y_primera_fila = 140

    documento = fitz.open(ruta_pdf_ejemplo)
    pagina = documento[0]

    techo = _techo_de_datos(pagina, formato="mnk")

    assert techo is not None
    assert y_titulos < techo < y_primera_fila
    documento.close()


def test_encabezado_una_sola_vez_si_la_poliza_cabe_en_una_hoja(tmp_path):
    """Un mismo oficial partido en 3 páginas del mismo PDF (misma póliza)
    solo debe arrastrar el encabezado una vez, si las filas caben en una
    sola hoja de salida."""
    ruta_poliza = tmp_path / "poliza_a.pdf"
    _crear_pdf_planilla(
        ruta_poliza,
        "EMPRESA POLIZA A",
        filas_por_pagina=[
            [("111111111", "JUAN", "PEREZ MORA", "Ninguna")],
            [("222222222", "MARIA", "SOLANO MORA", "Ninguna")],
            [("333333333", "LUIS", "ZUNIGA RAMIREZ", "Ninguna")],
        ],
    )
    registros = [
        {"cedula": "111111111", "cliente": "Cliente Unico", "nombre": "Juan Perez"},
        {"cedula": "222222222", "cliente": "Cliente Unico", "nombre": "Maria Solano"},
        {"cedula": "333333333", "cliente": "Cliente Unico", "nombre": "Luis Zuniga"},
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento = fitz.open(resultado["archivos_por_cliente"]["Cliente Unico"])
    try:
        assert documento.page_count == 1
        assert documento[0].get_text().count("EMPRESA POLIZA A") == 1
    finally:
        documento.close()


def test_encabezado_sale_de_la_pagina_1_aunque_el_cliente_empiece_en_la_pagina_2(tmp_path):
    """Reproduce un reporte real de varias páginas (ej. MNK): la página 1
    trae el logo/título/fecha completos, las páginas siguientes solo
    repiten la fila de títulos de columna, más arriba. Si el primer oficial
    de un cliente aparece recién en la página 2, igual se debe llevar el
    encabezado COMPLETO de la página 1 -no el recortado de su propia
    página, que le faltaría el logo/título/fecha."""
    documento = fitz.open()

    # página 1: encabezado completo (letterhead) + una fila de otro cliente
    pagina1 = documento.new_page(width=595, height=842)
    pagina1.insert_text((36, 40), "REPORTE MNK - LETTERHEAD COMPLETO", fontsize=13)
    pagina1.insert_text((36, 60), "Fecha: 01/01/2026", fontsize=10)
    _escribir_fila(pagina1, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina1, 140, [("999999999", "OTRO", "CLIENTE VIEJO", "Ninguna")])

    # página 2: sin letterhead, solo repite la fila de títulos más arriba
    pagina2 = documento.new_page(width=595, height=842)
    _escribir_fila(pagina2, 40, _TITULOS_COLUMNAS)
    _escribir_fila(pagina2, 80, [("111111111", "JUAN", "PEREZ NUEVO", "Ninguna")])

    ruta_poliza = tmp_path / "poliza_multipagina.pdf"
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Pagina Dos", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    ruta_salida = resultado["archivos_por_cliente"]["Cliente Pagina Dos"]
    documento_salida = fitz.open(ruta_salida)
    try:
        texto = documento_salida[0].get_text()
        assert "REPORTE MNK - LETTERHEAD COMPLETO" in texto
        assert "JUAN" in texto
    finally:
        documento_salida.close()


def test_pie_de_pagina_con_total_se_agrega_al_final_del_documento(tmp_path):
    """El pie de página con el total (tal cual viene en el original, ver
    _PERFILES_PIE_PAGINA) se agrega al final del PDF de cada cliente -como
    imagen, no como copia vectorial: así se lleva también el valor de
    campos de formulario (ej. el total real de CCSS, que la Oficina
    Virtual rellena como widget en vez de como texto de la página), que
    show_pdf_page no arrastra. Por eso se verifica que se insertó una
    imagen, no que el texto se pueda extraer."""
    ruta_poliza = tmp_path / "poliza_con_total.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA CON TOTAL", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina.insert_text((36, 700), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina.insert_text((36, 720), "TOTAL DE SALARIO 405,710.71", fontsize=10)
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Con Total", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Con Total"])
    try:
        total_imagenes = sum(len(pagina.get_images(full=True)) for pagina in documento_salida)
        assert total_imagenes >= 1
    finally:
        documento_salida.close()


def test_pie_de_pagina_separa_total_de_leyenda_sin_el_hueco_del_medio(tmp_path):
    """Si el original trae un hueco en blanco grande entre el total y la
    leyenda/firma (como en CCSS real), se recortan como DOS franjas
    separadas -no se arrastra ese hueco como un solo bloque enorme."""
    ruta_poliza = tmp_path / "poliza_con_hueco.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA CON HUECO", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina.insert_text((36, 200), "TOTAL SALARIOS", fontsize=10)
    # hueco grande en blanco entre el total (y=200) y la leyenda (y=800),
    # igual que en el documento real de la CCSS
    pagina.insert_text((36, 800), "Ajuste al minimo base diferenciada SEM", fontsize=8)
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Con Hueco", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="ccss"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Con Hueco"])
    try:
        total_imagenes = sum(len(p.get_images(full=True)) for p in documento_salida)
        assert total_imagenes == 2  # una franja para el total, otra para la leyenda
    finally:
        documento_salida.close()


def test_pie_de_pagina_de_cada_poliza_queda_junto_a_sus_propias_filas(tmp_path):
    """Si el cliente tiene oficiales en dos pólizas distintas, el total de
    la póliza 1 tiene que quedar junto a SUS propias filas -no amontonado
    con el de la póliza 2 al final de todo el documento."""
    ruta_poliza_a = tmp_path / "poliza_a.pdf"
    ruta_poliza_b = tmp_path / "poliza_b.pdf"

    documento_a = fitz.open()
    pagina_a = documento_a.new_page(width=595, height=842)
    pagina_a.insert_text((36, 40), "EMPRESA POLIZA A", fontsize=13)
    _escribir_fila(pagina_a, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina_a, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina_a.insert_text((36, 200), "TOTAL SALARIOS", fontsize=10)
    documento_a.save(str(ruta_poliza_a))
    documento_a.close()

    documento_b = fitz.open()
    pagina_b = documento_b.new_page(width=595, height=842)
    pagina_b.insert_text((36, 40), "EMPRESA POLIZA B", fontsize=13)
    _escribir_fila(pagina_b, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina_b, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina_b.insert_text((36, 200), "TOTAL SALARIOS", fontsize=10)
    documento_b.save(str(ruta_poliza_b))
    documento_b.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Dos Polizas", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza_a), str(ruta_poliza_b)], registros, str(tmp_path / "salida"), formato="ccss"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Dos Polizas"])
    try:
        # cada póliza (encabezado + fila + su propio total) cabe en su
        # propia hoja -no queda un total separado al final de todo
        assert documento_salida.page_count == 2

        texto_pagina_1 = documento_salida[0].get_text()
        assert "EMPRESA POLIZA A" in texto_pagina_1
        assert "EMPRESA POLIZA B" not in texto_pagina_1
        assert len(documento_salida[0].get_images(full=True)) >= 1

        texto_pagina_2 = documento_salida[1].get_text()
        assert "EMPRESA POLIZA B" in texto_pagina_2
        assert "EMPRESA POLIZA A" not in texto_pagina_2
        assert len(documento_salida[1].get_images(full=True)) >= 1
    finally:
        documento_salida.close()


def test_varias_polizas_seguidas_conservan_todo_el_contenido_tras_guardarse_a_disco(tmp_path):
    """Por memoria, el PDF de cada cliente se guarda a disco y se libera de
    RAM en cada cambio de póliza (ver _flush_a_disco), en vez de mantenerse
    completo en memoria hasta el final -esto obliga a reabrir el archivo y
    seguir agregándole páginas con guardados incrementales. Con tres
    pólizas seguidas para el mismo cliente se ejercita ese guardado
    incremental más de una vez seguida, para asegurar que ninguna página
    anterior se pierda ni se corrompa por el camino."""
    rutas = []
    for letra in ("A", "B", "C"):
        ruta = tmp_path / f"poliza_{letra.lower()}.pdf"
        documento = fitz.open()
        pagina = documento.new_page(width=595, height=842)
        pagina.insert_text((36, 40), f"EMPRESA POLIZA {letra}", fontsize=13)
        _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
        _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
        pagina.insert_text((36, 200), "TOTAL SALARIOS", fontsize=10)
        documento.save(str(ruta))
        documento.close()
        rutas.append(str(ruta))

    registros = [{"cedula": "111111111", "cliente": "Cliente Tres Polizas", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        rutas, registros, str(tmp_path / "salida"), formato="ccss"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Tres Polizas"])
    try:
        assert documento_salida.page_count == 3
        for indice, letra in enumerate(("A", "B", "C")):
            texto_pagina = documento_salida[indice].get_text()
            assert f"EMPRESA POLIZA {letra}" in texto_pagina
            for otra_letra in ("A", "B", "C"):
                if otra_letra != letra:
                    assert f"EMPRESA POLIZA {otra_letra}" not in texto_pagina
            assert len(documento_salida[indice].get_images(full=True)) >= 1
    finally:
        documento_salida.close()


def test_pie_de_pagina_se_comparte_entre_clientes_de_la_misma_poliza(tmp_path):
    """El render del pie de página se cachea por archivo (para no
    reabrirlo/re-renderizarlo por cada cliente que comparte la misma
    póliza) -pero cada cliente igual debe llevarse su propia copia en su
    PDF, sin que el cacheo se lo pierda ni se lo mezcle con otro."""
    ruta_poliza = tmp_path / "poliza_compartida.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA COMPARTIDA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    _escribir_fila(pagina, 170, [("222222222", 90), ("MARIA", 80), ("SOLANO MORA", 100), ("Ninguna", 90)])
    pagina.insert_text((36, 220), "TOTAL SALARIOS", fontsize=10)
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [
        {"cedula": "111111111", "cliente": "Cliente Uno Compartido", "nombre": "Juan Perez"},
        {"cedula": "222222222", "cliente": "Cliente Dos Compartido", "nombre": "Maria Solano"},
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="ccss"
    )

    for cliente in ("Cliente Uno Compartido", "Cliente Dos Compartido"):
        documento_salida = fitz.open(resultado["archivos_por_cliente"][cliente])
        try:
            total_imagenes = sum(len(p.get_images(full=True)) for p in documento_salida)
            assert total_imagenes >= 1
        finally:
            documento_salida.close()


def test_pie_de_pagina_no_se_agrega_si_el_formato_no_tiene_perfil_conocido(tmp_path):
    """Si el formato no tiene un perfil de pie de página conocido (ej.
    INS), no se agrega nada -mejor omitirlo que recortar cualquier cosa."""
    ruta_poliza = tmp_path / "poliza_sin_perfil.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA SIN PERFIL", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Sin Perfil", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="ins"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Sin Perfil"])
    try:
        assert documento_salida.page_count == 1
    finally:
        documento_salida.close()


def test_pie_de_pagina_usa_la_ultima_fila_si_el_texto_del_total_no_se_encuentra(tmp_path):
    """Algunos PDFs reales de MNK no traen el rótulo del total (ni
    "CODIFICACIÓN") como texto buscable con search_for, aunque se vean
    perfectamente al abrir el archivo. En ese caso se usa la posición de
    la ÚLTIMA fila de empleado real (por su número de identificación)
    como referencia: todo lo que hay debajo de ella, hasta el final de
    la hoja, se copia como cierre de la póliza, sea o no texto buscable
    con las anclas conocidas."""
    ruta_poliza = tmp_path / "poliza_sin_texto_de_total.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA SIN TEXTO DE TOTAL", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    # cierre que no calza con ninguna ancla conocida de total/leyenda
    pagina.insert_text((36, 200), "RESUMEN FINAL DE LA PLANILLA - 1 REGISTRO", fontsize=10)
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Sin Texto Total", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Sin Texto Total"])
    try:
        total_imagenes = sum(len(p.get_images(full=True)) for p in documento_salida)
        assert total_imagenes >= 1
    finally:
        documento_salida.close()


def test_respaldo_posicional_no_arrastra_el_espacio_en_blanco_hasta_el_fondo_de_la_hoja(tmp_path):
    """Igual que la prueba anterior (texto de total no buscable), pero en
    una hoja donde el contenido real termina bien antes del final de la
    página -como pasa en pólizas grandes reales donde el total/codificación
    quedan a media hoja y el resto queda en blanco. El respaldo posicional
    no debe arrastrar ese espacio vacío hasta el borde físico de la hoja:
    debe cortar justo después del último texto real, para que la
    leyenda/firma de la hoja siguiente pueda quedar pegada en el mismo
    bloque de salida en vez de dejar un salto de página feo en el medio."""
    ruta_poliza = tmp_path / "poliza_con_hueco_grande.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA CON HUECO GRANDE", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    # cierre que no calza con ninguna ancla conocida, y termina bien
    # arriba -el resto de la hoja (hasta y=842) queda en blanco de verdad
    pagina.insert_text((36, 200), "RESUMEN FINAL DE LA PLANILLA - 1 REGISTRO", fontsize=10)
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Hueco Grande", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Hueco Grande"])
    try:
        alto_maximo_imagen_pt = 0.0
        for pagina_salida in documento_salida:
            for xref, *_resto in pagina_salida.get_images(full=True):
                pixmap_img = fitz.Pixmap(documento_salida.extract_image(xref)["image"])
                # la imagen se renderiza a 2x (fitz.Matrix(2, 2)), así que
                # se divide entre 2 para volver a puntos de PDF
                alto_maximo_imagen_pt = max(alto_maximo_imagen_pt, pixmap_img.height / 2)

        # el texto real termina cerca de y=210; si se arrastrara hasta el
        # fondo de la hoja (842) la imagen tendría más de 600pt de alto.
        # Cortando pegado al contenido debe quedar muy por debajo de eso.
        assert alto_maximo_imagen_pt < 100
    finally:
        documento_salida.close()


def test_franja_del_total_no_incluye_encabezado_repetido_pegado_arriba():
    """Cuando a la última hoja de una póliza le quedan pocas filas, CCSS
    repite el renglón de títulos de columna justo antes del total. La
    franja recortada para el pie de página no debe incluir ese renglón
    -si lo hiciera, saldría un pedazo de encabezado repetido y
    desordenado justo arriba del total."""
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    # renglón de títulos repetido, muy pegado al total -el caso real que
    # se ve cuando a la hoja le quedan pocas filas
    _escribir_fila(pagina, 150, _TITULOS_COLUMNAS)
    pagina.insert_text((36, 168), "TOTAL SALARIOS", fontsize=10)

    franjas = _franjas_pie_de_pagina(pagina, formato="ccss")

    assert len(franjas) == 1
    rect_encabezado_repetido = pagina.search_for("OBSERVACION")[-1]
    assert franjas[0].y0 >= rect_encabezado_repetido.y1


def test_pie_de_pagina_cuando_el_total_se_repite_en_hoja_de_aviso_legal(tmp_path):
    """MNK reparte el pie de página real en dos hojas: la que trae los
    datos + el total + "CODIFICACIÓN", y una hoja de aviso legal/firma
    aparte donde el total se repite pero "CODIFICACIÓN" no aparece. El
    total y la leyenda deben capturarse cada uno de la última hoja donde
    de verdad aparecen -no asumir que están juntos, ni duplicar el total
    solo porque sale en las dos hojas."""
    ruta_poliza = tmp_path / "poliza_dos_hojas.pdf"
    documento = fitz.open()

    pagina_datos = documento.new_page(width=595, height=842)
    pagina_datos.insert_text((36, 40), "EMPRESA DOS HOJAS", fontsize=13)
    _escribir_fila(pagina_datos, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina_datos, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina_datos.insert_text((36, 200), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina_datos.insert_text((36, 220), "TOTAL DE SALARIO 405710.71", fontsize=10)
    pagina_datos.insert_text((36, 250), "CODIFICACIÓN", fontsize=10)

    pagina_aviso = documento.new_page(width=595, height=842)
    pagina_aviso.insert_text((36, 40), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina_aviso.insert_text((36, 60), "TOTAL DE SALARIO 405710.71", fontsize=10)
    pagina_aviso.insert_text((36, 100), "La documentacion contractual y la nota tecnica...", fontsize=8)

    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Dos Hojas", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Dos Hojas"])
    try:
        imagenes = [xref for p in documento_salida for xref, *_r in p.get_images(full=True)]
        # una imagen para el total (de la hoja de aviso, la última donde
        # aparece) y una para la leyenda (de la hoja de datos, la única
        # con "CODIFICACIÓN") -nunca deben salir tres o más
        assert len(imagenes) == 2

        # la de "CODIFICACIÓN" (con la firma) viene de la PRIMERA página del
        # original (pagina_datos, índice 0); la del total repetido viene de
        # la SEGUNDA (pagina_aviso, índice 1) -aunque el bloque del total se
        # calcule primero en el código, en la salida debe listarse después
        # de la leyenda, respetando el orden real de las páginas. Como la
        # franja de la leyenda llega hasta el fondo de la hoja (mucho más
        # alta que la del total, que es solo el renglón), se distinguen por
        # su alto.
        alturas = [fitz.Pixmap(documento_salida.extract_image(xref)["image"]).height for xref in imagenes]
        assert alturas[0] > alturas[1], (
            "la leyenda (más alta, de la página 1) debe ir antes que el total repetido "
            "(más bajo, de la página 2) -no al revés"
        )
    finally:
        documento_salida.close()


def test_extender_leyenda_para_incluir_total_agrega_el_total_completo_arriba():
    """En MNK "CODIFICACIÓN" a veces viene tan pegada al total que no hay
    margen seguro para separarlos sin arriesgarse a cortar algo -en vez de
    intentar una división perfecta, se prefiere que la leyenda vuelva a
    incluir el total completo en su propio techo (el total ya se agregó
    también como su propio bloque aparte): mejor que se vea duplicado a
    que falte contenido. Así pedido explícitamente por VMA."""
    franja_total = fitz.Rect(0, 200, 595, 248)
    franja_leyenda = fitz.Rect(0, 215, 595, 800)

    extendida = _extender_leyenda_para_incluir_total(franja_leyenda, franja_total, misma_pagina=True)

    assert extendida.y0 == franja_total.y0
    assert extendida.y1 == franja_leyenda.y1


def test_extender_leyenda_para_incluir_total_no_toca_franja_que_ya_lo_incluye():
    """Si la leyenda ya arrancaba antes que el total (o están en páginas
    distintas), no hay nada que extender."""
    franja_total = fitz.Rect(0, 200, 595, 248)
    franja_leyenda_ya_incluye_el_total = fitz.Rect(0, 190, 595, 800)

    extendida = _extender_leyenda_para_incluir_total(
        franja_leyenda_ya_incluye_el_total, franja_total, misma_pagina=True
    )
    assert extendida == franja_leyenda_ya_incluye_el_total

    # en páginas distintas tampoco tiene sentido extender una contra la otra
    franja_leyenda_otra_pagina = fitz.Rect(0, 215, 595, 800)
    extendida_paginas_distintas = _extender_leyenda_para_incluir_total(
        franja_leyenda_otra_pagina, franja_total, misma_pagina=False
    )
    assert extendida_paginas_distintas == franja_leyenda_otra_pagina


def test_recortar_total_antes_de_leyenda_evita_tragarse_la_codificacion():
    """Caso real de MNK: "CODIFICACIÓN" viene pegada casi sin espacio bajo
    el total -el margen fijo de abajo del total (_MARGEN_ABAJO_TOTAL, 26pt)
    se pasa de largo y termina capturando la barra de "CODIFICACIÓN" entera
    dentro del recorte del total, dejándola faltante en el de la leyenda.
    Debe recortarse el total para que pare justo donde arranca la leyenda."""
    # el total "crudo" (con su margen fijo de 26pt) alcanza hasta y=239,
    # pero la leyenda arranca en y=220 -mucho antes de que el margen del
    # total termine
    franja_total_con_margen_generoso = fitz.Rect(0, 190, 842, 239)
    franja_leyenda = fitz.Rect(0, 220, 842, 595)

    recortada = _recortar_total_antes_de_leyenda(franja_total_con_margen_generoso, franja_leyenda, misma_pagina=True)

    assert recortada.y1 == franja_leyenda.y0
    assert recortada.y0 == franja_total_con_margen_generoso.y0  # nunca toca su propio inicio


def test_recortar_total_antes_de_leyenda_no_toca_franja_que_no_se_superpone():
    """Si el margen del total ya paraba antes de donde arranca la leyenda,
    no hay nada que recortar."""
    franja_total = fitz.Rect(0, 190, 842, 210)
    franja_leyenda = fitz.Rect(0, 220, 842, 595)

    recortada = _recortar_total_antes_de_leyenda(franja_total, franja_leyenda, misma_pagina=True)
    assert recortada == franja_total

    # en páginas distintas tampoco hay nada que recortar, aunque las
    # coordenadas coincidan
    franja_total_otra_pagina = fitz.Rect(0, 190, 842, 239)
    recortada_paginas_distintas = _recortar_total_antes_de_leyenda(
        franja_total_otra_pagina, franja_leyenda, misma_pagina=False
    )
    assert recortada_paginas_distintas == franja_total_otra_pagina


def test_recortar_total_antes_de_leyenda_nunca_corta_su_propia_etiqueta():
    """Caso extremo: si la leyenda arranca ANTES incluso de donde empieza
    el total (algo raro, pero si pasara), no hay que recortar el total
    hasta dejarlo con alto cero o negativo -mejor dejarlo como está."""
    franja_total = fitz.Rect(0, 190, 842, 239)
    franja_leyenda_antes_del_total = fitz.Rect(0, 185, 842, 595)

    recortada = _recortar_total_antes_de_leyenda(franja_total, franja_leyenda_antes_del_total, misma_pagina=True)
    assert recortada == franja_total


def test_pie_de_pagina_duplica_el_total_cuando_esta_pegado_a_codificacion(tmp_path):
    """Prueba de punta a punta del caso real de MNK: "CODIFICACIÓN" pegada
    casi sin espacio bajo el total. El total propio no debe tragarse la
    barra de "CODIFICACIÓN" (se recorta antes de ella), y la leyenda debe
    volver a incluir el total completo en su propio techo -aunque salga
    duplicado, es preferible a perder contenido. Se compara contra lo que
    las funciones reales calculan de forma independiente sobre la misma
    página, así la prueba no depende de adivinar métricas de fuente a mano."""
    ruta_poliza = tmp_path / "poliza_total_pegado.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina.insert_text((36, 200), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina.insert_text((36, 220), "TOTAL DE SALARIO 405710.71", fontsize=10)
    # pegada casi sin espacio -bastante antes de que termine el margen fijo
    # de abajo del total (_MARGEN_ABAJO_TOTAL, 26pt desde su último renglón)
    pagina.insert_text((36, 232), "CODIFICACIÓN", fontsize=10)
    documento.save(str(ruta_poliza))
    documento.close()

    # referencia independiente: lo que las funciones reales calculan para
    # esta misma página, antes de pasar por el pipeline completo
    documento_ref = fitz.open(ruta_poliza)
    pagina_ref = documento_ref[0]
    franja_total_ref = _franja_total_en_pagina(pagina_ref, "mnk")
    franja_leyenda_ref = _franja_leyenda_en_pagina(pagina_ref, "mnk")
    alto_hoja = pagina_ref.rect.height
    documento_ref.close()
    assert franja_total_ref is not None and franja_leyenda_ref is not None
    # confirma que este PDF de prueba de verdad reproduce el caso "pegado"
    assert franja_leyenda_ref.y0 < franja_total_ref.y1

    registros = [{"cedula": "111111111", "cliente": "Cliente Total Pegado", "nombre": "Juan Perez"}]
    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Total Pegado"])
    try:
        imagenes = [xref for p in documento_salida for xref, *_r in p.get_images(full=True)]
        assert len(imagenes) == 2

        # las imágenes se guardan a 2x de escala (ver _pixmaps_pie_de_poliza)
        alto_total, alto_leyenda = (
            fitz.Pixmap(documento_salida.extract_image(xref)["image"]).height / 2 for xref in imagenes
        )

        # el total propio no debe llegar hasta "CODIFICACIÓN"
        assert alto_total < (franja_leyenda_ref.y0 - franja_total_ref.y0) + 1

        # la leyenda debe volver a incluir el total completo -su alto debe
        # acercarse a la distancia desde el techo del total hasta el fondo
        # de la hoja, no solo desde "CODIFICACIÓN"
        alto_esperado_leyenda = alto_hoja - franja_total_ref.y0
        assert alto_leyenda > alto_esperado_leyenda - 5
    finally:
        documento_salida.close()


def test_techo_de_datos_no_confunde_numero_patronal_con_primera_fila():
    """En CCSS, el número patronal (arriba del todo, ANTES del título de
    columnas) también tiene 9+ dígitos -no debe confundirse con la
    primera fila de un empleado ni recortar el encabezado de más de lo
    debido."""
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 30), "PLANILLA MENSUAL", fontsize=13)
    # número patronal, con guiones, arriba del título de columnas
    pagina.insert_text((36, 60), "2-03101682626-001-001", fontsize=10)
    _escribir_fila(pagina, 100, [("APELLIDOS Y NOMBRES", 200), ("OBSERVACIONES", 100)])

    techo = _techo_de_datos(pagina, formato="ccss")

    assert techo is not None
    # debe quedar cerca del título de columnas (y=100), no arriba, cerca
    # del número patronal (y=60)
    assert techo > 90


def test_encabezado_no_arrastra_la_primera_fila_de_datos_si_esta_muy_pegada(tmp_path):
    """Si la primera fila de datos de la página queda muy pegada al
    encabezado (menos que el margen fijo que usa _techo_de_datos), el
    bloque de encabezado que se repite en el PDF de cada cliente no debe
    arrastrar al primer empleado de esa página -aunque no sea la persona
    buscada ni pertenezca a este cliente. Esto reproduce el caso real
    donde el primer empleado de la póliza (ej. "Gamboa") aparecía pegado
    justo después del encabezado en el PDF de CUALQUIER cliente de esa
    póliza, sin importar a quién se buscara."""
    ruta_poliza = tmp_path / "poliza_header_apretado.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA HEADER APRETADO", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    # primer empleado de la página, pegado al encabezado -no es la
    # persona buscada ni está en el Excel de este cliente
    _escribir_fila(pagina, 112, [("999999999", 90), ("PRIMERO", 80), ("GAMBOA", 100), ("Ninguna", 90)])
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    documento.save(str(ruta_poliza))
    documento.close()

    registros = [{"cedula": "111111111", "cliente": "Cliente Header Apretado", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Header Apretado"])
    try:
        texto_completo = "".join(p.get_text() for p in documento_salida)
        assert "GAMBOA" not in texto_completo
        assert "PEREZ MORA" in texto_completo
    finally:
        documento_salida.close()


def test_franja_del_total_no_incluye_la_ultima_fila_de_datos_pegada_arriba():
    """Igual que con el encabezado repetido, el margen fijo del total
    también puede pasarse de largo hacia la ÚLTIMA fila de datos real de
    la página (no un encabezado repetido, sino un empleado cualquiera) si
    quedan muy pegados, como en tablas de MNK con filas apretadas."""
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    # el total queda pegado a la última fila de datos (pero con margen
    # suficiente para no tener que elegir entre cortar la etiqueta del
    # total o arrastrar la fila -ver la otra prueba para el caso límite)
    pagina.insert_text((36, 156), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina.insert_text((36, 176), "TOTAL DE SALARIO 405710.71", fontsize=10)

    franja_total = _franja_total_en_pagina(pagina, formato="mnk")

    assert franja_total is not None
    rect_ultima_fila = pagina.search_for("PEREZ MORA")[0]
    assert franja_total.y0 >= rect_ultima_fila.y1


def test_franja_del_total_nunca_corta_su_propia_etiqueta(tmp_path):
    """Si la última fila de datos queda TAN pegada al total que no hay
    espacio para evitar el traslape sin cortar la etiqueta del total,
    hay que priorizar no cortarla -mejor arrastrar un poco de la fila de
    al lado que dejar el total con la etiqueta rota (lo que pasó de
    verdad con un PDF real de CCSS)."""
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    # caso extremo: el total queda pegadísimo a la última fila, sin
    # espacio real para separarlos del todo
    pagina.insert_text((36, 150), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina.insert_text((36, 170), "TOTAL DE SALARIO 405710.71", fontsize=10)

    franja_total = _franja_total_en_pagina(pagina, formato="mnk")

    assert franja_total is not None
    rect_etiqueta_total = pagina.search_for("TOTAL DE TRABAJADORES")[0]
    assert franja_total.y0 <= rect_etiqueta_total.y0


def test_polizas_distintas_conservan_su_propio_encabezado(tmp_path):
    """Si el mismo cliente tiene oficiales en dos pólizas (archivos)
    distintas, cada una debe llegar en su propia hoja con su propio
    encabezado -sin mezclarse con el de la otra póliza."""
    ruta_poliza_a = tmp_path / "poliza_a.pdf"
    ruta_poliza_b = tmp_path / "poliza_b.pdf"
    _crear_pdf_planilla(ruta_poliza_a, "EMPRESA POLIZA A", [[("111111111", "JUAN", "PEREZ MORA", "Ninguna")]])
    _crear_pdf_planilla(ruta_poliza_b, "EMPRESA POLIZA B", [[("111111111", "JUAN", "PEREZ MORA", "Ninguna")]])

    registros = [{"cedula": "111111111", "cliente": "Cliente Multi Poliza", "nombre": "Juan Perez"}]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza_a), str(ruta_poliza_b)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento = fitz.open(resultado["archivos_por_cliente"]["Cliente Multi Poliza"])
    try:
        assert documento.page_count == 2
        texto_pagina_1 = documento[0].get_text()
        texto_pagina_2 = documento[1].get_text()
        assert "EMPRESA POLIZA A" in texto_pagina_1 and "EMPRESA POLIZA B" not in texto_pagina_1
        assert "EMPRESA POLIZA B" in texto_pagina_2 and "EMPRESA POLIZA A" not in texto_pagina_2
    finally:
        documento.close()


def test_encabezado_se_repite_si_la_poliza_desborda_una_hoja(tmp_path):
    """Si una misma póliza trae tantos oficiales que no caben en una sola
    hoja de salida, la segunda hoja también debe traer el encabezado
    arriba (si no, se pierde de vista qué póliza es)."""
    cantidad_filas = 60
    filas = [
        (f"{100000000 + i}", "NOMBRE", f"APELLIDO {i}", "Ninguna")
        for i in range(cantidad_filas)
    ]
    ruta_poliza = tmp_path / "poliza_larga.pdf"
    _crear_pdf_planilla(ruta_poliza, "EMPRESA POLIZA LARGA", [filas], espacio_filas=10)

    registros = [
        {"cedula": f"{100000000 + i}", "cliente": "Cliente Lote Grande", "nombre": f"Nombre {i}"}
        for i in range(cantidad_filas)
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento = fitz.open(resultado["archivos_por_cliente"]["Cliente Lote Grande"])
    try:
        assert documento.page_count > 1
        texto_completo = ""
        for pagina_salida in documento:
            assert "EMPRESA POLIZA LARGA" in pagina_salida.get_text()
            texto_completo += pagina_salida.get_text()
        # el desborde de hoja ahora también guarda a disco y libera memoria
        # a mitad de camino (ver _flush_a_disco en _agregar_bloque) -esto
        # confirma que ninguna fila se pierde ni se corrompe al reabrir el
        # documento para seguir agregando contenido
        for i in range(cantidad_filas):
            assert f"APELLIDO {i}" in texto_completo
    finally:
        documento.close()


def test_oficial_asignado_a_multiples_clientes_aparece_en_ambos_pdfs(ruta_pdf_ejemplo, tmp_path):
    """Si Juan Perez está asignado en el Excel tanto a 'Cliente Walmart' como
    a 'Cliente BAC', su fila debe exportarse al PDF de ambos clientes."""
    registros_multiples = [
        {"cedula": "111111111", "cliente": "Cliente Walmart", "nombre": "Juan Perez"},
        {"cedula": "111111111", "cliente": "Cliente BAC", "nombre": "Juan Perez"},
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_multiples, str(tmp_path / "salida"), formato="mnk"
    )

    # 1. Ambos clientes deben tener su propio archivo generado
    assert "Cliente Walmart" in resultado["archivos_por_cliente"]
    assert "Cliente BAC" in resultado["archivos_por_cliente"]

    # 2. Ambos PDFs deben contener la fila de Juan Perez
    doc_walmart = fitz.open(resultado["archivos_por_cliente"]["Cliente Walmart"])
    doc_bac = fitz.open(resultado["archivos_por_cliente"]["Cliente BAC"])
    try:
        assert "JUAN" in doc_walmart[0].get_text()
        assert "JUAN" in doc_bac[0].get_text()
    finally:
        doc_walmart.close()
        doc_bac.close()

    # 3. No debe quedar como cédula no encontrada
    assert resultado["no_encontrados"] == []


def test_exportar_por_cliente_sin_resaltado_no_agrega_anotaciones(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    """Cuando resaltar_filas=False, el PDF generado debe contener los textos
    recortados pero ninguna anotación de resaltado (highlight)."""
    carpeta_salida = tmp_path / "salida_limpia"

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo],
        registros_ejemplo,
        str(carpeta_salida),
        formato="mnk",
        resaltar_filas=False,
    )

    ruta_cliente = resultado["archivos_por_cliente"]["Cliente Prueba Uno"]
    documento = fitz.open(ruta_cliente)
    try:
        pagina = documento[0]
        # Verificar que el texto existe
        assert "JUAN" in pagina.get_text()
        # Verificar que NO tiene ninguna anotación de resaltado
        assert pagina.first_annot is None
    finally:
        documento.close()


def test_exportar_por_cliente_con_resaltado_agrega_anotacion(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    """Por defecto (resaltar_filas=True), el PDF debe incluir las anotaciones."""
    carpeta_salida = tmp_path / "salida_resaltada"

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo],
        registros_ejemplo,
        str(carpeta_salida),
        formato="mnk",
        resaltar_filas=True,
    )

    ruta_cliente = resultado["archivos_por_cliente"]["Cliente Prueba Uno"]
    documento = fitz.open(ruta_cliente)
    try:
        pagina = documento[0]
        assert pagina.first_annot is not None
        assert pagina.first_annot.type[0] == fitz.PDF_ANNOT_HIGHLIGHT
    finally:
        documento.close()


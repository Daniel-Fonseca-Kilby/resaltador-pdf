"""Pruebas de resaltar_por_cedula_y_exportar_por_cliente (datos inventados, ver conftest.py)."""
from pathlib import Path

import pymupdf as fitz

from resaltado_pdf import _techo_de_datos, resaltar_por_cedula_y_exportar_por_cliente


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


def test_segunda_pasada_rescata_por_nombre_si_la_cedula_no_calza(tmp_path):
    """Si el Excel trae un identificador que no aparece en el PDF (ej. un
    DIMEX en vez del número que imprime la CCSS para esa persona), pero el
    nombre completo sí aparece en una sola fila de todo el lote, se rescata
    por nombre en la segunda pasada."""
    ruta_poliza = tmp_path / "poliza.pdf"
    _crear_pdf_planilla(
        ruta_poliza,
        "EMPRESA AUTOSTAR",
        filas_por_pagina=[[("716322536", "ROBERTO", "REYES RODRIGUEZ", "Ninguna")]],
    )
    registros = [
        {"cedula": "155802367909", "cliente": "AUTOSTAR", "nombre": "Roberto Reyes Rodriguez"},
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    assert resultado["no_encontrados"] == []
    assert "AUTOSTAR" in resultado["archivos_por_cliente"]

    detalle = resultado["detalle_registros"][0]
    assert detalle["encontrado"] is True
    assert detalle["encontrado_por"] == "nombre"

    documento = fitz.open(resultado["archivos_por_cliente"]["AUTOSTAR"])
    try:
        assert "REYES" in documento[0].get_text()
    finally:
        documento.close()


def test_segunda_pasada_no_rescata_si_el_nombre_es_ambiguo(tmp_path):
    """Si el nombre completo aparece en más de una fila del lote, no se
    arriesga a adivinar cuál es la persona correcta -mejor dejarla como
    no encontrada que asignarla mal."""
    ruta_poliza = tmp_path / "poliza.pdf"
    _crear_pdf_planilla(
        ruta_poliza,
        "EMPRESA AMBIGUA",
        filas_por_pagina=[[
            ("111111111", "JUAN", "PEREZ MORA", "Ninguna"),
            ("222222222", "JUAN", "PEREZ MORA", "Ninguna"),
        ]],
    )
    registros = [
        {"cedula": "999999999", "cliente": "Cliente Ambiguo", "nombre": "Juan Perez Mora"},
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    assert len(resultado["no_encontrados"]) == 1
    assert resultado["detalle_registros"][0]["encontrado"] is False
    assert resultado["archivos_por_cliente"] == {}


def test_segunda_pasada_no_le_roba_la_fila_a_otro_empleado_conocido(tmp_path):
    """Si la fila que calza por nombre ya tiene una cédula que pertenece a
    OTRO empleado ya registrado en el Excel, no se la asigna -evita
    confundir a dos personas que solo comparten nombre."""
    ruta_poliza = tmp_path / "poliza.pdf"
    _crear_pdf_planilla(
        ruta_poliza,
        "EMPRESA DUPLICADA",
        filas_por_pagina=[[("111111111", "JUAN", "PEREZ MORA", "Ninguna")]],
    )
    registros = [
        {"cedula": "111111111", "cliente": "Cliente Correcto", "nombre": "Juan Perez Mora"},
        {"cedula": "999999999", "cliente": "Cliente Equivocado", "nombre": "Juan Perez Mora"},
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    assert "Cliente Correcto" in resultado["archivos_por_cliente"]
    assert "Cliente Equivocado" not in resultado["archivos_por_cliente"]

    no_encontrados_clientes = {r["cliente"] for r in resultado["no_encontrados"]}
    assert no_encontrados_clientes == {"Cliente Equivocado"}


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
    _PERFILES_PIE_PAGINA) se agrega al final del PDF de cada cliente."""
    ruta_poliza = tmp_path / "poliza_con_total.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA CON TOTAL", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", "JUAN", "PEREZ MORA", "Ninguna")])
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
        texto_completo = "".join(p.get_text() for p in documento_salida)
        assert "TOTAL DE TRABAJADORES" in texto_completo
        assert "TOTAL DE SALARIO" in texto_completo
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
    _escribir_fila(pagina, 140, [("111111111", "JUAN", "PEREZ MORA", "Ninguna")])
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
        for pagina_salida in documento:
            assert "EMPRESA POLIZA LARGA" in pagina_salida.get_text()
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


from pathlib import Path

import pymupdf as fitz

from resaltado_pdf import (
    _MARGEN_ABAJO_TOTAL,
    _extender_leyenda_para_incluir_total,
    _franja_leyenda_en_pagina,
    _franja_total_en_pagina,
    _franjas_pie_de_pagina,
    _recortar_total_antes_de_leyenda,
    _techo_de_datos,
    resaltar_por_cedula_sin_recortar,
    resaltar_por_cedula_y_exportar_por_cliente,
)


def _escribir_fila(pagina, y, celdas, fontsize=10):
    x = 36
    for texto, ancho in celdas:
        pagina.insert_text((x, y), texto, fontsize=fontsize)
        x += ancho


_TITULOS_COLUMNAS = [
    ("IDENTIFICACION", 90), ("NOMBRE", 80), ("APELLIDOS", 100), ("OBSERVACION", 90),
]


def _crear_pdf_planilla(ruta, empresa, filas_por_pagina, espacio_filas=30):

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
    y_titulos = 100
    y_primera_fila = 140

    documento = fitz.open(ruta_pdf_ejemplo)
    pagina = documento[0]

    techo = _techo_de_datos(pagina, formato="mnk")

    assert techo is not None
    assert y_titulos < techo < y_primera_fila
    documento.close()


def test_encabezado_una_sola_vez_si_la_poliza_cabe_en_una_hoja(tmp_path):
    
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
  
    documento = fitz.open()

    pagina1 = documento.new_page(width=595, height=842)
    pagina1.insert_text((36, 40), "REPORTE MNK - LETTERHEAD COMPLETO", fontsize=13)
    pagina1.insert_text((36, 60), "Fecha: 01/01/2026", fontsize=10)
    _escribir_fila(pagina1, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina1, 140, [("999999999", 90), ("OTRO", 80), ("CLIENTE VIEJO", 100), ("Ninguna", 90)])

    pagina2 = documento.new_page(width=595, height=842)
    _escribir_fila(pagina2, 40, _TITULOS_COLUMNAS)
    _escribir_fila(pagina2, 80, [("111111111", 90), ("JUAN", 80), ("PEREZ NUEVO", 100), ("Ninguna", 90)])

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
    ruta_poliza = tmp_path / "poliza_hueco_grande.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA CON HUECO GRANDE", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])

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

                alto_maximo_imagen_pt = max(alto_maximo_imagen_pt, pixmap_img.height / 2)

        assert alto_maximo_imagen_pt < 100
    finally:
        documento_salida.close()


def test_franja_del_total_no_incluye_encabezado_repetido_pegado_arriba():

    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])

    _escribir_fila(pagina, 150, _TITULOS_COLUMNAS)
    pagina.insert_text((36, 168), "TOTAL SALARIOS", fontsize=10)

    franjas = _franjas_pie_de_pagina(pagina, formato="ccss")

    assert len(franjas) == 1
    rect_encabezado_repetido = pagina.search_for("OBSERVACION")[-1]
    assert franjas[0].y0 >= rect_encabezado_repetido.y1


def test_pie_de_pagina_cuando_el_total_se_repite_en_hoja_de_aviso_legal(tmp_path):
    
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
        
        assert len(imagenes) == 2

        alturas = [fitz.Pixmap(documento_salida.extract_image(xref)["image"]).height for xref in imagenes]
        assert alturas[0] > alturas[1], (
            "la leyenda (más alta, de la página 1) debe ir antes que el total repetido "
            "(más bajo, de la página 2) -no al revés"
        )
    finally:
        documento_salida.close()


def test_extender_leyenda_para_incluir_total_agrega_el_total_completo_arriba():

    franja_total = fitz.Rect(0, 200, 595, 248)
    franja_leyenda = fitz.Rect(0, 215, 595, 800)

    extendida = _extender_leyenda_para_incluir_total(franja_leyenda, franja_total, misma_pagina=True)

    assert extendida.y0 == franja_total.y0
    assert extendida.y1 == franja_leyenda.y1


def test_extender_leyenda_para_incluir_total_no_toca_franja_que_ya_lo_incluye():

    franja_total = fitz.Rect(0, 200, 595, 248)
    franja_leyenda_ya_incluye_el_total = fitz.Rect(0, 190, 595, 800)

    extendida = _extender_leyenda_para_incluir_total(
        franja_leyenda_ya_incluye_el_total, franja_total, misma_pagina=True
    )
    assert extendida == franja_leyenda_ya_incluye_el_total

    franja_leyenda_otra_pagina = fitz.Rect(0, 215, 595, 800)
    extendida_paginas_distintas = _extender_leyenda_para_incluir_total(
        franja_leyenda_otra_pagina, franja_total, misma_pagina=False
    )
    assert extendida_paginas_distintas == franja_leyenda_otra_pagina


def test_recortar_total_antes_de_leyenda_evita_tragarse_la_codificacion():

    franja_total_con_margen_generoso = fitz.Rect(0, 190, 842, 239)
    franja_leyenda = fitz.Rect(0, 220, 842, 595)

    recortada = _recortar_total_antes_de_leyenda(franja_total_con_margen_generoso, franja_leyenda, misma_pagina=True)

    assert recortada.y1 == franja_leyenda.y0
    assert recortada.y0 == franja_total_con_margen_generoso.y0  # nunca toca su propio inicio


def test_recortar_total_antes_de_leyenda_nunca_corta_por_debajo_de_su_contenido_real(tmp_path):

    franja_total = fitz.Rect(0, 190, 842, 190 + _MARGEN_ABAJO_TOTAL + 32)  # y1 = 248
    franja_leyenda = fitz.Rect(0, 210, 842, 595)

    recortada = _recortar_total_antes_de_leyenda(franja_total, franja_leyenda, misma_pagina=True)

    y1_real_del_total = franja_total.y1 - _MARGEN_ABAJO_TOTAL  # 222
    assert recortada.y1 == y1_real_del_total
    assert recortada.y1 > franja_leyenda.y0  # se acepta el traslape antes que cortar contenido


def test_recortar_total_antes_de_leyenda_no_toca_franja_que_no_se_superpone():

    franja_total = fitz.Rect(0, 190, 842, 210)
    franja_leyenda = fitz.Rect(0, 220, 842, 595)

    recortada = _recortar_total_antes_de_leyenda(franja_total, franja_leyenda, misma_pagina=True)
    assert recortada == franja_total


    franja_total_otra_pagina = fitz.Rect(0, 190, 842, 239)
    recortada_paginas_distintas = _recortar_total_antes_de_leyenda(
        franja_total_otra_pagina, franja_leyenda, misma_pagina=False
    )
    assert recortada_paginas_distintas == franja_total_otra_pagina


def test_recortar_total_antes_de_leyenda_nunca_corta_su_propia_etiqueta():

    franja_total = fitz.Rect(0, 190, 842, 239)
    franja_leyenda_antes_del_total = fitz.Rect(0, 185, 842, 595)

    recortada = _recortar_total_antes_de_leyenda(franja_total, franja_leyenda_antes_del_total, misma_pagina=True)
    assert recortada == franja_total


def test_pie_de_pagina_duplica_el_total_cuando_esta_pegado_a_codificacion(tmp_path):

    ruta_poliza = tmp_path / "poliza_total_pegado.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina.insert_text((36, 200), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina.insert_text((36, 220), "TOTAL DE SALARIO 405710.71", fontsize=10)
    pagina.insert_text((36, 232), "CODIFICACIÓN", fontsize=10)
    documento.save(str(ruta_poliza))
    documento.close()

    documento_ref = fitz.open(ruta_poliza)
    pagina_ref = documento_ref[0]
    franja_total_ref = _franja_total_en_pagina(pagina_ref, "mnk")
    franja_leyenda_ref = _franja_leyenda_en_pagina(pagina_ref, "mnk")
    alto_hoja = pagina_ref.rect.height
    documento_ref.close()
    assert franja_total_ref is not None and franja_leyenda_ref is not None
    assert franja_leyenda_ref.y0 < franja_total_ref.y1

    registros = [{"cedula": "111111111", "cliente": "Cliente Total Pegado", "nombre": "Juan Perez"}]
    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [str(ruta_poliza)], registros, str(tmp_path / "salida"), formato="mnk"
    )

    documento_salida = fitz.open(resultado["archivos_por_cliente"]["Cliente Total Pegado"])
    try:
        imagenes = [xref for p in documento_salida for xref, *_r in p.get_images(full=True)]
        assert len(imagenes) == 2

        alto_total, alto_leyenda = (
            fitz.Pixmap(documento_salida.extract_image(xref)["image"]).height / 2 for xref in imagenes
        )

        alto_minimo_total = (franja_total_ref.y1 - _MARGEN_ABAJO_TOTAL) - franja_total_ref.y0
        assert alto_total >= alto_minimo_total - 1

        alto_esperado_leyenda = alto_hoja - franja_total_ref.y0
        assert alto_leyenda > alto_esperado_leyenda - 5
    finally:
        documento_salida.close()


def test_techo_de_datos_no_confunde_numero_patronal_con_primera_fila():

    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 30), "PLANILLA MENSUAL", fontsize=13)
    pagina.insert_text((36, 60), "2-03101682626-001-001", fontsize=10)
    _escribir_fila(pagina, 100, [("APELLIDOS Y NOMBRES", 200), ("OBSERVACIONES", 100)])

    techo = _techo_de_datos(pagina, formato="ccss")

    assert techo is not None

    assert techo > 90


def test_encabezado_no_arrastra_la_primera_fila_de_datos_si_esta_muy_pegada(tmp_path):

    ruta_poliza = tmp_path / "poliza_header_apretado.pdf"
    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA HEADER APRETADO", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)

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

    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])

    pagina.insert_text((36, 156), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina.insert_text((36, 176), "TOTAL DE SALARIO 405710.71", fontsize=10)

    franja_total = _franja_total_en_pagina(pagina, formato="mnk")

    assert franja_total is not None
    rect_ultima_fila = pagina.search_for("PEREZ MORA")[0]
    assert franja_total.y0 >= rect_ultima_fila.y1


def test_franja_del_total_nunca_corta_su_propia_etiqueta(tmp_path):

    documento = fitz.open()
    pagina = documento.new_page(width=595, height=842)
    pagina.insert_text((36, 40), "EMPRESA PRUEBA", fontsize=13)
    _escribir_fila(pagina, 100, _TITULOS_COLUMNAS)
    _escribir_fila(pagina, 140, [("111111111", 90), ("JUAN", 80), ("PEREZ MORA", 100), ("Ninguna", 90)])
    pagina.insert_text((36, 150), "TOTAL DE TRABAJADORES 1", fontsize=10)
    pagina.insert_text((36, 170), "TOTAL DE SALARIO 405710.71", fontsize=10)
    franja_total = _franja_total_en_pagina(pagina, formato="mnk")
    assert franja_total is not None
    rect_etiqueta_total = pagina.search_for("TOTAL DE TRABAJADORES")[0]
    assert franja_total.y0 <= rect_etiqueta_total.y0


def test_polizas_distintas_conservan_su_propio_encabezado(tmp_path):

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

        for i in range(cantidad_filas):
            assert f"APELLIDO {i}" in texto_completo
    finally:
        documento.close()


def test_oficial_asignado_a_multiples_clientes_aparece_en_ambos_pdfs(ruta_pdf_ejemplo, tmp_path):
    registros_multiples = [
        {"cedula": "111111111", "cliente": "Cliente Walmart", "nombre": "Juan Perez"},
        {"cedula": "111111111", "cliente": "Cliente BAC", "nombre": "Juan Perez"},
    ]

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_multiples, str(tmp_path / "salida"), formato="mnk"
    )


    assert "Cliente Walmart" in resultado["archivos_por_cliente"]
    assert "Cliente BAC" in resultado["archivos_por_cliente"]


    doc_walmart = fitz.open(resultado["archivos_por_cliente"]["Cliente Walmart"])
    doc_bac = fitz.open(resultado["archivos_por_cliente"]["Cliente BAC"])
    try:
        assert "JUAN" in doc_walmart[0].get_text()
        assert "JUAN" in doc_bac[0].get_text()
    finally:
        doc_walmart.close()
        doc_bac.close()


    assert resultado["no_encontrados"] == []


def test_exportar_por_cliente_sin_resaltado_no_agrega_anotaciones(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
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

        assert "JUAN" in pagina.get_text()

        assert pagina.first_annot is None
    finally:
        documento.close()


def test_exportar_por_cliente_con_resaltado_agrega_anotacion(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):

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





def test_solo_resaltar_mantiene_el_pdf_completo_e_intacto(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):

    resultado = resaltar_por_cedula_sin_recortar(
        [ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "salida")
    )

    nombre_archivo = Path(ruta_pdf_ejemplo).name
    assert nombre_archivo in resultado["archivos_resaltados"]

    documento_original = fitz.open(ruta_pdf_ejemplo)
    documento_salida = fitz.open(resultado["archivos_resaltados"][nombre_archivo])
    try:
        assert documento_salida.page_count == documento_original.page_count
        for pagina_original, pagina_salida in zip(documento_original, documento_salida):
            assert pagina_salida.get_text() == pagina_original.get_text()
    finally:
        documento_original.close()
        documento_salida.close()


def test_solo_resaltar_agrega_una_anotacion_por_fila_encontrada(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    resultado = resaltar_por_cedula_sin_recortar(
        [ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "salida")
    )

    nombre_archivo = Path(ruta_pdf_ejemplo).name
    documento_salida = fitz.open(resultado["archivos_resaltados"][nombre_archivo])
    try:
        anotaciones = list(documento_salida[0].annots())
        assert len(anotaciones) == 3
        assert all(a.type[0] == fitz.PDF_ANNOT_HIGHLIGHT for a in anotaciones)
    finally:
        documento_salida.close()


def test_solo_resaltar_entrega_un_pdf_por_archivo_no_por_cliente(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):

    resultado = resaltar_por_cedula_sin_recortar(
        [ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "salida")
    )

    assert len(resultado["archivos_resaltados"]) == 1


def test_solo_resaltar_reporta_cedulas_no_encontradas(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    resultado = resaltar_por_cedula_sin_recortar(
        [ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "salida")
    )

    cedulas_no_encontradas = {r["cedula"] for r in resultado["no_encontrados"]}
    assert cedulas_no_encontradas == {"999999999"}


def test_solo_resaltar_encuentra_extranjero_por_numero_de_asegurado(tmp_path):

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

    resultado = resaltar_por_cedula_sin_recortar(
        [str(ruta_poliza)], registros, str(tmp_path / "salida")
    )

    assert resultado["no_encontrados"] == []
    detalle = resultado["detalle_registros"][0]
    assert detalle["encontrado"] is True
    assert detalle["encontrado_por"] == "numero_asegurado"
    assert detalle["cedula"] == "155812345678"

    nombre_archivo = Path(ruta_poliza).name
    documento = fitz.open(resultado["archivos_resaltados"][nombre_archivo])
    try:
        assert "GOMEZ" in documento[0].get_text()
        assert len(list(documento[0].annots())) == 1
    finally:
        documento.close()


def test_solo_resaltar_varios_archivos_da_un_pdf_por_cada_uno(tmp_path):
    ruta_poliza_a = tmp_path / "poliza_a.pdf"
    ruta_poliza_b = tmp_path / "poliza_b.pdf"
    _crear_pdf_planilla(ruta_poliza_a, "EMPRESA POLIZA A", [[("111111111", "JUAN", "PEREZ MORA", "Ninguna")]])
    _crear_pdf_planilla(ruta_poliza_b, "EMPRESA POLIZA B", [[("222222222", "MARIA", "SOLANO MORA", "Ninguna")]])

    registros = [
        {"cedula": "111111111", "cliente": "Cliente Multi Archivo", "nombre": "Juan Perez"},
        {"cedula": "222222222", "cliente": "Cliente Multi Archivo", "nombre": "Maria Solano"},
    ]

    resultado = resaltar_por_cedula_sin_recortar(
        [str(ruta_poliza_a), str(ruta_poliza_b)], registros, str(tmp_path / "salida")
    )

    assert set(resultado["archivos_resaltados"]) == {"poliza_a.pdf", "poliza_b.pdf"}
    assert resultado["no_encontrados"] == []


def test_solo_resaltar_pdf_con_contrasena_se_reporta_como_error_sin_tumbar_el_proceso(tmp_path):
    documento = fitz.open()
    documento.new_page()
    ruta = tmp_path / "protegido.pdf"
    documento.save(str(ruta), encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="clave123")
    documento.close()

    resultado = resaltar_por_cedula_sin_recortar(
        [str(ruta)],
        [{"cedula": "111111111", "cliente": "Cliente X"}],
        str(tmp_path / "salida"),
    )

    assert "protegido.pdf" in resultado["errores_por_archivo"]
    assert resultado["archivos_resaltados"] == {}



def test_por_cliente_reporta_tiempos_por_fase(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "salida"), formato="mnk"
    )
    tiempos = resultado["tiempos"]
    assert tiempos["conteos"]["paginas"] == 1
    assert tiempos["conteos"]["filas_encontradas"] > 0
    for fase in ("abrir_pdf", "extraer_texto", "buscar_cedulas", "copiar_filas", "guardar_pdf_cliente"):
        assert fase in tiempos["segundos"]


def test_sin_recortar_reporta_tiempos_por_fase(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    resultado = resaltar_por_cedula_sin_recortar(
        [ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "salida")
    )
    assert resultado["tiempos"]["conteos"]["paginas"] == 1
    assert "guardar_disco" in resultado["tiempos"]["segundos"]


def test_pdf_de_cliente_queda_compactado_y_legible(ruta_pdf_ejemplo, registros_ejemplo, tmp_path):
    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        [ruta_pdf_ejemplo], registros_ejemplo, str(tmp_path / "salida"), formato="mnk"
    )
    assert "compactar_pdf_cliente" in resultado["tiempos"]["segundos"]
    for ruta in resultado["archivos_por_cliente"].values():
        assert not list(Path(ruta).parent.glob("*.compacto.pdf"))
        documento = fitz.open(ruta)
        assert documento.page_count >= 1
        assert documento.xref_length() > 1
        documento.close()


def test_compactar_pdf_mantiene_logo_y_pie_en_cada_pagina(tmp_path):
    """Imita el flujo real: el PDF de cliente se guarda y se reabre en cada
    cambio de póliza, y cada vez se vuelve a insertar el mismo logo -eso
    deja copias duplicadas. Compactar debe dejar una sola copia guardada
    pero la imagen tiene que seguir viéndose en TODAS las páginas, igual
    que antes (píxel por píxel)."""
    from resaltado_pdf import _compactar_pdf

    logo = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 60, 60), False)
    logo.set_rect(logo.irect, (200, 30, 30))
    png_logo = logo.tobytes("png")

    ruta = tmp_path / "cliente.pdf"
    for numero_pagina in range(3):
        documento = fitz.open(str(ruta)) if ruta.exists() else fitz.open()
        pagina = documento.new_page(width=595, height=842)
        pagina.insert_image(fitz.Rect(36, 20, 96, 80), stream=png_logo)  # encabezado
        pagina.insert_text((36, 800), f"TOTAL DE TRABAJADORES pagina {numero_pagina}", fontsize=9)  # pie
        if ruta.exists():
            documento.save(str(ruta), incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        else:
            documento.save(str(ruta))
        documento.close()

    def _estado(ruta_pdf):
        documento = fitz.open(str(ruta_pdf))
        imagenes_por_pagina = [len(p.get_images()) for p in documento]
        xrefs_distintos = {img[0] for p in documento for img in p.get_images()}
        render = [p.get_pixmap(matrix=fitz.Matrix(1, 1)).samples for p in documento]
        documento.close()
        return imagenes_por_pagina, xrefs_distintos, render

    antes_por_pagina, antes_xrefs, antes_render = _estado(ruta)
    assert len(antes_xrefs) == 3  # una copia del logo por cada guardado

    _compactar_pdf(ruta)

    despues_por_pagina, despues_xrefs, despues_render = _estado(ruta)
    assert despues_por_pagina == antes_por_pagina == [1, 1, 1]  # el logo sigue en cada página
    assert len(despues_xrefs) == 1  # pero guardado una sola vez
    assert despues_render == antes_render  # y se ve idéntico


def test_compactar_pdf_une_logos_identicos_con_distinto_name(tmp_path):
    """Caso real de MNK: la planilla trae el mismo logo 3 veces, idéntico
    salvo por /Name (im1, im2, im3). Compactar debe dejar una sola copia y
    que cada página se siga viendo igual."""
    from resaltado_pdf import _compactar_pdf

    logo = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 60, 60), False)
    logo.set_rect(logo.irect, (20, 60, 180))
    png_logo = logo.tobytes("png")

    ruta = tmp_path / "planilla_mnk.pdf"
    for _ in range(3):
        documento = fitz.open(str(ruta)) if ruta.exists() else fitz.open()
        pagina = documento.new_page(width=595, height=842)
        pagina.insert_image(fitz.Rect(36, 20, 96, 80), stream=png_logo)
        if ruta.exists():
            documento.save(str(ruta), incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        else:
            documento.save(str(ruta))
        documento.close()

    documento = fitz.open(str(ruta))
    xrefs = sorted({img[0] for p in documento for img in p.get_images()})
    for numero, xref in enumerate(xrefs, start=1):
        documento.xref_set_key(xref, "Name", f"/im{numero}")
    documento.save(str(ruta), incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
    antes_render = [p.get_pixmap().samples for p in documento]
    documento.close()
    assert len(xrefs) == 3

    _compactar_pdf(ruta)

    documento = fitz.open(str(ruta))
    assert [len(p.get_images()) for p in documento] == [1, 1, 1]
    assert len({img[0] for p in documento for img in p.get_images()}) == 1
    assert [p.get_pixmap().samples for p in documento] == antes_render
    documento.close()

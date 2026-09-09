"""
resaltado_pdf.py

Busca nombres o cédulas dentro de PDFs de planillas y les agrega
resaltado, sin tocar el resto del documento (mismo formato, mismas
fuentes, etc.).
"""

import io
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import pymupdf as fitz  # alias tradicional de PyMuPDF
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


@dataclass
class ResultadoResaltado:
    """Resultado de procesar un PDF en modo simple."""
    archivo: str
    coincidencias_por_nombre: dict  # {"JUAN PEREZ": 3, "MARIA LOPEZ": 0}
    ruta_salida: str


def _normalizar(texto: str) -> str:
    """Mayúsculas y sin tildes, para comparar sin depender del acento."""
    texto = texto.strip().upper()
    texto_sin_tildes = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto_sin_tildes if not unicodedata.combining(c))


def _variantes_ene(texto: str) -> list[str]:
    """En algunos PDFs de MNK la Ñ sale como un espacio en blanco."""
    if "Ñ" not in texto and "ñ" not in texto:
        return []
    return [
        texto.replace("Ñ", " ").replace("ñ", " "),
        texto.replace("Ñ", "").replace("ñ", ""),
    ]


def _buscar_con_variantes(pagina, texto: str, textpage=None):
    """Prueba el texto tal cual, normalizado y con variantes de Ñ.

    textpage se puede pasar ya extraído (page.get_textpage()) para no
    reextraerlo en cada búsqueda dentro de la misma página.
    """
    coincidencias = pagina.search_for(texto, quads=False, textpage=textpage)
    if coincidencias:
        return coincidencias

    coincidencias = pagina.search_for(_normalizar(texto), quads=False, textpage=textpage)
    if coincidencias:
        return coincidencias

    for variante in _variantes_ene(texto):
        coincidencias = pagina.search_for(variante, quads=False, textpage=textpage)
        if coincidencias:
            return coincidencias

    return []


def _todas_las_palabras_en_texto(texto_norm: str, palabras: list[str]) -> bool:
    """Chequeo rápido antes de meterse a buscar con PyMuPDF: si alguna
    palabra ni aparece en el texto de la página, ahí no hay nada que
    buscar."""
    if not palabras:
        return False
    return all(
        any(_normalizar(v) in texto_norm for v in (palabra, *_variantes_ene(palabra)))
        for palabra in palabras
    )


def _puede_aparecer_en_pagina(texto_pagina_norm: str, texto: str, palabras: list[str]) -> bool:
    """Igual idea que _todas_las_palabras_en_texto pero probando primero
    la frase completa. Con miles de nombres este filtro barato ahorra
    mandar a buscar con search_for los que de plano no están en la
    página."""
    candidatos_frase = [_normalizar(texto), *[_normalizar(v) for v in _variantes_ene(texto)]]
    if any(c in texto_pagina_norm for c in candidatos_frase):
        return True
    return _todas_las_palabras_en_texto(texto_pagina_norm, palabras)


_TOLERANCIA_FILA = 6  # variación en Y (puntos) tolerada para considerar la misma fila


def _buscar_por_fila(pagina, palabras: list[str], textpage=None):
    """Busca cada palabra por separado y agrupa las que caen en la misma
    fila -para cuando nombre y apellido quedan en columnas distintas."""
    rects_por_palabra = []
    for palabra in palabras:
        rects = _buscar_con_variantes(pagina, palabra, textpage=textpage)
        if not rects:
            return []
        rects_por_palabra.append(rects)

    indice_ancla = min(range(len(rects_por_palabra)), key=lambda i: len(rects_por_palabra[i]))
    filas_encontradas = []
    for rect_ancla in rects_por_palabra[indice_ancla]:
        fila = [rect_ancla]
        completa = True
        for i, rects in enumerate(rects_por_palabra):
            if i == indice_ancla:
                continue
            candidato = next((r for r in rects if abs(r.y0 - rect_ancla.y0) <= _TOLERANCIA_FILA), None)
            if candidato is None:
                completa = False
                break
            fila.append(candidato)
        if completa:
            filas_encontradas.append(fila)
    return filas_encontradas


def resaltar_nombres_en_pdf(ruta_pdf: str, nombres: list[str], ruta_salida: str) -> ResultadoResaltado:
    """Resalta cada nombre de la lista en el PDF y guarda la copia en ruta_salida."""
    documento = fitz.open(ruta_pdf)
    if documento.is_encrypted:
        documento.close()
        raise ValueError(
            "El PDF está protegido con contraseña. Quite la protección e inténtelo de nuevo."
        )

    conteo = {nombre: 0 for nombre in nombres}

    for pagina in documento:
        textpage = pagina.get_textpage()
        texto_pagina_norm = _normalizar(pagina.get_text(textpage=textpage))

        for nombre in nombres:
            nombre_limpio = nombre.strip()
            if not nombre_limpio:
                continue

            palabras = [p for p in nombre_limpio.split() if len(p) >= 3]
            if not _puede_aparecer_en_pagina(texto_pagina_norm, nombre_limpio, palabras):
                continue

            coincidencias = _buscar_con_variantes(pagina, nombre_limpio, textpage=textpage)

            if coincidencias:
                for rect in coincidencias:
                    anotacion = pagina.add_highlight_annot(rect)
                    anotacion.update()
                    conteo[nombre] += 1
                continue

            # no aparece junto: puede que nombre y apellido queden en columnas distintas
            for fila in _buscar_por_fila(pagina, palabras, textpage=textpage):
                for rect in fila:
                    anotacion = pagina.add_highlight_annot(rect)
                    anotacion.update()
                conteo[nombre] += 1

    documento.save(ruta_salida)
    documento.close()

    return ResultadoResaltado(
        archivo=Path(ruta_pdf).name,
        coincidencias_por_nombre=conteo,
        ruta_salida=ruta_salida,
    )


def _nombre_archivo_seguro(texto: str) -> str:
    limpio = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in texto).strip()
    return limpio or "SIN_CLIENTE"


def _normalizar_cedula(digitos: str) -> str:
    """Quita ceros a la izquierda (si el Excel guardó la cédula como
    número, se pierde el cero inicial)."""
    return digitos.lstrip("0") or "0"


def _coincide_cliente(digitos_palabra: str, mapa_cedulas: dict) -> list[str]:
    """Clientes a los que pertenece esta cédula -un oficial puede estar
    asignado a más de uno. Tolera el dígito de tipo de identificación que
    antepone la CCSS y diferencias de ceros iniciales. mapa_cedulas debe
    venir con las claves ya normalizadas."""
    clientes = mapa_cedulas.get(_normalizar_cedula(digitos_palabra))
    if not clientes and len(digitos_palabra) > 1:
        clientes = mapa_cedulas.get(_normalizar_cedula(digitos_palabra[1:]))
    return clientes or []


# frases que solo aparecen en la fila de títulos de cada formato conocido
_PERFILES_ENCABEZADO = {
    "ccss": ["APELLIDOS Y NOMBRES", "OBSERVACIONES"],
    "mnk": ["IDENTIFICACIÓN", "OBSERVACIÓN"],
    "ins": ["Salario Reportado", "Descripción Ocupación"],
}


def _techo_por_anclas(pagina, anclas: list[str], textpage=None) -> float | None:
    """Y justo debajo de la fila donde aparecen las anclas."""
    y1_maximo = None
    for ancla in anclas:
        coincidencias = pagina.search_for(ancla, quads=False, textpage=textpage)
        if not coincidencias:
            coincidencias = pagina.search_for(_normalizar(ancla), quads=False, textpage=textpage)
        for rect in coincidencias:
            if y1_maximo is None or rect.y1 > y1_maximo:
                y1_maximo = rect.y1
    if y1_maximo is None:
        return None
    return y1_maximo + 14


def _techo_de_datos(pagina, formato: str = "auto", textpage=None) -> float | None:
    """Y donde arranca la tabla de empleados. Prueba primero las anclas
    del formato indicado y luego las de los demás formatos conocidos (por
    si el desplegable quedó mal puesto); si nada calza, cae al detector
    de tablas de PyMuPDF. Si tampoco encuentra nada, mejor no devolver
    encabezado que arriesgarse a mostrar la fila de otra persona."""
    perfiles_a_probar = []
    anclas_formato = _PERFILES_ENCABEZADO.get(formato)
    if anclas_formato:
        perfiles_a_probar.append(anclas_formato)
    for anclas_perfil in _PERFILES_ENCABEZADO.values():
        if anclas_perfil not in perfiles_a_probar:
            perfiles_a_probar.append(anclas_perfil)

    for anclas in perfiles_a_probar:
        techo = _techo_por_anclas(pagina, anclas, textpage=textpage)
        if techo is not None:
            return techo

    for estrategia in ("lines_strict", "lines", "text"):
        try:
            tablas = pagina.find_tables(strategy=estrategia).tables
        except Exception:
            continue
        if not tablas:
            continue
        tabla_datos = max(tablas, key=lambda t: len(t.rows))
        if tabla_datos.rows:
            return fitz.Rect(tabla_datos.rows[0].bbox).y0
    return None


# frases del renglón del total, al final de cada póliza
_PERFILES_PIE_PAGINA = {
    "mnk": ["TOTAL DE TRABAJADORES", "TOTAL DE SALARIO"],
    "ccss": ["TOTAL SALARIOS"],
}

# frases donde arranca la leyenda/firma, después del total
_PERFILES_LEYENDA_PIE = {
    "mnk": ["CODIFICACIÓN"],
    "ccss": ["Ajuste al mínimo base diferenciada SEM"],
}

_MARGEN_ARRIBA_PIE = 12
_MARGEN_ABAJO_TOTAL = 26  # alto aprox. del renglón del total + su caja de valor


def _franja_por_anclas(pagina, anclas: list[str], textpage=None) -> tuple[float, float] | None:
    """(y0, y1) del renglón donde aparecen las anclas, sin margen."""
    y0_minimo = None
    y1_maximo = None
    for ancla in anclas:
        coincidencias = pagina.search_for(ancla, quads=False, textpage=textpage)
        if not coincidencias:
            coincidencias = pagina.search_for(_normalizar(ancla), quads=False, textpage=textpage)
        for rect in coincidencias:
            if y0_minimo is None or rect.y0 < y0_minimo:
                y0_minimo = rect.y0
            if y1_maximo is None or rect.y1 > y1_maximo:
                y1_maximo = rect.y1
    if y0_minimo is None:
        return None
    return (y0_minimo, y1_maximo)


def _franja_por_perfiles(pagina, perfiles: dict, formato: str, textpage=None) -> tuple[float, float] | None:
    """Mismo criterio que _techo_de_datos: prueba el formato indicado y
    después los demás perfiles conocidos."""
    perfiles_a_probar = []
    anclas_formato = perfiles.get(formato)
    if anclas_formato:
        perfiles_a_probar.append(anclas_formato)
    for anclas_perfil in perfiles.values():
        if anclas_perfil not in perfiles_a_probar:
            perfiles_a_probar.append(anclas_perfil)

    for anclas in perfiles_a_probar:
        franja = _franja_por_anclas(pagina, anclas, textpage=textpage)
        if franja is not None:
            return franja
    return None


def _franja_total_en_pagina(pagina, formato: str = "auto", textpage=None) -> "fitz.Rect | None":
    """Franja del renglón del total en esta página, o None si no aparece."""
    total = _franja_por_perfiles(pagina, _PERFILES_PIE_PAGINA, formato, textpage=textpage)
    if total is None:
        return None
    y0_total, y1_total = total

    # cuando a la hoja le quedan pocas filas, CCSS repite el renglón de
    # títulos de columna justo antes del total -si ese renglón cae más
    # cerca del total que el margen fijo, hay que recortar justo debajo
    # de él para no arrastrarlo al pie de página. Se usa un margen chico
    # (no el de _techo_de_datos, pensado para el espacio más amplio antes
    # de la primera fila de datos) para no pasarse de largo y comerse el
    # total.
    margen_superior = y0_total - _MARGEN_ARRIBA_PIE
    encabezado_repetido = _franja_por_perfiles(pagina, _PERFILES_ENCABEZADO, formato, textpage=textpage)
    if encabezado_repetido is not None:
        _y0_encabezado, y1_encabezado = encabezado_repetido
        limite_tras_encabezado = y1_encabezado + 4
        if margen_superior < limite_tras_encabezado < y0_total:
            margen_superior = limite_tras_encabezado

    return fitz.Rect(pagina.rect.x0, margen_superior, pagina.rect.x1, y1_total + _MARGEN_ABAJO_TOTAL)


def _franja_leyenda_en_pagina(pagina, formato: str = "auto", textpage=None) -> "fitz.Rect | None":
    """Franja de la leyenda/firma en esta página, o None si no aparece."""
    leyenda = _franja_por_perfiles(pagina, _PERFILES_LEYENDA_PIE, formato, textpage=textpage)
    if leyenda is None:
        return None
    y0_leyenda, _y1_leyenda = leyenda
    return fitz.Rect(pagina.rect.x0, y0_leyenda - _MARGEN_ARRIBA_PIE, pagina.rect.x1, pagina.rect.height)


def _recortar_leyenda_tras_total(
    franja_leyenda: "fitz.Rect", franja_total: "fitz.Rect | None", misma_pagina: bool,
) -> "fitz.Rect":
    """Si el total y la leyenda cayeron en la misma página y el margen de
    la leyenda se superpone con el total, la recorta para que empiece
    justo donde termina el total -en MNK la barra de "CODIFICACIÓN" viene
    pegada justo debajo de la del total, casi sin espacio, y sin este
    ajuste la leyenda vuelve a arrastrar el total que ya se capturó por
    separado."""
    if misma_pagina and franja_total is not None and franja_leyenda.y0 < franja_total.y1:
        return fitz.Rect(franja_leyenda.x0, franja_total.y1, franja_leyenda.x1, franja_leyenda.y1)
    return franja_leyenda


def _franjas_pie_de_pagina(pagina, formato: str = "auto", textpage=None) -> list["fitz.Rect"]:
    """Franjas del pie de página de ESTA hoja: el total y, si aparece, la
    leyenda con la firma. Cada una se recorta pegada a su propio
    contenido, sin el espacio en blanco que las separa en el documento
    original. Si el formato no tiene perfil de total conocido (ej. INS),
    no se agrega nada.

    Asume que el total y la leyenda están en la MISMA página -para
    buscarlos cada uno en la última página del documento donde
    realmente aparezcan (pueden caer en hojas distintas, ver
    _pixmaps_pie_de_poliza), usar _franja_total_en_pagina y
    _franja_leyenda_en_pagina por separado."""
    franjas = []

    franja_total = _franja_total_en_pagina(pagina, formato, textpage=textpage)
    if franja_total is None:
        return franjas
    franjas.append(franja_total)

    franja_leyenda = _franja_leyenda_en_pagina(pagina, formato, textpage=textpage)
    if franja_leyenda is not None:
        franjas.append(franja_leyenda)

    return franjas


_MARGEN_PAGINA = 24  # margen arriba/abajo de cada página de salida
_ESPACIO_ENTRE_FILAS = 3


def resaltar_por_cedula_y_exportar_por_cliente(
    rutas_pdfs: list[str],
    registros: list[dict],
    carpeta_salida: str,
    formato: str = "auto",
    resaltar_filas: bool = True,
) -> dict:
    """Busca las cédulas de 'registros' en los PDFs, recorta cada fila
    encontrada y arma un PDF por cliente.

    Una misma cédula puede estar asignada a varios clientes (un oficial
    que cubrió turnos en más de un puesto durante la quincena): su fila
    se agrega al PDF de cada uno.

    El encabezado se saca siempre de la página 1 del archivo, no de donde
    arranca cada cliente -en reportes de varias páginas (MNK, por
    ejemplo) solo la primera trae el logo/título completos. Se repite una
    vez por póliza y de nuevo si esa póliza desborda a una segunda hoja.

    A los registros que no calzan por cédula se les hace una segunda
    pasada buscándolos por nombre completo (rescata casos como un DIMEX
    en el Excel contra el número de CCSS que imprime la planilla). Solo
    se acepta si el nombre aparece en exactamente una fila de todo el
    lote y esa fila no es ya de otro empleado conocido.

    Al cerrar cada póliza se le agrega su propio pie de página (el total
    y, si aparece, la leyenda con la firma) tal cual sale en el original
    -así lo pidió VMA. Se copia como imagen, no de forma vectorial, porque
    en CCSS el total lo rellena la Oficina Virtual como campo de
    formulario y show_pdf_page no arrastra ese valor.
    """
    registros_unicos: dict[tuple[str, str], dict] = {}
    for r in registros:
        cedula, cliente = r.get("cedula"), r.get("cliente")
        if not cedula or not cliente:
            continue
        clave = (_normalizar_cedula(cedula), cliente)
        if clave not in registros_unicos:
            registros_unicos[clave] = {"cedula": cedula, "cliente": cliente, "nombre": r.get("nombre", "")}

    mapa_cedulas: dict[str, list[str]] = {}
    for clave_cedula, cliente in registros_unicos:
        clientes = mapa_cedulas.setdefault(clave_cedula, [])
        if cliente not in clientes:
            clientes.append(cliente)

    estado_por_cliente: dict[str, dict] = {}
    polizas_encontradas: dict[tuple[str, str], set] = {}
    archivos_usados_por_cliente: dict[str, set] = {}  # para que la segunda pasada no reabra una póliza ya cerrada
    errores_por_archivo: dict[str, str] = {}

    carpeta = Path(carpeta_salida)
    carpeta.mkdir(parents=True, exist_ok=True)

    def _obtener_estado(cliente: str, pagina_origen) -> dict:
        estado = estado_por_cliente.get(cliente)
        if estado is None:
            estado = {
                "documento": fitz.open(),
                "pagina": None,
                "y": 0.0,
                "ancho": pagina_origen.rect.width,
                "alto": pagina_origen.rect.height,
                "archivo_actual": None,
                "encabezado_actual": None,
                "ruta_salida": carpeta / f"{_nombre_archivo_seguro(cliente)}.pdf",
                "guardado_en_disco": False,
            }
            estado_por_cliente[cliente] = estado
        return estado

    def _asegurar_documento(estado: dict) -> None:
        """Reabre el PDF del cliente si ya se había guardado y cerrado
        (ver _flush_a_disco)."""
        if estado["documento"] is None:
            estado["documento"] = fitz.open(str(estado["ruta_salida"]))
            estado["pagina"] = None

    def _flush_a_disco(estado: dict) -> None:
        """Guarda lo acumulado del cliente hasta ahora y cierra el
        documento en memoria. Con lotes grandes, mantener el PDF completo
        de cada cliente en RAM hasta el final del proceso es lo que
        termina agotando la memoria del servidor -por eso esto se llama
        en cada cambio de póliza, una vez que el bloque anterior ya quedó
        cerrado y no se vuelve a tocar."""
        documento = estado["documento"]
        if documento is None:
            return
        ruta = str(estado["ruta_salida"])
        if estado["guardado_en_disco"]:
            documento.save(ruta, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
        else:
            documento.save(ruta)
            estado["guardado_en_disco"] = True
        documento.close()
        estado["documento"] = None
        estado["pagina"] = None

    def _agregar_bloque(
        estado: dict,
        documento_origen,
        pagina_origen,
        franja: "fitz.Rect",
        resaltar: bool,
        repetir_encabezado: bool = True,
    ) -> None:
        escala = estado["ancho"] / pagina_origen.rect.width
        alto_bloque = franja.height * escala
        if alto_bloque <= 0:
            return

        _asegurar_documento(estado)
        if estado["pagina"] is None or estado["y"] + alto_bloque > estado["alto"] - _MARGEN_PAGINA:
            # a diferencia del cambio de póliza, acá no se hace flush:
            # guardar y reabrir a mitad de una misma póliza obliga a
            # PyMuPDF a volver a copiar el logo/fuentes del encabezado en
            # cada corte, duplicándolos. Mejor dejar que la póliza se
            # termine de escribir de un tirón.
            estado["pagina"] = estado["documento"].new_page(width=estado["ancho"], height=estado["alto"])
            estado["y"] = _MARGEN_PAGINA
            encabezado = estado["encabezado_actual"]
            if repetir_encabezado and encabezado is not None:
                _agregar_bloque(
                    estado, encabezado["documento"], encabezado["pagina"], encabezado["franja"],
                    resaltar=False, repetir_encabezado=False,
                )

        destino = fitz.Rect(0, estado["y"], estado["ancho"], estado["y"] + alto_bloque)
        estado["pagina"].show_pdf_page(destino, documento_origen, pagina_origen.number, clip=franja)
        if resaltar:
            anotacion = estado["pagina"].add_highlight_annot(destino)
            anotacion.update()
        estado["y"] += alto_bloque + _ESPACIO_ENTRE_FILAS

    def _agregar_imagen_cacheada(estado: dict, png_bytes: bytes, ancho_original: float, alto_original: float) -> None:
        """Como _agregar_bloque pero para una imagen ya renderizada (ver
        _pixmaps_pie_de_poliza) -sirve para no reabrir el archivo ni
        volver a renderizar el pie de página por cada cliente que
        comparte la misma póliza."""
        if ancho_original <= 0:
            return
        escala = estado["ancho"] / ancho_original
        alto_bloque = alto_original * escala
        if alto_bloque <= 0:
            return

        _asegurar_documento(estado)
        if estado["pagina"] is None or estado["y"] + alto_bloque > estado["alto"] - _MARGEN_PAGINA:
            estado["pagina"] = estado["documento"].new_page(width=estado["ancho"], height=estado["alto"])
            estado["y"] = _MARGEN_PAGINA

        destino = fitz.Rect(0, estado["y"], estado["ancho"], estado["y"] + alto_bloque)
        estado["pagina"].insert_image(destino, stream=png_bytes)
        estado["y"] += alto_bloque + _ESPACIO_ENTRE_FILAS

    pixmaps_pie_por_archivo: dict[str, list[tuple]] = {}

    def _pixmaps_pie_de_poliza(ruta_pdf_saliente: str) -> list[tuple]:
        """Renderiza el pie de página de esta póliza una sola vez (varios
        clientes suelen compartirla) y lo guarda en caché como PNG. Se
        copia como imagen porque en CCSS el total lo rellena la Oficina
        Virtual como campo de formulario, y show_pdf_page no arrastra ese
        valor.

        El total y la leyenda se buscan cada uno por separado, retrocediendo
        desde la última página -MNK a veces repite el total al principio de
        una página de aviso legal/datos de contacto que viene DESPUÉS de la
        que trae "CODIFICACIÓN", así que no siempre caen en la misma hoja.
        Si ninguna de las últimas páginas tiene el total, no se agrega nada
        (mejor que arriesgarse a recortar cualquier cosa)."""
        if ruta_pdf_saliente in pixmaps_pie_por_archivo:
            return pixmaps_pie_por_archivo[ruta_pdf_saliente]

        resultado: list[tuple] = []
        try:
            documento_pie = fitz.open(ruta_pdf_saliente)
        except Exception:
            pixmaps_pie_por_archivo[ruta_pdf_saliente] = resultado
            return resultado
        try:
            if not documento_pie.is_encrypted:
                ultimo_indice = documento_pie.page_count - 1
                num_a_revisar = min(3, documento_pie.page_count)

                pagina_total = franja_total = None
                for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
                    candidata = documento_pie[indice]
                    franja = _franja_total_en_pagina(candidata, formato)
                    if franja is not None:
                        pagina_total, franja_total = candidata, franja
                        break

                if pagina_total is not None:
                    # si el PDF no trae ya generada la apariencia del campo,
                    # PyMuPDF lo captura en blanco -forzar que la regenere
                    for widget in pagina_total.widgets() or []:
                        try:
                            widget.update()
                        except Exception:
                            pass
                    pixmap = pagina_total.get_pixmap(clip=franja_total, matrix=fitz.Matrix(2, 2))
                    resultado.append((pixmap.tobytes("png"), pagina_total.rect.width, franja_total.height))

                pagina_leyenda = franja_leyenda = None
                for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
                    candidata = documento_pie[indice]
                    franja = _franja_leyenda_en_pagina(candidata, formato)
                    if franja is not None:
                        pagina_leyenda, franja_leyenda = candidata, franja
                        break

                if pagina_leyenda is not None:
                    franja_leyenda = _recortar_leyenda_tras_total(
                        franja_leyenda, franja_total, pagina_leyenda is pagina_total,
                    )
                    pixmap = pagina_leyenda.get_pixmap(clip=franja_leyenda, matrix=fitz.Matrix(2, 2))
                    resultado.append((pixmap.tobytes("png"), pagina_leyenda.rect.width, franja_leyenda.height))
        except Exception:
            pass
        finally:
            documento_pie.close()

        pixmaps_pie_por_archivo[ruta_pdf_saliente] = resultado
        return resultado

    def _agregar_pie_de_poliza(estado: dict, ruta_pdf_saliente: str) -> None:
        """Cierra la póliza que este cliente está dejando atrás con su
        propio total, antes de pasar a la siguiente."""
        for png_bytes, ancho_pagina, alto_franja in _pixmaps_pie_de_poliza(ruta_pdf_saliente):
            _agregar_imagen_cacheada(estado, png_bytes, ancho_pagina, alto_franja)

    for ruta_pdf in rutas_pdfs:
        nombre_archivo = Path(ruta_pdf).name
        try:
            documento = fitz.open(ruta_pdf)
        except Exception as error:
            errores_por_archivo[nombre_archivo] = f"No se pudo abrir el archivo: {error}"
            continue

        if documento.is_encrypted:
            documento.close()
            errores_por_archivo[nombre_archivo] = (
                "El PDF está protegido con contraseña. Quite la protección e inténtelo de nuevo."
            )
            continue

        encabezado_archivo = None
        encabezado_calculado = False

        try:
            for pagina in documento:
                textpage = pagina.get_textpage()
                palabras = pagina.get_text("words", textpage=textpage)

                franjas_vistas: dict[str, list] = {}
                for x0, y0, x1, y1, palabra, *_resto in palabras:
                    digitos = "".join(c for c in palabra if c.isdigit())
                    if not digitos:
                        continue
                    clientes = _coincide_cliente(digitos, mapa_cedulas)
                    if not clientes:
                        continue
                    clave_cedula = _normalizar_cedula(digitos)
                    if clave_cedula not in mapa_cedulas and len(digitos) > 1:
                        clave_cedula = _normalizar_cedula(digitos[1:])

                    fila_y0 = min(w[1] for w in palabras if abs(w[1] - y0) <= _TOLERANCIA_FILA)
                    fila_y1 = max(w[3] for w in palabras if abs(w[1] - y0) <= _TOLERANCIA_FILA)
                    franja = fitz.Rect(pagina.rect.x0, fila_y0 - 2, pagina.rect.x1, fila_y1 + 2)

                    for cliente in clientes:
                        polizas_encontradas.setdefault((clave_cedula, cliente), set()).add(nombre_archivo)

                        vistas = franjas_vistas.setdefault(cliente, [])
                        if franja in vistas:
                            continue
                        vistas.append(franja)

                        estado = _obtener_estado(cliente, pagina)
                        archivos_usados_por_cliente.setdefault(cliente, set()).add(ruta_pdf)

                        if estado["archivo_actual"] != ruta_pdf:
                            if estado["archivo_actual"] is not None:
                                _agregar_pie_de_poliza(estado, estado["archivo_actual"])
                                _flush_a_disco(estado)
                            if not encabezado_calculado:
                                primera_pagina = documento[0]
                                techo_pagina1 = _techo_de_datos(primera_pagina, formato)
                                if techo_pagina1 is not None and techo_pagina1 > 4:
                                    encabezado_archivo = {
                                        "documento": documento,
                                        "pagina": primera_pagina,
                                        "franja": fitz.Rect(
                                            primera_pagina.rect.x0, 0, primera_pagina.rect.x1, techo_pagina1 - 2
                                        ),
                                    }
                                encabezado_calculado = True
                            estado["pagina"] = None
                            estado["archivo_actual"] = ruta_pdf
                            estado["encabezado_actual"] = encabezado_archivo
                            if encabezado_archivo is not None:
                                _agregar_bloque(
                                    estado, encabezado_archivo["documento"], encabezado_archivo["pagina"],
                                    encabezado_archivo["franja"], resaltar=False, repetir_encabezado=False,
                                )

                        _agregar_bloque(estado, documento, pagina, franja, resaltar=resaltar_filas)
        except Exception as error:
            errores_por_archivo[nombre_archivo] = f"No se pudo procesar el archivo: {error}"
        finally:
            documento.close()

    # segunda pasada: a los que no calzaron por cédula se les busca el
    # nombre completo. Solo se acepta si aparece en exactamente una fila
    # de todo el lote y esa fila no tiene ya la cédula de otro empleado
    encontrados_por_nombre: set = set()
    # clave -> por qué la segunda pasada no lo rescató (nombre ambiguo, y
    # dónde); si no aparece acá y tampoco se encontró, es que ni siquiera
    # apareció el nombre en ninguna página
    motivos_no_rescatado: dict[tuple[str, str], str] = {}
    pendientes_por_nombre: dict[tuple[str, str], list[str]] = {}
    for clave, datos in registros_unicos.items():
        if polizas_encontradas.get(clave):
            continue
        palabras_nombre = [p for p in _normalizar(datos.get("nombre") or "").split() if len(p) >= 3]
        if len(palabras_nombre) >= 2:
            pendientes_por_nombre[clave] = palabras_nombre

    if pendientes_por_nombre:
        candidatos_por_pendiente: dict[tuple[str, str], list[tuple[str, int, tuple]]] = {}
        for ruta_pdf in rutas_pdfs:
            try:
                documento = fitz.open(ruta_pdf)
            except Exception:
                continue
            if documento.is_encrypted:
                documento.close()
                continue
            try:
                for pagina in documento:
                    textpage = pagina.get_textpage()
                    palabras_pagina = pagina.get_text("words", textpage=textpage)
                    texto_pagina_norm = _normalizar(pagina.get_text(textpage=textpage))
                    for clave, palabras_nombre in pendientes_por_nombre.items():
                        if not _todas_las_palabras_en_texto(texto_pagina_norm, palabras_nombre):
                            continue
                        for fila_rects in _buscar_por_fila(pagina, palabras_nombre, textpage=textpage):
                            fila_y0 = min(r.y0 for r in fila_rects)
                            fila_y1 = max(r.y1 for r in fila_rects)

                            cedula_de_la_fila = None
                            for wx0, wy0, wx1, wy1, wpalabra, *_resto in palabras_pagina:
                                if abs(wy0 - fila_y0) > _TOLERANCIA_FILA:
                                    continue
                                wdigitos = "".join(c for c in wpalabra if c.isdigit())
                                if not wdigitos:
                                    continue
                                wclave_cedula = _normalizar_cedula(wdigitos)
                                if wclave_cedula not in mapa_cedulas and len(wdigitos) > 1:
                                    wclave_cedula = _normalizar_cedula(wdigitos[1:])
                                if wclave_cedula in mapa_cedulas:
                                    cedula_de_la_fila = wclave_cedula
                                    break
                            if cedula_de_la_fila is not None and cedula_de_la_fila != clave[0]:
                                continue  # esa fila ya es de otro empleado conocido

                            franja = fitz.Rect(pagina.rect.x0, fila_y0 - 2, pagina.rect.x1, fila_y1 + 2)
                            candidatos_por_pendiente.setdefault(clave, []).append(
                                (ruta_pdf, pagina.number, (franja.x0, franja.y0, franja.x1, franja.y1))
                            )
            except Exception:
                pass
            finally:
                documento.close()

        for clave, candidatos in candidatos_por_pendiente.items():
            if len(candidatos) != 1:
                # ambiguo: no se arriesga, pero se deja registrado el motivo
                # exacto (y dónde) para que no quede como un simple "no
                # encontrado" sin explicación
                ubicaciones = ", ".join(
                    f"{Path(ruta).name} pág. {pagina_num + 1}" for ruta, pagina_num, _coords in candidatos
                )
                motivos_no_rescatado[clave] = f"nombre ambiguo -{len(candidatos)} coincidencias: {ubicaciones}"
                continue
            ruta_pdf_ganador, numero_pagina, coords = candidatos[0]
            cliente = clave[1]

            # si el cliente ya usó esa póliza y no es la que tiene abierta
            # ahora mismo, reabrirla la duplicaría al final del documento
            estado_previo = estado_por_cliente.get(cliente)
            archivo_actual_previo = estado_previo["archivo_actual"] if estado_previo else None
            usados_por_este_cliente = archivos_usados_por_cliente.get(cliente, set())
            if ruta_pdf_ganador in usados_por_este_cliente and ruta_pdf_ganador != archivo_actual_previo:
                motivos_no_rescatado[clave] = (
                    f"nombre encontrado en {Path(ruta_pdf_ganador).name} pág. {numero_pagina + 1}, "
                    "pero esa póliza ya se había cerrado para este cliente -no se reabre para no duplicar"
                )
                continue

            franja = fitz.Rect(*coords)

            try:
                documento_ganador = fitz.open(ruta_pdf_ganador)
            except Exception:
                continue
            try:
                pagina_ganadora = documento_ganador[numero_pagina]
                estado = _obtener_estado(cliente, pagina_ganadora)

                primera_pagina_ganadora = documento_ganador[0]
                techo_ganador = _techo_de_datos(primera_pagina_ganadora, formato)
                encabezado_ganador = None
                if techo_ganador is not None and techo_ganador > 4:
                    encabezado_ganador = {
                        "documento": documento_ganador,
                        "pagina": primera_pagina_ganadora,
                        "franja": fitz.Rect(
                            primera_pagina_ganadora.rect.x0, 0, primera_pagina_ganadora.rect.x1, techo_ganador - 2
                        ),
                    }
                estado["encabezado_actual"] = encabezado_ganador

                if estado["archivo_actual"] != ruta_pdf_ganador:
                    if estado["archivo_actual"] is not None:
                        _agregar_pie_de_poliza(estado, estado["archivo_actual"])
                        _flush_a_disco(estado)
                    estado["pagina"] = None
                    estado["archivo_actual"] = ruta_pdf_ganador
                    if encabezado_ganador is not None:
                        _agregar_bloque(
                            estado, encabezado_ganador["documento"], encabezado_ganador["pagina"],
                            encabezado_ganador["franja"], resaltar=False, repetir_encabezado=False,
                        )

                _agregar_bloque(estado, documento_ganador, pagina_ganadora, franja, resaltar=resaltar_filas)

                polizas_encontradas.setdefault(clave, set()).add(Path(ruta_pdf_ganador).name)
                archivos_usados_por_cliente.setdefault(cliente, set()).add(ruta_pdf_ganador)
                encontrados_por_nombre.add(clave)
            finally:
                documento_ganador.close()

    for clave in pendientes_por_nombre:
        if clave not in encontrados_por_nombre and clave not in motivos_no_rescatado:
            motivos_no_rescatado[clave] = "nombre no aparece en ninguna página de los PDFs de este lote"

    for estado in estado_por_cliente.values():
        if estado["archivo_actual"] is not None:
            _agregar_pie_de_poliza(estado, estado["archivo_actual"])

    archivos_por_cliente = {}
    for cliente, estado in estado_por_cliente.items():
        _flush_a_disco(estado)
        archivos_por_cliente[cliente] = str(estado["ruta_salida"])

    detalle_registros = []
    for clave, datos in registros_unicos.items():
        cliente = clave[1]
        polizas = sorted(polizas_encontradas.get(clave, set()))
        encontrado = bool(polizas)
        detalle_registros.append({
            "cedula": datos["cedula"],
            "nombre": datos["nombre"],
            "cliente": cliente,
            "polizas": polizas,
            "encontrado": encontrado,
            # "cedula": calzó por número; "nombre": rescatado en la segunda
            # pasada (conviene revisarlo); None: no se encontró
            "encontrado_por": ("nombre" if clave in encontrados_por_nombre else "cedula") if encontrado else None,
            "motivo_no_rescatado": motivos_no_rescatado.get(clave),
        })

    no_encontrados = [
        {
            "cedula": d["cedula"], "cliente": d["cliente"], "nombre": d["nombre"],
            "motivo_no_rescatado": d["motivo_no_rescatado"],
        }
        for d in detalle_registros if not d["encontrado"]
    ]

    return {
        "archivos_por_cliente": archivos_por_cliente,
        "no_encontrados": no_encontrados,
        "errores_por_archivo": errores_por_archivo,
        "detalle_registros": detalle_registros,
    }


def generar_pdf_resumen(titulo: str, secciones: list[tuple[str, list[str]]]) -> bytes:
    """PDF simple de texto con listas por sección (cédulas no encontradas,
    archivos con error, etc.), para meter dentro del zip."""
    documento = fitz.open()
    pagina = documento.new_page()
    margen = 40
    alto_util = pagina.rect.height - margen

    y = margen
    pagina.insert_text((margen, y), titulo, fontsize=14)
    y += 26

    for encabezado, lineas in secciones:
        if not lineas:
            continue
        if y > alto_util - 20:
            pagina = documento.new_page()
            y = margen
        pagina.insert_text((margen, y), encabezado, fontsize=11)
        y += 18
        for linea in lineas:
            if y > alto_util:
                pagina = documento.new_page()
                y = margen
            pagina.insert_text((margen + 10, y), f"- {linea}", fontsize=9)
            y += 14
        y += 16

    datos = documento.tobytes()
    documento.close()
    return datos


_AZUL_VMA = "0A1F3D"
_RELLENO_ENCABEZADO = PatternFill("solid", fgColor=_AZUL_VMA)
_FUENTE_ENCABEZADO = Font(color="FFFFFF", bold=True)
_BORDE_CELDA = Border(*(Side(style="thin", color="D7DBE3"),) * 4)


def _escribir_encabezado_excel(hoja, titulos: list[str]) -> None:
    hoja.append(titulos)
    for celda in hoja[1]:
        celda.fill = _RELLENO_ENCABEZADO
        celda.font = _FUENTE_ENCABEZADO
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autoajustar_columnas_excel(hoja) -> None:
    for columna in hoja.columns:
        ancho = max((len(str(celda.value)) for celda in columna if celda.value is not None), default=8)
        hoja.column_dimensions[columna[0].column_letter].width = min(ancho + 4, 60)


def _bordear_filas_excel(hoja) -> None:
    for fila in hoja.iter_rows(min_row=2, max_row=hoja.max_row):
        for celda in fila:
            celda.border = _BORDE_CELDA


def generar_excel_resumen(detalle_registros: list[dict]) -> bytes:
    """Arma el Excel de facturación (dos pestañas) que acompaña al zip del
    modo cliente, para que Facturación no tenga que abrir los PDFs uno
    por uno a contar oficiales a mano.

    detalle_registros es la lista que devuelve
    resaltar_por_cedula_y_exportar_por_cliente bajo esa misma llave: un
    dict por cada (cédula, cliente) único, con cedula, nombre, cliente,
    polizas (archivos donde se encontró), encontrado y encontrado_por
    ("cedula", "nombre" o None -"nombre" significa que se rescató en la
    segunda pasada y conviene revisarlo).

    Pestaña 1 (Resumen por Cliente): una fila por cliente con el total de
    oficiales listos para cobrar, cuántos faltan y en qué pólizas
    aparecieron. Pestaña 2 (Detalle de Oficiales): una fila por oficial,
    para conciliar reclamos puntuales.
    """
    libro = openpyxl.Workbook()

    hoja_resumen = libro.active
    hoja_resumen.title = "Resumen por Cliente"
    _escribir_encabezado_excel(hoja_resumen, [
        "Cliente", "Pólizas Involucradas", "Oficiales Encontrados", "Cédulas Faltantes", "Estado",
    ])

    por_cliente: dict[str, dict] = {}
    for d in detalle_registros:
        entrada = por_cliente.setdefault(d["cliente"], {"polizas": set(), "encontrados": 0, "faltantes": 0})
        if d["encontrado"]:
            entrada["encontrados"] += 1
            entrada["polizas"].update(d["polizas"])
        else:
            entrada["faltantes"] += 1

    total_encontrados = 0
    total_faltantes = 0
    for cliente in sorted(por_cliente):
        datos = por_cliente[cliente]
        estado = "🟢 Completo" if datos["faltantes"] == 0 else f"🟡 Pendiente ({datos['faltantes']})"
        hoja_resumen.append([
            cliente,
            ", ".join(sorted(datos["polizas"])) or "—",
            datos["encontrados"],
            datos["faltantes"],
            estado,
        ])
        total_encontrados += datos["encontrados"]
        total_faltantes += datos["faltantes"]

    hoja_resumen.append(["TOTAL GENERAL", "—", total_encontrados, total_faltantes, ""])
    for celda in hoja_resumen[hoja_resumen.max_row]:
        celda.font = Font(bold=True)
    _bordear_filas_excel(hoja_resumen)
    _autoajustar_columnas_excel(hoja_resumen)

    hoja_detalle = libro.create_sheet("Detalle de Oficiales")
    _escribir_encabezado_excel(hoja_detalle, [
        "Cédula", "Nombre Completo", "Cliente Asignado", "Póliza / Archivo de Origen", "¿Aparece en Planilla?",
    ])
    for d in sorted(detalle_registros, key=lambda d: (d["cliente"], d["cedula"])):
        if d.get("encontrado_por") == "nombre":
            estado_fila = "✅ Sí (por nombre -revisar)"
        elif d["encontrado"]:
            estado_fila = "✅ Sí"
        else:
            estado_fila = "❌ No encontrado"
        hoja_detalle.append([
            d["cedula"],
            d["nombre"] or "—",
            d["cliente"],
            ", ".join(d["polizas"]) if d["polizas"] else "(Ninguno)",
            estado_fila,
        ])
    _bordear_filas_excel(hoja_detalle)
    _autoajustar_columnas_excel(hoja_detalle)

    buffer = io.BytesIO()
    libro.save(buffer)
    return buffer.getvalue()


def procesar_carpeta(ruta_carpeta_entrada: str, nombres: list[str], ruta_carpeta_salida: str) -> list[ResultadoResaltado]:
    """Aplica resaltar_nombres_en_pdf a todos los PDFs de una carpeta."""
    carpeta_entrada = Path(ruta_carpeta_entrada)
    carpeta_salida = Path(ruta_carpeta_salida)
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    resultados = []
    for archivo_pdf in sorted(carpeta_entrada.glob("*.pdf")):
        ruta_salida = carpeta_salida / f"{archivo_pdf.stem}_resaltado.pdf"
        resultado = resaltar_nombres_en_pdf(str(archivo_pdf), nombres, str(ruta_salida))
        resultados.append(resultado)

    return resultados

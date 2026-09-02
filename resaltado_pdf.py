"""
resaltado_pdf.py

Disclaimer: buscar texto (nombres) dentro de un PDF y agregarle
anotaciones de resaltado ("highlight"), SIN modificar el resto del documento
(mismo encabezado, mismo formato, mismas fuentes, etc.).

"""

import pymupdf as fitz  # PyMuPDF (alias 'fitz' por compatibilidad con ejemplos/documentación)
import unicodedata
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ResultadoResaltado:
    """Mensaje de resultado que devuele el aplicativo al usuario"""
    archivo: str
    coincidencias_por_nombre: dict  # {"JUAN PEREZ": 3, "MARIA LOPEZ": 0}
    ruta_salida: str


def _normalizar(texto: str) -> str:
    """Normaliza el texto para tolerar diferencias de mayúsculas/minúsculas y tildes. Esto ayuda a"""
    texto = texto.strip().upper()
    texto_sin_tildes = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto_sin_tildes if not unicodedata.combining(c))


def _variantes_ene(texto: str) -> list[str]:
    """La Ñ/ñ a veces sale como espacio en blanco en el texto del PDF (pasa con MNK)."""
    if "Ñ" not in texto and "ñ" not in texto:
        return []
    return [
        texto.replace("Ñ", " ").replace("ñ", " "),
        texto.replace("Ñ", "").replace("ñ", ""),
    ]


def _buscar_con_variantes(pagina, texto: str, textpage=None):
    """Prueba la forma exacta, sin tildes, y variantes de Ñ si aplica.
    'textpage' (opcional) evita que PyMuPDF reextraiga el texto de la
    página en cada búsqueda -se puede sacar una vez con page.get_textpage()
    y reusarla para todas las búsquedas de esa página."""
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


_TOLERANCIA_FILA = 6  # puntos: cuánto pueden variar en Y las palabras de una misma fila


def _buscar_por_fila(pagina, palabras: list[str], textpage=None):
    """Busca cada palabra suelta y agrupa las que caen en la misma fila
    (nombre y apellido en columnas separadas, por ejemplo)."""
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
        textpage = pagina.get_textpage()  # se extrae 1 vez y se reusa en todas las búsquedas de la página
        for nombre in nombres:
            nombre_limpio = nombre.strip()
            if not nombre_limpio:
                continue

            coincidencias = _buscar_con_variantes(pagina, nombre_limpio, textpage=textpage)

            if coincidencias:
                for rect in coincidencias:
                    anotacion = pagina.add_highlight_annot(rect)
                    anotacion.update()
                    conteo[nombre] += 1
                continue

            # nombre completo no aparece junto -> capaz nombre/apellido
            # quedaron en columnas distintas, se busca palabra por palabra
            palabras = [p for p in nombre_limpio.split() if len(p) >= 3]
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
    """Quita ceros a la izquierda: una celda numérica del Excel (en vez de
    texto) puede perder el cero inicial de la cédula, y con esto igual calza
    contra la que viene completa desde el PDF."""
    return digitos.lstrip("0") or "0"


def _coincide_cliente(digitos_palabra: str, mapa_cedulas: dict) -> list[str]:
    """Devuelve todos los clientes a los que pertenece esta cédula (relación
    1 a N: un oficial puede cubrir turnos en más de un cliente durante la
    misma quincena). Tolera que la planilla anteponga un dígito de tipo de
    id (ej. CCSS: '0-303370238') y que a la cédula le falten ceros iniciales
    frente a la del Excel. 'mapa_cedulas' debe venir con las claves ya
    normalizadas (ver _normalizar_cedula)."""
    clientes = mapa_cedulas.get(_normalizar_cedula(digitos_palabra))
    if not clientes and len(digitos_palabra) > 1:
        clientes = mapa_cedulas.get(_normalizar_cedula(digitos_palabra[1:]))
    return clientes or []


# frases que solo salen en la fila de títulos de cada formato conocido
_PERFILES_ENCABEZADO = {
    "ccss": ["APELLIDOS Y NOMBRES", "OBSERVACIONES"],
    "mnk": ["IDENTIFICACIÓN", "OBSERVACIÓN"],
    "ins": ["Salario Reportado", "Descripción Ocupación"],
}


def _techo_por_anclas(pagina, anclas: list[str], textpage=None) -> float | None:
    """Y justo debajo de la fila donde aparecen las anclas, con margen."""
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
    """Y donde empieza la tabla de empleados (debajo de empresa/período/títulos).
    Prueba primero las anclas del formato indicado y despues las de los
    demás perfiles conocidos por si el desplegable quedó mal puesto -cada
    proveedor tiene frases únicas asi que no calzan cruzadas. Si nada
    calza cae al detector de tablas de PyMuPDF, y si tampoco encuentra
    nada mejor no mostrar encabezado que arriesgarse a mostrar una fila
    de otra persona."""
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


_MARGEN_PAGINA = 24  # puntos de margen arriba/abajo en cada página de salida
_ESPACIO_ENTRE_FILAS = 3  # separación vertical entre filas apiladas


def resaltar_por_cedula_y_exportar_por_cliente(
    rutas_pdfs: list[str],
    registros: list[dict],
    carpeta_salida: str,
    formato: str = "auto",
    resaltar_filas: bool = True,
) -> dict:
    """Busca las cédulas de 'registros' en los PDFs, recorta cada fila
    encontrada y arma un PDF por cliente, apilando filas sin dejar huecos.
    'formato' solo ayuda a ubicar el encabezado de cada página (ver
    _PERFILES_ENCABEZADO / _techo_de_datos) para no arrastrar filas de
    gente que no está en el Excel. 'resaltar_filas' controla si cada fila
    se marca con la franja amarilla o se deja tal cual salió de la
    planilla original -algunos clientes (auditorías, entidades públicas)
    piden el documento limpio, sin marcas encima.

    Una misma cédula puede estar asignada a más de un cliente en 'registros'
    (relación 1 a N -un oficial que cubrió turnos en varios puestos durante
    la quincena): en ese caso su fila se agrega al PDF de cada cliente al
    que está asignado, no solo al último que aparece en el Excel.

    El encabezado se repite una vez por archivo/póliza (no por página de
    origen): si un cliente tiene oficiales en varias páginas del mismo PDF
    solo se lleva el encabezado una vez, pero si cambia de archivo (otra
    póliza) se fuerza una hoja de salida limpia con su propio encabezado.
    Y si esa póliza termina desbordando a una segunda hoja de salida, el
    encabezado se repite ahí también para no perder el contexto."""
    mapa_cedulas: dict[str, list[str]] = {}
    for r in registros:
        cedula, cliente = r.get("cedula"), r.get("cliente")
        if not cedula or not cliente:
            continue
        clientes = mapa_cedulas.setdefault(_normalizar_cedula(cedula), [])
        if cliente not in clientes:
            clientes.append(cliente)
    estado_por_cliente: dict[str, dict] = {}
    cedulas_encontradas: set[str] = set()
    errores_por_archivo: dict[str, str] = {}

    def _obtener_estado(cliente: str, pagina_origen) -> dict:
        estado = estado_por_cliente.get(cliente)
        if estado is None:
            estado = {
                "documento": fitz.open(),
                "pagina": None,
                "y": 0.0,
                "ancho": pagina_origen.rect.width,
                "alto": pagina_origen.rect.height,
                "archivo_actual": None,  # ruta_pdf de la póliza que se está volcando ahora
                "encabezado_actual": None,  # bloque a repetir si la póliza desborda una hoja
            }
            estado_por_cliente[cliente] = estado
        return estado

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

        if estado["pagina"] is None or estado["y"] + alto_bloque > estado["alto"] - _MARGEN_PAGINA:
            estado["pagina"] = estado["documento"].new_page(width=estado["ancho"], height=estado["alto"])
            estado["y"] = _MARGEN_PAGINA
            encabezado = estado["encabezado_actual"]
            if repetir_encabezado and encabezado is not None:
                # la póliza sigue en la hoja de salida siguiente -se repite el encabezado
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

        try:
            for pagina in documento:
                textpage = pagina.get_textpage()
                palabras = pagina.get_text("words", textpage=textpage)
                techo_datos = None
                techo_calculado = False

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
                    cedulas_encontradas.add(clave_cedula)

                    fila_y0 = min(w[1] for w in palabras if abs(w[1] - y0) <= _TOLERANCIA_FILA)
                    fila_y1 = max(w[3] for w in palabras if abs(w[1] - y0) <= _TOLERANCIA_FILA)
                    franja = fitz.Rect(pagina.rect.x0, fila_y0 - 2, pagina.rect.x1, fila_y1 + 2)

                    # un mismo oficial puede estar asignado a varios clientes
                    # (relación 1 a N) -su fila se agrega al PDF de cada uno
                    for cliente in clientes:
                        vistas = franjas_vistas.setdefault(cliente, [])
                        if franja in vistas:
                            continue
                        vistas.append(franja)

                        estado = _obtener_estado(cliente, pagina)

                        if estado["archivo_actual"] != ruta_pdf:
                            # nueva póliza para este cliente -hoja limpia y encabezado propio
                            # (si el cliente ya venía de otro archivo se repite este mismo
                            # bloque de encabezado en cada hoja que la póliza necesite)
                            if not techo_calculado:
                                techo_datos = _techo_de_datos(pagina, formato, textpage=textpage)
                                techo_calculado = True
                            estado["pagina"] = None
                            estado["archivo_actual"] = ruta_pdf
                            if techo_datos is not None and techo_datos > 4:
                                franja_encabezado = fitz.Rect(pagina.rect.x0, 0, pagina.rect.x1, techo_datos - 2)
                                estado["encabezado_actual"] = {
                                    "documento": documento,
                                    "pagina": pagina,
                                    "franja": franja_encabezado,
                                }
                                _agregar_bloque(
                                    estado, documento, pagina, franja_encabezado,
                                    resaltar=False, repetir_encabezado=False,
                                )
                            else:
                                estado["encabezado_actual"] = None

                        _agregar_bloque(estado, documento, pagina, franja, resaltar=resaltar_filas)
        except Exception as error:
            errores_por_archivo[nombre_archivo] = f"No se pudo procesar el archivo: {error}"
        finally:
            documento.close()

    carpeta = Path(carpeta_salida)
    carpeta.mkdir(parents=True, exist_ok=True)
    archivos_por_cliente = {}
    for cliente, estado in estado_por_cliente.items():
        ruta = carpeta / f"{_nombre_archivo_seguro(cliente)}.pdf"
        estado["documento"].save(str(ruta))
        estado["documento"].close()
        archivos_por_cliente[cliente] = str(ruta)

    no_encontrados = [
        r for r in registros
        if r.get("cedula") and _normalizar_cedula(r["cedula"]) not in cedulas_encontradas
    ]

    return {
        "archivos_por_cliente": archivos_por_cliente,
        "no_encontrados": no_encontrados,
        "errores_por_archivo": errores_por_archivo,
    }


def generar_pdf_resumen(titulo: str, secciones: list[tuple[str, list[str]]]) -> bytes:
    """PDF simple de texto con listas por sección (ej. cédulas no encontradas,
    archivos con error), para meter dentro del zip en vez de mostrarlas en pantalla."""
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

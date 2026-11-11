import csv
import io
import logging
import os
import shutil
import subprocess
import tempfile
import threading
import time
import unicodedata
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

import openpyxl
import pymupdf as fitz  # alias tradicional de PyMuPDF
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

log = logging.getLogger("resaltado")
# Los PDF vienen de afuera: uno hecho a propósito no puede ocupar el servidor para todos (PyMuPDF atiende de a un lote).
MAXIMO_PAGINAS_POR_PDF = 5000
MAXIMO_SEGUNDOS_POR_LOTE = 15 * 60


def _desactivar_busy_loop_de_pymupdf() -> None:
    """PyMuPDF 1.28.2 llama en CADA Annot.update() a un `update_timing_test()` que es un bucle vacío de 30 000 vueltas
    (suma 0..30000 y descarta el resultado) -instrumentación olvidada que no cambia el PDF, solo quema ~1 ms por
    anotación. En un lote con miles de resaltados se van varios segundos. Lo volvemos un no-op. Con guardas por si una
    versión futura ya lo arregló o lo quitó (requirements fija pymupdf==1.28.2, así que hoy siempre aplica)."""
    try:
        if hasattr(fitz.Annot, "update_timing_test"):
            fitz.Annot.update_timing_test = staticmethod(lambda: 0)
    except Exception as error:  # nunca debe impedir que cargue el módulo
        log.debug("No se pudo neutralizar update_timing_test de PyMuPDF: %s", error)


_desactivar_busy_loop_de_pymupdf()


class _Cronometro:
    """Acumula segundos (y conteos) por fase, para ver en el log dónde se
    va el tiempo de un lote -ver deploy/benchmark_rendimiento.py."""

    def __init__(self):
        self.segundos: dict[str, float] = defaultdict(float)
        self.conteos: dict[str, int] = defaultdict(int)

    @contextmanager
    def medir(self, fase: str): 
        inicio = time.perf_counter()
        try:
            yield
        finally:
            self.segundos[fase] += time.perf_counter() - inicio

    def contar(self, que: str, cantidad: int = 1) -> None:
        self.conteos[que] += cantidad

    def resumen(self) -> dict:
        return {
            "segundos": {fase: round(s, 3) for fase, s in self.segundos.items()},
            "conteos": dict(self.conteos),
        }


@dataclass
class ResultadoResaltado:
    """Resultado de procesar un PDF en modo simple."""
    archivo: str
    coincidencias_por_nombre: dict  # {"JUAN PEREZ": 3, "MARIA LOPEZ": 0}
    ruta_salida: str
    tiempos: dict = field(default_factory=dict)
    uso_ocr: bool = False  # True si alguna página venía escaneada y se leyó con OCR


@lru_cache(maxsize=4096)
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
   
    if not palabras:
        return False
    return all(
        any(_normalizar(v) in texto_norm for v in (palabra, *_variantes_ene(palabra)))
        for palabra in palabras
    )


def _puede_aparecer_en_pagina(texto_pagina_norm: str, texto: str, palabras: list[str]) -> bool:
    candidatos_frase = [_normalizar(texto), *[_normalizar(v) for v in _variantes_ene(texto)]]
    if any(c in texto_pagina_norm for c in candidatos_frase):
        return True
    return _todas_las_palabras_en_texto(texto_pagina_norm, palabras)


_TOLERANCIA_FILA = 3  # variación en Y (puntos) tolerada para considerar la misma fila

_ALTURA_MAXIMA_FILA_SIN_VECINA = 45  # puntos: tope cuando no hay otra fila abajo contra qué recortar


def _limites_fila_por_anclas_vecinas(
    y0_objetivo: float, y0s_anclas_pagina: list[float],
) -> tuple[float | None, float | None]:

    anteriores = [y for y in y0s_anclas_pagina if y < y0_objetivo]
    siguientes = [y for y in y0s_anclas_pagina if y > y0_objetivo]
    limite_superior = (anteriores[-1] + y0_objetivo) / 2 if anteriores else None
    limite_inferior = (y0_objetivo + siguientes[0]) / 2 if siguientes else None
    return limite_superior, limite_inferior


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


def _regenerar_campos(pagina) -> None:
    """Si el PDF no trae ya generada la apariencia de un campo, PyMuPDF lo captura en blanco. Si no se puede, sale como
    venía: es un detalle visual, no frena el lote."""
    for widget in pagina.widgets() or []:
        try:
            widget.update()
        except Exception as error:
            log.debug("No se pudo regenerar un campo del PDF: %s", error)


def _demasiadas_paginas() -> str:
    return f"El PDF tiene demasiadas páginas (máximo {MAXIMO_PAGINAS_POR_PDF}). Divídalo e inténtelo de nuevo."


# --- OCR para planillas escaneadas -------------------------------------------------------------------------------
# Algunas planillas llegan escaneadas (cada página es una imagen, sin capa de texto): search_for/get_text no encuentran
# nada que resaltar. Cuando una página no trae texto, se lee con Tesseract (empaquetado en tesseract/, portable) y su
# lectura se escribe como una CAPA DE TEXTO INVISIBLE sobre la página, con la posición de cada palabra. Así todo lo de
# abajo (nombres, cédula, anclas del encabezado y del pie, franjas) funciona igual que en un PDF digital. El OCR interno
# de PyMuPDF no sirve acá: usa el modo de segmentación PSM 6, que se come las tablas densas de la CCSS (leía 1 de 20
# filas). El arranque del servidor fija TESSERACT_EXE (al tesseract.exe del paquete) y TESSDATA_PREFIX (al idioma); sin
# eso el OCR no corre y todo sigue igual (un PDF digital nunca entra al OCR).
UMBRAL_TEXTO_OCR = 10   # menos de estos caracteres en una página (sin imagen grande) = escaneada
OCR_IDIOMA = "spa"
OCR_DPI = 200           # resolución del render para OCR. A 200 el Tesseract tarda ~la mitad que a 300 y lee lo mismo en
                        # las pruebas; si un escaneo real de mala calidad leyera peor, subir a 300 (más lento).
OCR_PSM = "11"          # "texto disperso": lee bien las tablas densas; el PSM 6 (interno de PyMuPDF) se las come
OCR_CONF_MINIMA = 40    # se ignoran las palabras que Tesseract leyó con poca confianza
# Las páginas escaneadas se leen en PARALELO (el OCR de cada una es un proceso aparte de Tesseract, el cuello del lote).
# Tope de procesos simultáneos: usa la mitad de los hilos lógicos (deja la otra mitad para el resto del server -el
# portal, Licitaciones, los hilos de waitress-), con un piso de 2 y un techo de 8. En JUPITER (Xeon 4210, 20 hilos,
# 128 GB) quedan 8; en una PC chica, 2-4. Cada Tesseract se fija a 1 hilo (OMP_THREAD_LIMIT) para que el paralelismo sea
# a nivel de página y no se pisen entre sí. Variable de entorno OCR_MAX_WORKERS para ajustarlo sin tocar código.
OCR_MAX_WORKERS = max(2, min(int(os.environ.get("OCR_MAX_WORKERS") or (os.cpu_count() or 2) // 2), 8))
UMBRAL_COBERTURA_IMAGEN = 0.5  # una imagen que cubre al menos la mitad de la página = página escaneada (aunque traiga texto)
OCR_MINIMO_PALABRAS = 3        # si el OCR lee menos palabras que esto, la página escaneada se cuenta como "no leída" (aviso)

# Contadores del lote actual, para avisar en la web. _paginas_ocr: páginas leídas con OCR. _paginas_escaneadas_sin_ocr:
# páginas que parecían escaneadas pero el OCR no pudo leer (así se avisa claro en vez de resaltar mal en silencio). El
# trabajo de PyMuPDF está serializado (un lote a la vez, ver _turno_pymupdf en api/index.py), así que contadores de
# módulo alcanzan; aun así se protegen con un candado. La web los reinicia al empezar cada lote y los lee al responder.
_ocr_lock = threading.Lock()
_paginas_ocr = 0
_paginas_escaneadas_sin_ocr = 0


def reiniciar_contador_ocr() -> None:
    global _paginas_ocr, _paginas_escaneadas_sin_ocr
    with _ocr_lock:
        _paginas_ocr = 0
        _paginas_escaneadas_sin_ocr = 0


def paginas_leidas_con_ocr() -> int:
    with _ocr_lock:
        return _paginas_ocr


def paginas_escaneadas_sin_leer() -> int:
    with _ocr_lock:
        return _paginas_escaneadas_sin_ocr


def _contar_escaneada_sin_leer() -> None:
    global _paginas_escaneadas_sin_ocr
    with _ocr_lock:
        _paginas_escaneadas_sin_ocr += 1


def _contar_pagina_ocr() -> None:
    global _paginas_ocr
    with _ocr_lock:
        _paginas_ocr += 1


def _cobertura_imagen(pagina) -> float:
    """Fracción de la página cubierta por imágenes rasterizadas (0 a 1). Un escaneo es una imagen a página casi
    completa; un PDF digital, a lo sumo un logo chico. Se suman las imágenes (recortadas a la página), por si el escáner
    partió la hoja en varias bandas; se topa en 1.0."""
    area = pagina.rect.width * pagina.rect.height
    if area <= 0:
        return 0.0
    try:
        infos = pagina.get_image_info()
    except Exception:
        return 0.0
    cubierto = 0.0
    for info in infos:
        try:
            r = fitz.Rect(info["bbox"]) & pagina.rect
        except Exception:
            continue
        if r.is_valid and not r.is_empty:
            cubierto += abs(r.width * r.height)
    return min(1.0, cubierto / area)


def _pagina_parece_escaneada(pagina) -> bool:
    """Una página necesita OCR cuando una imagen la cubre casi entera -está escaneada, aunque traiga una capa de texto
    (la mala OCR del propio escáner, que si no se cuela como 'digital' y resalta filas equivocadas). También si casi no
    tiene texto y hay algo de imagen. Un PDF digital normal no entra: su texto es vectorial y sus imágenes son chicas."""
    cobertura = _cobertura_imagen(pagina)
    if cobertura >= UMBRAL_COBERTURA_IMAGEN:
        return True
    if cobertura > 0:
        return len((pagina.get_text() or "").strip()) < UMBRAL_TEXTO_OCR
    return False


def _tesseract_exe() -> str | None:
    """El tesseract.exe a usar: el empaquetado (variable TESSERACT_EXE, la fija el arranque del servidor) o uno del
    sistema; None si no hay (entonces no se hace OCR y los escaneos siguen sin texto)."""
    exe = os.environ.get("TESSERACT_EXE")
    if exe and os.path.exists(exe):
        return exe
    return shutil.which("tesseract")


def _palabras_ocr(exe: str, ruta_png: str) -> list[tuple]:
    """Corre Tesseract sobre la imagen y devuelve [(texto, left, top, ancho, alto), …] en píxeles, de las palabras
    leídas con confianza suficiente."""
    base = ruta_png[:-4] if ruta_png.lower().endswith(".png") else ruta_png
    entorno = dict(os.environ, OMP_THREAD_LIMIT="1")  # 1 hilo por Tesseract: el paralelismo es a nivel de página
    subprocess.run([exe, ruta_png, base, "--psm", OCR_PSM, "-l", OCR_IDIOMA, "tsv"],
                   check=True, capture_output=True, env=entorno)
    palabras: list[tuple] = []
    with open(base + ".tsv", encoding="utf-8", errors="replace", newline="") as f:
        for fila in csv.DictReader(f, delimiter="\t"):
            texto = (fila.get("text") or "").strip()
            try:
                conf = float(fila.get("conf", "-1"))
            except ValueError:
                conf = -1.0
            if texto and conf >= OCR_CONF_MINIMA:
                palabras.append((texto, float(fila["left"]), float(fila["top"]),
                                 float(fila["width"]), float(fila["height"])))
    return palabras


def _escribir_capa_invisible(pagina, palabras: list[tuple], escala: float) -> None:
    """Escribe cada palabra del OCR como texto invisible (render_mode=3) en su posición (escalada a puntos PDF), para
    que search_for/get_text la encuentren. Descarta detecciones repetidas (PSM 11 a veces lee dos veces la misma)."""
    vistas: set = set()
    for texto, left, top, ancho, alto in palabras:
        clave = (texto, round(left * escala), round(top * escala))
        if clave in vistas:
            continue
        vistas.add(clave)
        x = left * escala
        y = (top + alto) * escala                 # la línea base va cerca del borde inferior de la caja
        tam = max(1.0, alto * escala * 0.92)
        try:
            pagina.insert_text((x, y), texto, fontsize=tam, render_mode=3)  # render_mode 3 = invisible
        except Exception:
            pass  # un carácter que la fuente no cubra no debe cortar el lote


def _quitar_capa_de_texto(pagina) -> bool:
    """Borra la capa de texto de la página (la mala OCR del escáner) dejando la imagen del escaneo intacta, para que
    esos dígitos mal ubicados no compitan con la lectura buena del OCR ni provoquen resaltados equivocados. True si
    había texto y se quitó. Si no se puede, se sigue: peor es no leer la página."""
    if not (pagina.get_text() or "").strip():
        return False
    try:
        pagina.add_redact_annot(pagina.rect)
        pagina.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
        return True
    except Exception as error:
        log.debug("No se pudo quitar la capa de texto del escaneo: %s", error)
        return False


def _ocr_paginas_en_paralelo(escaneadas: list, exe: str) -> tuple[bool, bool]:
    """Lee las páginas escaneadas y les escribe la capa de texto invisible. El OCR de cada página (Tesseract, el cuello
    del lote) corre en PARALELO -un proceso por página-; las operaciones de PyMuPDF (quitar la capa basura, rasterizar,
    escribir la capa) van en el hilo principal, que no es seguro compartir entre hilos. Se procesa por tandas de
    OCR_MAX_WORKERS para no acumular en disco los PNG de un escaneo grande. Devuelve (hubo_ocr, hubo_cambio): hubo_cambio
    incluye las páginas a las que solo se les quitó la capa basura (se guardan igual para no sobre-resaltar con ella)."""
    hubo_ocr = False
    hubo_cambio = False
    escala = 72 / OCR_DPI
    for inicio in range(0, len(escaneadas), OCR_MAX_WORKERS):
        tanda = escaneadas[inicio:inicio + OCR_MAX_WORKERS]
        with tempfile.TemporaryDirectory(prefix="ocr_") as carpeta:
            preparadas = []  # (pagina, png) listas para OCR
            for i, pagina in enumerate(tanda):
                if _quitar_capa_de_texto(pagina):
                    hubo_cambio = True
                png = os.path.join(carpeta, f"p{i}.png")
                try:
                    pagina.get_pixmap(dpi=OCR_DPI).save(png)
                except Exception as error:
                    log.warning("No se pudo rasterizar una página para OCR (%s): se sigue sin OCR en esa página", error)
                    _contar_escaneada_sin_leer()
                    continue
                preparadas.append((pagina, png))

            def _leer(png):
                try:
                    return _palabras_ocr(exe, png)
                except Exception as error:  # un OCR que falla en una página no debe tumbar todo el lote
                    log.warning("OCR de una página falló (%s): se sigue sin OCR en esa página", error)
                    return None

            trabajadores = max(1, min(OCR_MAX_WORKERS, len(preparadas)))
            if trabajadores > 1:
                with ThreadPoolExecutor(max_workers=trabajadores) as ejecutor:
                    lecturas = list(ejecutor.map(lambda par: _leer(par[1]), preparadas))
            else:
                lecturas = [_leer(par[1]) for par in preparadas]

            for (pagina, _png), palabras in zip(preparadas, lecturas):
                if not palabras or len(palabras) < OCR_MINIMO_PALABRAS:
                    _contar_escaneada_sin_leer()
                    continue
                _escribir_capa_invisible(pagina, palabras, escala)
                _contar_pagina_ocr()
                hubo_ocr = True
                hubo_cambio = True
    return hubo_ocr, hubo_cambio


def _asegurar_buscable(ruta_pdf: str, carpeta_destino) -> tuple[str, bool]:
    """Si el PDF tiene páginas escaneadas (sin texto) y hay Tesseract, escribe en carpeta_destino una copia con capa de
    texto OCR (mismo nombre de archivo) y devuelve (esa ruta, True). Si ya es buscable o no hay OCR, (ruta original,
    False). Trabajar sobre una copia buscable deja que TODO el pipeline —búsqueda, encabezado, pie— la lea como un PDF
    digital (el pie, por ejemplo, re-abre el archivo, así que la capa tiene que estar en el archivo, no solo en memoria)."""
    if not _tesseract_exe():
        return ruta_pdf, False
    try:
        documento = fitz.open(ruta_pdf)
    except Exception:
        return ruta_pdf, False
    try:
        if documento.is_encrypted:
            return ruta_pdf, False  # el chequeo de contraseña (con mensaje claro) lo hace la función de entrada
        escaneadas = [p for p in documento if _pagina_parece_escaneada(p)]
        if not escaneadas:
            return ruta_pdf, False
        hubo_ocr, hubo_cambio = _ocr_paginas_en_paralelo(escaneadas, _tesseract_exe())
        if not hubo_cambio:
            return ruta_pdf, False
        carpeta = Path(carpeta_destino)
        carpeta.mkdir(parents=True, exist_ok=True)
        destino = carpeta / Path(ruta_pdf).name
        documento.save(str(destino))
        return str(destino), hubo_ocr
    finally:
        documento.close()


def resaltar_nombres_en_pdf(ruta_pdf: str, nombres: list[str], ruta_salida: str) -> ResultadoResaltado:
    """Resalta cada nombre de la lista en el PDF y guarda la copia en ruta_salida."""
    cronometro = _Cronometro()
    with cronometro.medir("ocr"):
        ruta_pdf, hubo_ocr = _asegurar_buscable(ruta_pdf, Path(ruta_salida).parent / "_buscables_ocr")
    with cronometro.medir("abrir_pdf"):
        documento = fitz.open(ruta_pdf)
    if documento.is_encrypted:
        documento.close()
        raise ValueError(
            "El PDF está protegido con contraseña. Quite la protección e inténtelo de nuevo."
        )
    if documento.page_count > MAXIMO_PAGINAS_POR_PDF:
        documento.close()
        raise ValueError(_demasiadas_paginas())

    conteo = {nombre: 0 for nombre in nombres}

    for pagina in documento:
        cronometro.contar("paginas")
        with cronometro.medir("extraer_texto"):
            textpage = pagina.get_textpage()
            texto_pagina_norm = _normalizar(pagina.get_text(textpage=textpage))

        for nombre in nombres:
            nombre_limpio = nombre.strip()
            if not nombre_limpio:
                continue

            with cronometro.medir("prefiltro_nombres"):
                palabras = [p for p in nombre_limpio.split() if len(p) >= 3]
                puede_aparecer = _puede_aparecer_en_pagina(texto_pagina_norm, nombre_limpio, palabras)
            if not puede_aparecer:
                continue
            cronometro.contar("busquedas_en_pagina")

            with cronometro.medir("buscar"):
                coincidencias = _buscar_con_variantes(pagina, nombre_limpio, textpage=textpage)

            if coincidencias:
                with cronometro.medir("resaltar"):
                    for rect in coincidencias:
                        anotacion = pagina.add_highlight_annot(rect)
                        anotacion.update()
                        conteo[nombre] += 1
                continue

            # no aparece junto: puede que nombre y apellido queden en columnas distintas
            with cronometro.medir("buscar_por_fila"):
                filas = _buscar_por_fila(pagina, palabras, textpage=textpage)
            with cronometro.medir("resaltar"):
                for fila in filas:
                    for rect in fila:
                        anotacion = pagina.add_highlight_annot(rect)
                        anotacion.update()
                    conteo[nombre] += 1

    with cronometro.medir("guardar_disco"):
        documento.save(ruta_salida)
        documento.close()

    return ResultadoResaltado(
        archivo=Path(ruta_pdf).name,
        coincidencias_por_nombre=conteo,
        ruta_salida=ruta_salida,
        tiempos=cronometro.resumen(),
        uso_ocr=hubo_ocr,
    )


def _nombre_archivo_seguro(texto: str) -> str:
    limpio = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in texto).strip()
    return limpio or "SIN_CLIENTE"


def _normalizar_cedula(digitos: str) -> str:
    """Quita ceros a la izquierda (si el Excel guardó la cédula como
    número, se pierde el cero inicial)."""
    return digitos.lstrip("0") or "0"


def _coincide_cliente(digitos_palabra: str, mapa_cedulas: dict) -> list[str]:
    """Devuelve la lista de clientes que tienen esa cédula, o [] si no"""
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


def _y0s_anclas_fila(pagina, textpage=None) -> list[float]:
    """Y0 de cada renglón donde aparece un número de cédula (9+ dígitos)"""
    limite_x = pagina.rect.x0 + pagina.rect.width * 0.35
    palabras = pagina.get_text("words", textpage=textpage)
    return sorted({
        w[1] for w in palabras
        if w[0] < limite_x and len("".join(c for c in w[4] if c.isdigit())) >= 9
    })


def _techo_de_datos(pagina, formato: str = "auto", textpage=None) -> float | None:
    """Y justo debajo de la primera fila de datos, sin incluir el renglón"""
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
           
            techo_crudo = techo - 14  # y1 real del título, sin el margen
            y0s_filas = _y0s_anclas_fila(pagina, textpage=textpage)
            candidatas = [y for y in y0s_filas if techo_crudo < y < techo]
            if candidatas:
                techo = candidatas[0] - 2
            return techo

    for estrategia in ("lines_strict", "lines", "text"):
        try:
            tablas = pagina.find_tables(strategy=estrategia).tables
        except Exception as error:
            log.debug("La estrategia %s no pudo leer las tablas: %s", estrategia, error)
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

    margen_superior = y0_total - _MARGEN_ARRIBA_PIE
    encabezado_repetido = _franja_por_perfiles(pagina, _PERFILES_ENCABEZADO, formato, textpage=textpage)
    if encabezado_repetido is not None:
        _y0_encabezado, y1_encabezado = encabezado_repetido
        limite_tras_encabezado = y1_encabezado + 4
        if margen_superior < limite_tras_encabezado < y0_total:
            margen_superior = limite_tras_encabezado

    
    y0s_filas = _y0s_anclas_fila(pagina, textpage=textpage)
    anteriores = [y for y in y0s_filas if y < y0_total]
    if anteriores:
        limite_tras_fila = anteriores[-1] + 10
        
        if margen_superior < limite_tras_fila < y0_total:
            margen_superior = limite_tras_fila

    return fitz.Rect(pagina.rect.x0, margen_superior, pagina.rect.x1, y1_total + _MARGEN_ABAJO_TOTAL)


def _franja_leyenda_en_pagina(pagina, formato: str = "auto", textpage=None) -> "fitz.Rect | None":
    """Franja de la leyenda/firma en esta página, o None si no aparece."""
    leyenda = _franja_por_perfiles(pagina, _PERFILES_LEYENDA_PIE, formato, textpage=textpage)
    if leyenda is None:
        return None
    y0_leyenda, _y1_leyenda = leyenda
    return fitz.Rect(pagina.rect.x0, y0_leyenda - _MARGEN_ARRIBA_PIE, pagina.rect.x1, pagina.rect.height)


def _extender_leyenda_para_incluir_total(
    franja_leyenda: "fitz.Rect", franja_total: "fitz.Rect | None", misma_pagina: bool,
) -> "fitz.Rect":
    """Cuando la leyenda cae pegada al total, se extiende para que el"""
    if (
        misma_pagina
        and franja_total is not None
        and franja_total.y0 < franja_leyenda.y0 < franja_total.y1
    ):
        return fitz.Rect(franja_leyenda.x0, franja_total.y0, franja_leyenda.x1, franja_leyenda.y1)
    return franja_leyenda


def _recortar_total_antes_de_leyenda(
    franja_total: "fitz.Rect", franja_leyenda: "fitz.Rect | None", misma_pagina: bool,
) -> "fitz.Rect":
    
    if (
        misma_pagina
        and franja_leyenda is not None
        and franja_total.y1 > franja_leyenda.y0 > franja_total.y0
    ):
        y1_real_del_total = franja_total.y1 - _MARGEN_ABAJO_TOTAL
        nuevo_y1 = max(franja_leyenda.y0, y1_real_del_total)
        if nuevo_y1 < franja_total.y1:
            return fitz.Rect(franja_total.x0, franja_total.y0, franja_total.x1, nuevo_y1)
    return franja_total


def _franjas_pie_de_pagina(pagina, formato: str = "auto", textpage=None) -> list["fitz.Rect"]:
   
    franjas = []

    franja_total = _franja_total_en_pagina(pagina, formato, textpage=textpage)
    if franja_total is None:
        return franjas
    franjas.append(franja_total)

    franja_leyenda = _franja_leyenda_en_pagina(pagina, formato, textpage=textpage)
    if franja_leyenda is not None:
        franjas.append(franja_leyenda)

    return franjas


def _compactar_pdf(ruta: Path) -> None:
    """Reescribe el PDF completo sin objetos duplicados y con todo
    comprimido. show_pdf_page arrastra las imágenes de la página de origen
    (el logo del encabezado) aunque el recorte no las muestre, y cada
    guardado incremental por cambio de póliza deja otra copia sin
    comprimir: en un lote real el 97 % del peso de salida eran imágenes,
    720 de ellas repetidas dentro del mismo PDF. Si falla, se queda el
    archivo como estaba -más pesado, pero correcto."""
    temporal = ruta.with_name(ruta.stem + ".compacto.pdf")
    try:
        documento = fitz.open(str(ruta))
        try:
            # las planillas de MNK traen el mismo logo 3 veces, idéntico
            # salvo por /Name (im1, im2, im3) -eso basta para que garbage=4
            # no las considere iguales. /Name en imágenes es obsoleto desde
            # PDF 1.1 y ningún visor lo usa, se puede quitar sin riesgo.
            for xref in range(1, documento.xref_length()):
                if (
                    documento.xref_get_key(xref, "Subtype")[1] == "/Image"
                    and documento.xref_get_key(xref, "Name")[0] != "null"
                ):
                    documento.xref_set_key(xref, "Name", "null")
            documento.save(
                str(temporal), garbage=4, deflate=True, deflate_images=True, deflate_fonts=True,
            )
        finally:
            documento.close()
        temporal.replace(ruta)
    except Exception:
        temporal.unlink(missing_ok=True)


_MARGEN_PAGINA = 24  # margen arriba/abajo de cada página de salida
_ESPACIO_ENTRE_FILAS = 3


def _construir_mapas_cedula(registros: list[dict]) -> tuple[dict, dict, dict]:
    """A partir de los registros del Excel/CSV arma:
    - registros_unicos: un registro por (cédula, cliente), sin duplicados.
    - mapa_cedulas: cédula normalizada (o número de asegurado) -> clientes que la reclaman.
    - mapa_id_a_cedula_real: para reportar el número de asegurado bajo la cédula real del cliente.
    Comparten esta lógica resaltar_por_cedula_y_exportar_por_cliente() y
    resaltar_por_cedula_sin_recortar()."""
    registros_unicos: dict[tuple[str, str], dict] = {}
    for r in registros:
        cedula, cliente = r.get("cedula"), r.get("cliente")
        if not cedula or not cliente:
            continue
        clave = (_normalizar_cedula(cedula), cliente)
        if clave not in registros_unicos:
            registros_unicos[clave] = {
                "cedula": cedula,
                "cliente": cliente,
                "nombre": r.get("nombre", ""),
                "numero_asegurado": r.get("numero_asegurado", ""),
            }

    mapa_cedulas: dict[str, list[str]] = {}
    mapa_id_a_cedula_real: dict[str, str] = {}
    for clave_cedula, cliente in registros_unicos:
        clientes = mapa_cedulas.setdefault(clave_cedula, [])
        if cliente not in clientes:
            clientes.append(cliente)
        mapa_id_a_cedula_real.setdefault(clave_cedula, clave_cedula)

        numero_asegurado_raw = registros_unicos[(clave_cedula, cliente)].get("numero_asegurado")
        if numero_asegurado_raw:
            numero_asegurado = _normalizar_cedula(numero_asegurado_raw)
            clientes_asegurado = mapa_cedulas.setdefault(numero_asegurado, [])
            if cliente not in clientes_asegurado:
                clientes_asegurado.append(cliente)
            mapa_id_a_cedula_real.setdefault(numero_asegurado, clave_cedula)

    return registros_unicos, mapa_cedulas, mapa_id_a_cedula_real


def _nombre_de_fila_coincide(palabras: list, y0: float, nombre_esperado: str) -> bool:
    """¿El nombre esperado del Excel es plausible en la fila que arranca en y0?
    Conservador a propósito: solo DESCARTA (False) cuando el nombre esperado trae
    >=2 palabras significativas, la fila trae su propio nombre (>=2 palabras de
    letras) y NINGUNA palabra del esperado aparece en la fila. Eso pasa cuando el
    OCR confundió una cédula con la de OTRA persona (visto en escaneos densos: la
    cédula de alguien se lee como la de otro empleado). Si el Excel no trae nombre
    útil o la fila no trae un nombre claro, no filtra -mejor no arriesgar un
    descarte. Un descarte deja esa fila sin resaltar (sale en "no encontrados"),
    que es seguro; un match equivocado facturaría a la persona equivocada."""
    tokens_esperados = [t for t in (nombre_esperado or "").split() if len(t) >= 3]
    if len(tokens_esperados) < 2:
        return True
    palabras_fila = [w[4] for w in palabras if abs(w[1] - y0) <= _TOLERANCIA_FILA]
    palabras_con_letras = [p for p in palabras_fila if len(p) >= 3 and any(c.isalpha() for c in p)]
    if len(palabras_con_letras) < 2:
        return True
    texto_fila_norm = _normalizar(" ".join(palabras_fila))
    return any(
        any(_normalizar(v) in texto_fila_norm for v in (tok, *_variantes_ene(tok)))
        for tok in tokens_esperados
    )


def _vecinos_distancia_1(digitos: str):
    """Variantes con UN dígito cambiado (misma longitud) -los típicos errores del OCR (0/8, 1/7, 5/6, 3/8…)."""
    for i in range(len(digitos)):
        for d in "0123456789":
            if d != digitos[i]:
                yield digitos[:i] + d + digitos[i + 1:]


def _recuperar_cedula_por_cercania(digitos: str, cedulas_oficiales: frozenset) -> str | None:
    """La ÚNICA cédula oficial a distancia-1 (misma longitud, 1 dígito distinto) de lo que leyó el OCR, o None si no hay
    o si hay más de una (ambiguo -no se arriesga). Solo tiene sentido con cédulas (>=8 dígitos), no con salarios."""
    if len(digitos) < 8:
        return None
    candidatas = {v for v in _vecinos_distancia_1(digitos) if v in cedulas_oficiales}
    return next(iter(candidatas)) if len(candidatas) == 1 else None


def _localizar_coincidencias_de_cedula(
    pagina, palabras: list, y0s_cedulas_pagina: list[float], mapa_cedulas: dict, mapa_id_a_cedula_real: dict,
    mapa_nombres: dict | None = None, verificar_nombre: bool = False, cedulas_oficiales: frozenset | None = None,
):
    for x0, y0, x1, y1, palabra, *_resto in palabras:
        digitos = "".join(c for c in palabra if c.isdigit())
        if not digitos:
            continue
        clientes = _coincide_cliente(digitos, mapa_cedulas)
        clave_cedula_hallada = _normalizar_cedula(digitos)
        if clave_cedula_hallada not in mapa_cedulas and len(digitos) > 1:
            clave_cedula_hallada = _normalizar_cedula(digitos[1:])
        clave_cedula = mapa_id_a_cedula_real.get(clave_cedula_hallada, clave_cedula_hallada)
        hallado_por_numero_asegurado = clave_cedula != clave_cedula_hallada

        if clientes:
            # Match exacto. Solo en escaneados (verificar_nombre): el OCR puede leer la cédula de alguien como la de otro
            # empleado, así que se confirma que el nombre de la fila concuerda con el del Excel antes de darlo por bueno.
            if verificar_nombre and mapa_nombres is not None:
                if not _nombre_de_fila_coincide(palabras, y0, mapa_nombres.get(clave_cedula, "")):
                    continue
        else:
            # Sin match exacto. SOLO en escaneados y SI el nombre confirma, se recupera un error de 1 dígito del OCR
            # (devuelve a un oficial que si no quedaría sin resaltar). El nombre es la red de seguridad: sin él no se
            # recupera, así que un número ajeno leído cerca de una cédula oficial no puede colarse.
            if not (verificar_nombre and mapa_nombres is not None and cedulas_oficiales):
                continue
            recuperada = _recuperar_cedula_por_cercania(_normalizar_cedula(digitos), cedulas_oficiales)
            if recuperada is None or not _nombre_de_fila_coincide(palabras, y0, mapa_nombres.get(recuperada, "")):
                continue
            clientes = mapa_cedulas.get(recuperada) or []
            if not clientes:
                continue
            clave_cedula = mapa_id_a_cedula_real.get(recuperada, recuperada)
            hallado_por_numero_asegurado = False

        franja = _franja_de_fila(pagina, palabras, y0, y0s_cedulas_pagina)

        yield franja, clientes, clave_cedula, hallado_por_numero_asegurado


def _franja_de_fila(pagina, palabras: list, y0: float, y0s_anclas_pagina: list[float]) -> "fitz.Rect":
    """Franja de ancho completo del renglón que arranca en y0: todas las
    palabras de esa misma altura, recortada a medio camino de las filas
    vecinas (para no comerse media fila de arriba o de abajo cuando el
    renglón ocupa dos líneas)."""
    misma_fila = [w for w in palabras if abs(w[1] - y0) <= _TOLERANCIA_FILA]
    fila_y0 = min((w[1] for w in misma_fila), default=y0)
    fila_y1 = max((w[3] for w in misma_fila), default=y0 + 10)

    limite_superior, limite_inferior = _limites_fila_por_anclas_vecinas(y0, y0s_anclas_pagina)
    if limite_superior is not None:
        fila_y0 = max(fila_y0, limite_superior)
    if limite_inferior is not None:
        fila_y1 = min(fila_y1, limite_inferior)
    else:
        fila_y1 = min(fila_y1, fila_y0 + _ALTURA_MAXIMA_FILA_SIN_VECINA)

    return fitz.Rect(pagina.rect.x0, fila_y0 - 2, pagina.rect.x1, fila_y1 + 2)


def _construir_detalle_registros(
    registros_unicos: dict, polizas_encontradas: dict, encontrados_por_numero_asegurado: set,
) -> tuple[list[dict], list[dict]]:
    
    detalle_registros = []
    for clave, datos in registros_unicos.items():
        cliente = clave[1]
        polizas = sorted(polizas_encontradas.get(clave, set()))
        encontrado = bool(polizas)
        detalle_registros.append({
            "cedula": datos["cedula"],
            "nombre": datos["nombre"],
            "cliente": cliente,
            "numero_asegurado": datos.get("numero_asegurado", ""),
            "polizas": polizas,
            "encontrado": encontrado,
            "encontrado_por": (
                "numero_asegurado" if clave in encontrados_por_numero_asegurado else "cedula"
            ) if encontrado else None,
        })

    no_encontrados = [
        {"cedula": d["cedula"], "cliente": d["cliente"], "nombre": d["nombre"]}
        for d in detalle_registros if not d["encontrado"]
    ]

    return detalle_registros, no_encontrados


def resaltar_por_cedula_y_exportar_por_cliente(
    rutas_pdfs: list[str],
    registros: list[dict],
    carpeta_salida: str,
    formato: str = "auto",
    resaltar_filas: bool = True,
    localizar=None,
) -> dict:
    """localizar: opcional, para recortar por otro criterio que no sea la
    cédula del Excel (ver resaltar_nombres_y_recortar). Recibe
    (pagina, palabras, y0s_anclas_fila, textpage) y devuelve tuplas
    (franja, grupos, clave, hallado_por_numero_asegurado) -cada grupo es
    un PDF de salida."""
    cronometro = _Cronometro()
    registros_unicos, mapa_cedulas, mapa_id_a_cedula_real = _construir_mapas_cedula(registros)
    mapa_nombres = {}
    for (ced, _cli), datos in registros_unicos.items():
        mapa_nombres.setdefault(ced, datos.get("nombre", ""))
    cedulas_oficiales = frozenset(mapa_cedulas)
    archivo_con_ocr = {"actual": False}  # el archivo que se procesa ahora vino escaneado -se verifica el nombre de la fila
    if localizar is None:
        def localizar(pagina, palabras, y0s_anclas, _textpage):
            return _localizar_coincidencias_de_cedula(
                pagina, palabras, y0s_anclas, mapa_cedulas, mapa_id_a_cedula_real,
                mapa_nombres=mapa_nombres, verificar_nombre=archivo_con_ocr["actual"],
                cedulas_oficiales=cedulas_oficiales,
            )

    estado_por_cliente: dict[str, dict] = {}
    polizas_encontradas: dict[tuple[str, str], set] = {}
    errores_por_archivo: dict[str, str] = {}
    # quiénes se encontraron por su número de asegurado (no por su cédula/
    # DIMEX) -conviene que Facturación lo sepa, igual que con "encontrado
    # por nombre"
    encontrados_por_numero_asegurado: set = set()

    carpeta = Path(carpeta_salida)
    carpeta.mkdir(parents=True, exist_ok=True)
    carpeta_buscables = carpeta / "_buscables_ocr"   # copias con capa OCR de los escaneos (las limpia la web con el temporal)
    hubo_ocr_lote = False

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
            cronometro.contar("reaperturas_pdf_cliente")
            with cronometro.medir("reabrir_pdf_cliente"):
                estado["documento"] = fitz.open(str(estado["ruta_salida"]))
            estado["pagina"] = None

    def _flush_a_disco(estado: dict) -> None:

        documento = estado["documento"]
        if documento is None:
            return
        cronometro.contar("guardados_pdf_cliente")
        with cronometro.medir("guardar_pdf_cliente"):
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
            
            estado["pagina"] = estado["documento"].new_page(width=estado["ancho"], height=estado["alto"])
            estado["y"] = _MARGEN_PAGINA
            encabezado = estado["encabezado_actual"]
            if repetir_encabezado and encabezado is not None:
                _agregar_bloque(
                    estado, encabezado["documento"], encabezado["pagina"], encabezado["franja"],
                    resaltar=False, repetir_encabezado=False,
                )

        destino = fitz.Rect(0, estado["y"], estado["ancho"], estado["y"] + alto_bloque)
        with cronometro.medir("show_pdf_page"):
            estado["pagina"].show_pdf_page(destino, documento_origen, pagina_origen.number, clip=franja)
        if resaltar:
            with cronometro.medir("resaltar_fila"):
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
     
        if ruta_pdf_saliente in pixmaps_pie_por_archivo:
            return pixmaps_pie_por_archivo[ruta_pdf_saliente]

        cronometro.contar("pies_renderizados")
        with cronometro.medir("renderizar_pie"):
            return _pixmaps_pie_de_poliza_sin_cache(ruta_pdf_saliente)

    def _pixmaps_pie_de_poliza_sin_cache(ruta_pdf_saliente: str) -> list[tuple]:
        resultado: list[tuple] = []
        try:
            documento_pie = fitz.open(ruta_pdf_saliente)
        except Exception:
            pixmaps_pie_por_archivo[ruta_pdf_saliente] = resultado
            return resultado
        try:
            if not documento_pie.is_encrypted:
                ultimo_indice = documento_pie.page_count - 1
                num_a_revisar = min(6, documento_pie.page_count)

                
                bloques_pie: list[tuple[int, float, tuple]] = []

                pagina_total = franja_total = None
                for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
                    candidata = documento_pie[indice]
                    franja = _franja_total_en_pagina(candidata, formato)
                    if franja is not None:
                        pagina_total, franja_total = candidata, franja
                        break

                pagina_leyenda = franja_leyenda = None
                for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
                    candidata = documento_pie[indice]
                    franja = _franja_leyenda_en_pagina(candidata, formato)
                    if franja is not None:
                        pagina_leyenda, franja_leyenda = candidata, franja
                        break
                misma_pagina = (
                    pagina_total is not None
                    and pagina_leyenda is not None
                    and pagina_leyenda.number == pagina_total.number
                )

                franja_total_original = franja_total
                franja_total = _recortar_total_antes_de_leyenda(franja_total, franja_leyenda, misma_pagina)

                if pagina_total is not None:
                    # si el PDF no trae ya generada la apariencia del campo,
                    # PyMuPDF lo captura en blanco -forzar que la regenere
                    _regenerar_campos(pagina_total)
                    pixmap = pagina_total.get_pixmap(clip=franja_total, matrix=fitz.Matrix(2, 2))
                    bloques_pie.append((
                        pagina_total.number, franja_total.y0,
                        (pixmap.tobytes("png"), pagina_total.rect.width, franja_total.height),
                    ))

                if pagina_leyenda is not None:
                    franja_leyenda = _extender_leyenda_para_incluir_total(
                        franja_leyenda, franja_total_original, misma_pagina,
                    )
                    pixmap = pagina_leyenda.get_pixmap(clip=franja_leyenda, matrix=fitz.Matrix(2, 2))
                    bloques_pie.append((
                        pagina_leyenda.number, franja_leyenda.y0,
                        (pixmap.tobytes("png"), pagina_leyenda.rect.width, franja_leyenda.height),
                    ))

                bloques_pie.sort(key=lambda b: (b[0], b[1]))
                resultado.extend(bloque for _pagina, _y0, bloque in bloques_pie)

               
                if not resultado and formato in ("mnk", "ccss"):
                    pagina_cierre = indice_cierre = None
                    for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
                        candidata = documento_pie[indice]
                        y0s_filas = _y0s_anclas_fila(candidata)
                        if y0s_filas:
                            pagina_cierre, indice_cierre = candidata, indice
                            y0_ultima_fila = y0s_filas[-1]
                            break

                    if pagina_cierre is not None:
                        _regenerar_campos(pagina_cierre)
                        y0_inicio_cierre = y0_ultima_fila + 14
                      
                        palabras_resto = [
                            w for w in pagina_cierre.get_text("words") if w[1] > y0_inicio_cierre
                        ]
                        if palabras_resto:
                            y1_cierre = min(max(w[3] for w in palabras_resto) + 10, pagina_cierre.rect.height)
                        else:
                            y1_cierre = pagina_cierre.rect.height
                        franja_cierre = fitz.Rect(
                            pagina_cierre.rect.x0, y0_inicio_cierre, pagina_cierre.rect.x1, y1_cierre,
                        )
                        # si de verdad no hay nada ahí (página de prueba
                        # sin pie de página, por ejemplo), mejor no
                        # agregar un recorte de puro espacio en blanco
                        hay_contenido = bool(pagina_cierre.get_text(clip=franja_cierre).strip())
                        if franja_cierre.height > 4 and hay_contenido:
                            pixmap = pagina_cierre.get_pixmap(clip=franja_cierre, matrix=fitz.Matrix(2, 2))
                            resultado.append((pixmap.tobytes("png"), pagina_cierre.rect.width, franja_cierre.height))

                        # todas las hojas que sigan y que NO tengan
                        # ninguna fila de empleado son continuación del
                        # cierre (firma, aviso legal, datos de contacto)
                        # -pueden ser una o varias, se copian TODAS,
                        # completas y en orden, para no perder ninguna ni
                        # mezclar el orden en el que aparecen en el
                        # original
                        for siguiente_indice in range(indice_cierre + 1, documento_pie.page_count):
                            pagina_siguiente = documento_pie[siguiente_indice]
                            if _y0s_anclas_fila(pagina_siguiente):
                                break  # ya empezó otra sección con filas de empleado
                            if not pagina_siguiente.get_text().strip():
                                continue  # hoja realmente en blanco, se salta
                            _regenerar_campos(pagina_siguiente)
                            pixmap = pagina_siguiente.get_pixmap(matrix=fitz.Matrix(2, 2))
                            resultado.append((
                                pixmap.tobytes("png"), pagina_siguiente.rect.width, pagina_siguiente.rect.height,
                            ))
        except Exception:
            log.warning("No se pudo copiar el pie de %s; el archivo sale sin él", ruta_pdf_saliente, exc_info=True)
        finally:
            documento_pie.close()

        pixmaps_pie_por_archivo[ruta_pdf_saliente] = resultado
        return resultado

    def _agregar_pie_de_poliza(estado: dict, ruta_pdf_saliente: str) -> None:
        """Cierra la póliza que este cliente está dejando atrás con su
        propio total, antes de pasar a la siguiente."""
        bloques_pie = _pixmaps_pie_de_poliza(ruta_pdf_saliente)
        with cronometro.medir("insertar_pie"):
            for png_bytes, ancho_pagina, alto_franja in bloques_pie:
                _agregar_imagen_cacheada(estado, png_bytes, ancho_pagina, alto_franja)

    vence = time.monotonic() + MAXIMO_SEGUNDOS_POR_LOTE
    for ruta_pdf in rutas_pdfs:
        nombre_archivo = Path(ruta_pdf).name
        if time.monotonic() > vence:
            errores_por_archivo[nombre_archivo] = "Se acabó el tiempo del lote antes de llegar a este archivo."
            continue
        with cronometro.medir("ocr"):
            ruta_pdf, _uso = _asegurar_buscable(ruta_pdf, carpeta_buscables)
            hubo_ocr_lote = hubo_ocr_lote or _uso
        archivo_con_ocr["actual"] = _uso  # en un escaneado se verifica el nombre de la fila (_localizar_coincidencias_de_cedula)
        try:
            with cronometro.medir("abrir_pdf"):
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
        if documento.page_count > MAXIMO_PAGINAS_POR_PDF:
            documento.close()
            errores_por_archivo[nombre_archivo] = _demasiadas_paginas()
            continue

        encabezado_archivo = None
        encabezado_calculado = False

        try:
            for pagina in documento:
                cronometro.contar("paginas")
                with cronometro.medir("extraer_texto"):
                    textpage = pagina.get_textpage()
                    palabras = pagina.get_text("words", textpage=textpage)
                    y0s_cedulas_pagina = _y0s_anclas_fila(pagina, textpage=textpage)

                with cronometro.medir("buscar_cedulas"):
                    coincidencias_pagina = list(localizar(pagina, palabras, y0s_cedulas_pagina, textpage))
                cronometro.contar("filas_encontradas", len(coincidencias_pagina))

                franjas_vistas: dict[str, list] = {}
                for franja, clientes, clave_cedula, hallado_por_numero_asegurado in coincidencias_pagina:
                    for cliente in clientes:
                        polizas_encontradas.setdefault((clave_cedula, cliente), set()).add(nombre_archivo)
                        if hallado_por_numero_asegurado:
                            encontrados_por_numero_asegurado.add((clave_cedula, cliente))

                        vistas = franjas_vistas.setdefault(cliente, [])
                        if franja in vistas:
                            continue
                        vistas.append(franja)

                        estado = _obtener_estado(cliente, pagina)

                        if estado["archivo_actual"] != ruta_pdf:
                            if estado["archivo_actual"] is not None:
                                cronometro.contar("cambios_de_poliza")
                                _agregar_pie_de_poliza(estado, estado["archivo_actual"])
                                # El PDF del cliente se deja ABIERTO entre pólizas: antes se guardaba+cerraba+reabría en
                                # cada cambio (cientos de reaperturas, el mayor costo del modo cliente). El contenido de
                                # cada póliza se copia con show_pdf_page mientras su PDF origen está abierto, así que no
                                # hace falta persistir el del cliente hasta el final. El guardado único es _flush_a_disco
                                # al cerrar el lote.
                            if not encabezado_calculado:
                                primera_pagina = documento[0]
                                with cronometro.medir("detectar_encabezado"):
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
                                with cronometro.medir("copiar_filas"):
                                    _agregar_bloque(
                                        estado, encabezado_archivo["documento"], encabezado_archivo["pagina"],
                                        encabezado_archivo["franja"], resaltar=False, repetir_encabezado=False,
                                    )

                        with cronometro.medir("copiar_filas"):
                            _agregar_bloque(estado, documento, pagina, franja, resaltar=resaltar_filas)
        except Exception as error:
            errores_por_archivo[nombre_archivo] = f"No se pudo procesar el archivo: {error}"
        finally:
            documento.close()

    for estado in estado_por_cliente.values():
        if estado["archivo_actual"] is not None:
            _agregar_pie_de_poliza(estado, estado["archivo_actual"])

    archivos_por_cliente = {}
    for cliente, estado in estado_por_cliente.items():
        _flush_a_disco(estado)
        with cronometro.medir("compactar_pdf_cliente"):
            _compactar_pdf(estado["ruta_salida"])
        archivos_por_cliente[cliente] = str(estado["ruta_salida"])

    detalle_registros, no_encontrados = _construir_detalle_registros(
        registros_unicos, polizas_encontradas, encontrados_por_numero_asegurado,
    )

    return {
        "archivos_por_cliente": archivos_por_cliente,
        "no_encontrados": no_encontrados,
        "errores_por_archivo": errores_por_archivo,
        "detalle_registros": detalle_registros,
        "uso_ocr": hubo_ocr_lote,
        "tiempos": cronometro.resumen(),
    }


_GRUPO_MODO_NOMBRE = "Nombres encontrados"


def _se_solapan(a: "fitz.Rect", b: "fitz.Rect") -> bool:
    """Dos franjas son el mismo renglón si comparten más de la mitad del alto."""
    alto_comun = min(a.y1, b.y1) - max(a.y0, b.y0)
    return alto_comun > 0.5 * min(a.height, b.height)


def _localizar_filas_de_nombres(pagina, palabras, y0s_anclas, textpage, nombres, conteo_por_archivo):
    """Franja del renglón de cada coincidencia de los nombres buscados, con
    el mismo criterio de búsqueda que resaltar_nombres_en_pdf (variantes de
    la Ñ, nombre y apellido en columnas distintas). Si dos nombres caen en
    el mismo renglón, el renglón sale una sola vez."""
    conteo = conteo_por_archivo.setdefault(Path(pagina.parent.name).name, {n: 0 for n in nombres})
    texto_pagina_norm = _normalizar(pagina.get_text(textpage=textpage))

    franjas_pagina: list = []
    for nombre in nombres:
        palabras_nombre = [p for p in nombre.split() if len(p) >= 3]
        if not _puede_aparecer_en_pagina(texto_pagina_norm, nombre, palabras_nombre):
            continue
        rects = _buscar_con_variantes(pagina, nombre, textpage=textpage)
        if rects:
            y0s = [r.y0 for r in rects]
        else:
            y0s = [min(r.y0 for r in fila) for fila in _buscar_por_fila(pagina, palabras_nombre, textpage=textpage)]

        franjas_nombre: list = []
        for y0 in y0s:
            franja = _franja_de_fila(pagina, palabras, y0, y0s_anclas)
            if any(_se_solapan(franja, f) for f in franjas_nombre):
                continue  # el mismo nombre dos veces en el mismo renglón cuenta una
            franjas_nombre.append(franja)
            conteo[nombre] += 1
            if not any(_se_solapan(franja, f) for f, _n in franjas_pagina):
                franjas_pagina.append((franja, nombre))

    franjas_pagina.sort(key=lambda par: par[0].y0)  # en el orden de la planilla
    for franja, nombre in franjas_pagina:
        yield franja, [_GRUPO_MODO_NOMBRE], nombre, False


def resaltar_nombres_y_recortar(
    rutas_pdfs: list[str],
    nombres: list[str],
    carpeta_salida: str,
    formato: str = "auto",
    resaltar_filas: bool = True,
) -> dict:
    """Modo por nombre con recorte: un solo PDF con el encabezado de cada
    planilla, solo los renglones donde aparece alguno de los nombres, y el
    pie de cada póliza -el mismo armado que el PDF por cliente."""
    nombres = [n.strip() for n in nombres if n.strip()]
    conteo_por_archivo: dict[str, dict[str, int]] = {}

    def localizar(pagina, palabras, y0s_anclas, textpage):
        return _localizar_filas_de_nombres(pagina, palabras, y0s_anclas, textpage, nombres, conteo_por_archivo)

    resultado = resaltar_por_cedula_y_exportar_por_cliente(
        rutas_pdfs, [], carpeta_salida, formato, resaltar_filas=resaltar_filas, localizar=localizar,
    )

    # las planillas donde no apareció nadie también van en el resumen, en cero
    for ruta in rutas_pdfs:
        nombre_archivo = Path(ruta).name
        if nombre_archivo not in resultado["errores_por_archivo"]:
            conteo_por_archivo.setdefault(nombre_archivo, {n: 0 for n in nombres})

    return {
        "ruta_pdf": resultado["archivos_por_cliente"].get(_GRUPO_MODO_NOMBRE),
        "coincidencias_por_archivo": conteo_por_archivo,
        "errores_por_archivo": resultado["errores_por_archivo"],
        "uso_ocr": resultado.get("uso_ocr", False),
        "tiempos": resultado["tiempos"],
    }


def resaltar_por_cedula_sin_recortar(
    rutas_pdfs: list[str],
    registros: list[dict],
    carpeta_salida: str,
) -> dict:
    cronometro = _Cronometro()
    registros_unicos, mapa_cedulas, mapa_id_a_cedula_real = _construir_mapas_cedula(registros)
    mapa_nombres = {}
    for (ced, _cli), datos in registros_unicos.items():
        mapa_nombres.setdefault(ced, datos.get("nombre", ""))
    cedulas_oficiales = frozenset(mapa_cedulas)

    polizas_encontradas: dict[tuple[str, str], set] = {}
    encontrados_por_numero_asegurado: set = set()
    errores_por_archivo: dict[str, str] = {}
    archivos_resaltados: dict[str, str] = {}

    carpeta = Path(carpeta_salida)
    carpeta.mkdir(parents=True, exist_ok=True)
    carpeta_buscables = carpeta / "_buscables_ocr"   # copias con capa OCR de los escaneos (las limpia la web con el temporal)
    hubo_ocr_lote = False

    vence = time.monotonic() + MAXIMO_SEGUNDOS_POR_LOTE
    for ruta_pdf in rutas_pdfs:
        nombre_archivo = Path(ruta_pdf).name
        if time.monotonic() > vence:
            errores_por_archivo[nombre_archivo] = "Se acabó el tiempo del lote antes de llegar a este archivo."
            continue
        with cronometro.medir("ocr"):
            ruta_pdf, _uso = _asegurar_buscable(ruta_pdf, carpeta_buscables)
            hubo_ocr_lote = hubo_ocr_lote or _uso
        try:
            with cronometro.medir("abrir_pdf"):
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
        if documento.page_count > MAXIMO_PAGINAS_POR_PDF:
            documento.close()
            errores_por_archivo[nombre_archivo] = _demasiadas_paginas()
            continue

        try:
            for pagina in documento:
                cronometro.contar("paginas")
                with cronometro.medir("extraer_texto"):
                    textpage = pagina.get_textpage()
                    palabras = pagina.get_text("words", textpage=textpage)
                    y0s_cedulas_pagina = _y0s_anclas_fila(pagina, textpage=textpage)

                with cronometro.medir("buscar_cedulas"):
                    coincidencias_pagina = list(_localizar_coincidencias_de_cedula(
                        pagina, palabras, y0s_cedulas_pagina, mapa_cedulas, mapa_id_a_cedula_real,
                        mapa_nombres=mapa_nombres, verificar_nombre=_uso, cedulas_oficiales=cedulas_oficiales,
                    ))
                cronometro.contar("filas_encontradas", len(coincidencias_pagina))

                franjas_ya_resaltadas: list = []
                for franja, clientes, clave_cedula, hallado_por_numero_asegurado in coincidencias_pagina:
                    for cliente in clientes:
                        polizas_encontradas.setdefault((clave_cedula, cliente), set()).add(nombre_archivo)
                        if hallado_por_numero_asegurado:
                            encontrados_por_numero_asegurado.add((clave_cedula, cliente))

                    # el resaltado en sí no depende del cliente -si la
                    # misma fila calzó para más de uno, se marca una sola
                    # vez, no una anotación encimada por cada cliente
                    if franja in franjas_ya_resaltadas:
                        continue
                    franjas_ya_resaltadas.append(franja)
                    with cronometro.medir("resaltar"):
                        anotacion = pagina.add_highlight_annot(franja)
                        anotacion.update()

            ruta_salida = carpeta / f"{Path(nombre_archivo).stem}_resaltado.pdf"
            with cronometro.medir("guardar_disco"):
                documento.save(str(ruta_salida))
            archivos_resaltados[nombre_archivo] = str(ruta_salida)
        except Exception as error:
            errores_por_archivo[nombre_archivo] = f"No se pudo procesar el archivo: {error}"
        finally:
            documento.close()

    detalle_registros, no_encontrados = _construir_detalle_registros(
        registros_unicos, polizas_encontradas, encontrados_por_numero_asegurado,
    )

    return {
        "archivos_resaltados": archivos_resaltados,
        "no_encontrados": no_encontrados,
        "errores_por_archivo": errores_por_archivo,
        "detalle_registros": detalle_registros,
        "uso_ocr": hubo_ocr_lote,
        "tiempos": cronometro.resumen(),
    }


def generar_pdf_resumen(titulo: str, secciones: list[tuple[str, list[str]]]) -> bytes:
    
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


def _sin_formulas(libro) -> None:
    """Los nombres y clientes vienen de los archivos subidos: uno que empiece con = se guardaría como fórmula y Excel lo
    ejecutaría al abrir el resumen. Quedan como texto, tal cual."""
    for hoja in libro:
        for fila in hoja.iter_rows():
            for celda in fila:
                if celda.data_type == "f":
                    celda.data_type = "s"


def generar_excel_resumen(detalle_registros: list[dict]) -> bytes:
    
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
        "Cédula", "Número de Asegurado", "Nombre Completo", "Cliente Asignado",
        "Póliza / Archivo de Origen", "¿Aparece en Planilla?",
    ])
    for d in sorted(detalle_registros, key=lambda d: (d["cliente"], d["cedula"])):
        if d.get("encontrado_por") == "numero_asegurado":
            estado_fila = "✅ Sí (por número de asegurado)"
        elif d["encontrado"]:
            estado_fila = "✅ Sí"
        else:
            estado_fila = "❌ No encontrado"
        hoja_detalle.append([
            d["cedula"],
            d.get("numero_asegurado") or "—",
            d["nombre"] or "—",
            d["cliente"],
            ", ".join(d["polizas"]) if d["polizas"] else "(Ninguno)",
            estado_fila,
        ])
    _bordear_filas_excel(hoja_detalle)
    _autoajustar_columnas_excel(hoja_detalle)

    buffer = io.BytesIO()
    _sin_formulas(libro)
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

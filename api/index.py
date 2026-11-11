"""
api/index.py
"""

import base64
import csv
import io
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import threading
import time
import unicodedata
import zipfile
from collections import defaultdict, deque
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlsplit

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
from flask import Flask, Response, g, jsonify, render_template, request
from werkzeug.exceptions import HTTPException

from resaltado_pdf import (
    generar_excel_resumen,
    generar_pdf_resumen,
    paginas_escaneadas_sin_leer,
    paginas_leidas_con_ocr,
    reiniciar_contador_ocr,
    resaltar_nombres_en_pdf,
    resaltar_nombres_y_recortar,
    resaltar_por_cedula_sin_recortar,
    resaltar_por_cedula_y_exportar_por_cliente,
)

from api import seguridad_web, sesion_vma

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024  # 60 MB, de sobra para una planilla


_DEBUG_GUARDAR_SUBIDAS = os.environ.get("DEBUG_GUARDAR_SUBIDAS") == "1"
_CARPETA_DEBUG_SUBIDAS = Path(tempfile.gettempdir()) / "debug_subidas"


def _guardar_copia_debug(ruta_origen: Path, nombre_archivo: str) -> None:
    if not _DEBUG_GUARDAR_SUBIDAS:
        return
    try:
        _CARPETA_DEBUG_SUBIDAS.mkdir(parents=True, exist_ok=True)
        shutil.copy(ruta_origen, _CARPETA_DEBUG_SUBIDAS / nombre_archivo)
    except Exception as error:
        app.logger.warning("No se pudo guardar copia de depuración de %s: %s", nombre_archivo, error)


def _limpiar_temporales_antiguos(segundos_vida: int = 3600) -> int:
    limite = time.time() - segundos_vida
    temp_dir = Path(tempfile.gettempdir())
    borradas = 0

    for prefijo in ("resaltado_simple_", "resaltado_cliente_"):
        for carpeta in temp_dir.glob(f"{prefijo}*"):
            try:
                if carpeta.is_dir() and carpeta.stat().st_mtime < limite:
                    shutil.rmtree(carpeta, ignore_errors=True)
                    borradas += 1
            except Exception as error:
                app.logger.debug("No se pudo borrar carpeta temporal %s: %s", carpeta, error)

    if borradas > 0:
        app.logger.info("Mantenimiento: se purgaron %d carpeta(s) temporales huérfanas", borradas)

    return borradas


_limpiar_temporales_antiguos()  # una pasada al arrancar el proceso


def _rss_pico_mb() -> float | None:
    """Pico de memoria del proceso (worker de gunicorn) desde que arrancó,
    en MB. None en Windows, donde no existe el módulo resource."""
    try:
        import resource
    except ImportError:
        return None
    return round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024, 1)  # Linux lo da en KB


def _sumar_tiempos(acumulado: dict, tiempos: dict) -> None:
    """Suma el resumen de un _Cronometro (resaltado_pdf) dentro de otro."""
    for seccion in ("segundos", "conteos"):
        destino = acumulado.setdefault(seccion, {})
        for clave, valor in tiempos.get(seccion, {}).items():
            destino[clave] = round(destino.get(clave, 0) + valor, 3)


def _registrar_tiempos(modo: str, fases: dict, tiempos_lote: dict, bytes_subidos: int, bytes_zip: int) -> None:
    """Una sola línea en el log por lote, para leerla con journalctl:
    fases del servidor (guardar subidas, procesar, zip) + fases internas
    del procesamiento."""
    app.logger.info("tiempos %s: %s", modo, json.dumps({
        "fases_servidor": {fase: round(s, 3) for fase, s in fases.items()},
        "procesamiento": tiempos_lote,
        "mb_subidos": round(bytes_subidos / 1024 / 1024, 2),
        "mb_zip": round(bytes_zip / 1024 / 1024, 2),
        "rss_pico_proceso_mb": _rss_pico_mb(),
    }, ensure_ascii=False))

# solo cuenta /api/procesar (lo pesado). Toda la oficina de VMA sale a
# internet por la misma IP (Fortinet), así que el límite es para el
# grupo completo, no por persona -por eso 30 y no 10. La vista previa
# (/api/detectar-modo-excel) no cuenta: se dispara cada vez que alguien
# elige un Excel y solo lee la hoja con openpyxl.
_RUTAS_CON_LIMITE = ("/api/procesar",)
_LIMITE_SOLICITUDES_POR_IP = 30
_VENTANA_LIMITE_SEGUNDOS = 5 * 60
_historial_solicitudes_por_ip: dict[str, deque] = defaultdict(deque)

# PyMuPDF no soporta usarse desde varios hilos a la vez. Gunicorn corre con
# varios hilos por worker para que la página y la vista previa respondan
# mientras un lote se procesa; este candado deja que solo un lote por
# worker toque PyMuPDF a la vez (el paralelismo real lo dan los workers,
# que son procesos separados).
_candado_pymupdf = threading.Lock()


@contextmanager
def _turno_pymupdf():
    inicio = time.perf_counter()
    with _candado_pymupdf:
        espera = time.perf_counter() - inicio
        if espera > 0.5:
            app.logger.info("el lote esperó %.1f s a que terminara otro en este worker", espera)
        yield


def _ip_supero_el_limite(ip: str) -> bool:
    ahora = time.time()
    historial = _historial_solicitudes_por_ip[ip]
    while historial and historial[0] < ahora - _VENTANA_LIMITE_SEGUNDOS:
        historial.popleft()
    if len(historial) >= _LIMITE_SOLICITUDES_POR_IP:
        return True
    historial.append(ahora)
    return False


# Inicio de sesión: con SESION_VMA_SECRETO (la misma clave del portal de servicios) solo entra quien inició sesión en el
# portal, que deja la cookie vma_sesion para todas las apps del servidor. En cada pedido se le pregunta al portal, por
# 127.0.0.1, quién es esa sesión (sesion_vma.ConsultaPortal recuerda la respuesta 30 segundos): una cuenta desactivada
# o una sesión cerrada deja de entrar aquí en menos de un minuto. Sin la clave la app queda abierta como antes (Vercel,
# desarrollo). PORTAL_PUERTO: dónde corre el portal en este mismo servidor (8080 por defecto). Además la cuenta tiene
# que tener acceso a Planillas (id «planillas» en las herramientas del portal); si no, vuelve al portal.
app.config["SESION_VMA_SECRETO"] = os.environ.get("SESION_VMA_SECRETO") or None
app.config["PORTAL_PUERTO"] = int(os.environ.get("PORTAL_PUERTO") or 8080)
app.config["CONSULTA_PORTAL"] = None  # se arma con el primer pedido (las pruebas ponen una de mentira)
if app.config["SESION_VMA_SECRETO"]:
    sesion_vma.revisar_secreto(app.config["SESION_VMA_SECRETO"])


def revisar_configuracion() -> None:
    seguridad_web.revisar_login_en_servidor(bool(app.config["SESION_VMA_SECRETO"]))


revisar_configuracion()
if _DEBUG_GUARDAR_SUBIDAS:
    app.logger.warning("DEBUG_GUARDAR_SUBIDAS está encendido: se guardan copias de las planillas subidas (con datos de "
                       "personas) en %s. Apágalo en producción.", _CARPETA_DEBUG_SUBIDAS)

_HOST_VALIDO = re.compile(r"^[A-Za-z0-9.-]{1,253}$")
_RUTAS_SIN_SESION = ("/salud", "/static/")  # la revisión del portal y del instalador, los estilos y JSZip
SERVICIO_PORTAL = "planillas"  # el id de esta app en servicios.json del portal
_CSP_SIN_HTML = seguridad_web.csp(estilos_en_linea=False)
_MAXIMO_EXCEL_DESCOMPRIMIDO = 500 * 1024 * 1024  # un .xlsx es un ZIP: unos KB pueden inflarse a varios GB
_RESERVADOS_WINDOWS = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}
_LARGO_MAXIMO_NOMBRE = 150


def _nombre_seguro(subido: str | None) -> str:
    """El nombre de un archivo subido, sin nada que lo saque de la carpeta temporal (../) ni que Windows interprete
    distinto: «x.pdf:oculto» (flujo alternativo de NTFS) o «NUL.pdf» (un dispositivo). Lo normal queda igual."""
    nombre = re.split(r"[\\/]", subido or "")[-1]
    nombre = re.sub(r'[\x00-\x1f\x7f<>:"|?*]', "", nombre).strip(" .")
    base, extension = os.path.splitext(nombre)
    if base.lower() in _RESERVADOS_WINDOWS:
        base = f"_{base}"
    return f"{base[:_LARGO_MAXIMO_NOMBRE - len(extension)]}{extension}" or "archivo"


def _es_pdf(archivo) -> bool:
    """Por el nombre y por el contenido: un programa renombrado a .pdf no pasa."""
    if not _nombre_seguro(archivo.filename).lower().endswith(".pdf"):
        return False
    inicio = archivo.stream.read(1024)
    archivo.stream.seek(0)
    return b"%PDF-" in inicio


def _host() -> str:
    """El nombre o IP con que se llegó a la app (sin puerto); si no es válido, localhost."""
    host = urlsplit(f"//{request.host}").hostname or ""
    host = host if _HOST_VALIDO.match(host) else "localhost"
    return seguridad_web.host_permitido(host)


@app.before_request
def _rechazar_otro_origen():
    """Con inicio de sesión, lo que manda datos solo se acepta desde la propia app: otra página abierta en el navegador
    no puede subir ni procesar nada en nombre de quien tiene la sesión (CSRF)."""
    if not app.config["SESION_VMA_SECRETO"] or request.method in ("GET", "HEAD", "OPTIONS"):
        return None
    origen = request.headers.get("Origin") or request.headers.get("Referer")
    if seguridad_web.mismo_origen(origen, request.host):
        return None
    return jsonify(error="Este pedido viene de otro sitio."), 403


@app.before_request
def _exigir_sesion_del_portal():
    secreto = app.config["SESION_VMA_SECRETO"]
    if not secreto or request.path.startswith(_RUTAS_SIN_SESION):
        return None
    consulta = app.config["CONSULTA_PORTAL"]
    if consulta is None:
        consulta = app.config["CONSULTA_PORTAL"] = sesion_vma.ConsultaPortal(
            sesion_vma.puerto_para_consultar(app.config["PORTAL_PUERTO"]), secreto)
    # Los pedidos automáticos de la página (X-Sin-Actividad: 1) no alargan la sesión.
    actividad = request.headers.get(sesion_vma.ENCABEZADO_SIN_ACTIVIDAD) != "1"
    try:
        g.persona = consulta.persona_de_encabezado(request.headers.get("Cookie"), actividad)
    except sesion_vma.PortalNoResponde as error:
        app.logger.warning("El portal no respondió la consulta de sesión: %s", error)
        mensaje = "El portal de servicios no responde. Pruebe de nuevo en un momento."
        if request.path.startswith("/api/"):
            return jsonify(error=mensaje), 503
        return Response(mensaje, status=503, mimetype="text/plain")
    if g.persona and not sesion_vma.puede_entrar(g.persona, SERVICIO_PORTAL):
        if request.path.startswith("/api/"):
            return jsonify(error="Tu cuenta no tiene acceso a Planillas. Pídeselo a un administrador del portal."), 403
        return Response(status=302, headers={
            "Location": sesion_vma.url_sin_acceso(_host(), app.config["PORTAL_PUERTO"], SERVICIO_PORTAL)})
    if g.persona:
        return None
    if request.path.startswith("/api/"):
        return jsonify(error="La sesión venció. Inicie sesión de nuevo en el portal de servicios."), 401
    # El portal solo devuelve a direcciones de este mismo servidor; lo revisa él. El host va por _host() (lista
    # blanca), no el del pedido, para no reflejar un Host falso en la dirección de vuelta.
    if request.method == "GET":
        puerto = urlsplit(f"//{request.host}").port
        destino_host = f"{_host()}:{puerto}" if puerto else _host()
        volver = f"{sesion_vma.esquema()}://{destino_host}{request.full_path.rstrip('?')}"
    else:
        volver = ""
    return Response(status=302, headers={"Location": sesion_vma.url_entrar(_host(), app.config["PORTAL_PUERTO"], volver)})


@app.after_request
def _encabezados_de_seguridad(respuesta: Response) -> Response:
    """Los encabezados de seguridad comunes de las apps de Grupo VMA (api/seguridad_web.py). La CSP de cada página
    admite justo sus scripts: los archivos de /static y los escritos en la plantilla, por su hash. La página y la API
    van sin caché (traen datos de personas)."""
    for nombre, valor in seguridad_web.encabezados():
        respuesta.headers.setdefault(nombre, valor)
    respuesta.headers["Server"] = "Planillas"  # sin la versión de Werkzeug/Python
    if respuesta.mimetype == "text/html" and not respuesta.is_streamed:
        politica = seguridad_web.csp(respuesta.get_data(as_text=True), estilos_en_linea=False)
    else:
        politica = _CSP_SIN_HTML
    respuesta.headers.setdefault("Content-Security-Policy", politica)
    if request.path == "/" or request.path.startswith("/api/"):
        respuesta.headers.setdefault("Cache-Control", "no-store")
    return respuesta


@app.before_request
def _limitar_solicitudes_api():
    if request.path not in _RUTAS_CON_LIMITE:
        return None
    ip = request.remote_addr or "desconocida"
    if _ip_supero_el_limite(ip):
        app.logger.warning("límite de solicitudes alcanzado para %s", ip)
        return jsonify(
            error="Demasiadas solicitudes seguidas desde esta conexión. Espere unos minutos e intente de nuevo."
        ), 429
    return None


def _enviar_zip_y_borrar(ruta: Path, nombre_descarga: str) -> Response:
    """Envía el ZIP temporal por partes y lo borra cuando la respuesta se cierra (ya enviada, con el archivo cerrado).

    No se usa send_file: cuando el servidor sabe enviar archivos por su cuenta (waitress en Windows, gunicorn en
    Linux), Flask le pasa el archivo directo y nunca avisa cuándo terminó, así que no había cuándo borrarlo; y
    borrarlo antes falla en Windows mientras está abierto (WinError 32). Así sirve igual en los dos sistemas y no
    carga el ZIP entero en memoria."""
    def contenido():
        with open(ruta, "rb") as archivo:
            while bloque := archivo.read(256 * 1024):
                yield bloque

    def borrar():
        try:
            ruta.unlink(missing_ok=True)
        except OSError as error:  # no debería pasar; un ZIP temporal que queda no justifica romper la descarga
            app.logger.warning("No se pudo borrar el ZIP temporal %s: %s", ruta, error)

    respuesta = Response(contenido(), mimetype="application/zip")
    respuesta.headers["Content-Disposition"] = f'attachment; filename="{nombre_descarga}"'
    respuesta.headers["Content-Length"] = str(ruta.stat().st_size)
    respuesta.call_on_close(borrar)  # corre después de cerrar contenido(), o aunque nunca se haya empezado a enviar
    return respuesta


@app.route("/", methods=["GET"])
def index():
    persona = g.get("persona")
    salir = sesion_vma.url_salir(_host(), app.config["PORTAL_PUERTO"]) if persona else None
    portal = sesion_vma.url_portal(_host(), app.config["PORTAL_PUERTO"]) if persona else None
    return render_template("index.html", persona=persona, salir=salir, portal=portal)


@app.route("/salud", methods=["GET"])
def salud():
    return jsonify(estado="ok")


@app.errorhandler(Exception)
def _manejar_error(error):
    if isinstance(error, HTTPException) and error.code == 413:
        return jsonify(error="El archivo es demasiado grande (máximo 60 MB en total)."), 413
    if isinstance(error, HTTPException):
        return jsonify(error=error.description or "Ocurrió un error."), error.code
    app.logger.exception("Error no manejado al procesar la solicitud")
    return jsonify(error="Ocurrió un error inesperado al procesar la solicitud."), 500


def _excel_no_se_pudo_leer(archivo) -> ValueError:
    return ValueError(
        f"El archivo Excel '{_nombre_seguro(archivo.filename)}' no se pudo leer. "
        "Verifique que sea un .xlsx válido y no esté dañado."
    )


def _trae_doctype(contenido) -> bool:
    """Por partes y hasta el final: se podría esconder detrás de un comentario largo."""
    anterior = b""
    while bloque := contenido.read(64 * 1024):
        if b"<!DOCTYPE" in (anterior + bloque).upper():
            return True
        anterior = bloque[-8:]
    return False


def _revisar_excel(archivo) -> None:
    """Antes de abrirlo: que no se infle de más al descomprimirlo y que su XML no traiga un DOCTYPE, que un Excel real
    no usa y es por donde entran los ataques de entidades XML. Si ni siquiera es un ZIP, lo dirá openpyxl."""
    try:
        with zipfile.ZipFile(archivo) as libro:
            partes = libro.infolist()
            if sum(parte.file_size for parte in partes) > _MAXIMO_EXCEL_DESCOMPRIMIDO:
                raise ValueError(f"El archivo Excel '{_nombre_seguro(archivo.filename)}' es demasiado grande.")
            for parte in partes:
                if parte.filename.lower().endswith((".xml", ".rels")):
                    with libro.open(parte) as contenido:
                        if _trae_doctype(contenido):
                            raise _excel_no_se_pudo_leer(archivo)
    except (zipfile.BadZipFile, RuntimeError, NotImplementedError):
        pass
    finally:
        archivo.seek(0)


def _abrir_libro(archivo):
    _revisar_excel(archivo)
    try:
        return openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    except Exception:
        raise _excel_no_se_pudo_leer(archivo)


def _decodificar_csv(contenido_bytes: bytes) -> str:

    for codificacion in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return contenido_bytes.decode(codificacion)
        except UnicodeDecodeError:
            continue
    return contenido_bytes.decode("latin-1", errors="replace")


def _detectar_delimitador(texto: str) -> str:

    primera_linea = texto.splitlines()[0] if texto else ""
    if primera_linea.count(";") > primera_linea.count(","):
        return ";"
    return ","


def _filas_desde_archivo(archivo):
    
    nombre = _nombre_seguro(archivo.filename).lower()
    archivo.seek(0)

    if nombre.endswith((".xlsx", ".xlsm")):
        libro = _abrir_libro(archivo)
        try:
            hoja = libro.active
            for fila in hoja.iter_rows(values_only=True):
                yield fila
        finally:
            libro.close()
    elif nombre.endswith(".csv"):
        contenido_bytes = archivo.read()
        archivo.seek(0)
        texto = _decodificar_csv(contenido_bytes)
        delimitador = _detectar_delimitador(texto)
        lector = csv.reader(io.StringIO(texto), delimiter=delimitador)
        for fila in lector:
            yield tuple(fila)
    else:
        raise ValueError("Formato de archivo no soportado. Suba un .xlsx, .xlsm o .csv.")


def _nombres_desde_excel(archivo) -> list[str]:

    nombres = []
    for i, fila in enumerate(_filas_desde_archivo(archivo)):
        valor = next((c for c in fila if c not in (None, "")), None)
        if valor is None:
            continue
        texto = str(valor).strip()
        if i == 0 and texto.upper() in ("NOMBRE", "NOMBRES", "NOMBRE COMPLETO"):
            continue
        if texto:
            nombres.append(texto)
    return nombres


def _combinar_nombres(texto_nombres: str, archivo_excel) -> list[str]:

    candidatos = [n.strip() for n in texto_nombres.split(",") if n.strip()]
    if archivo_excel and archivo_excel.filename:
        candidatos.extend(_nombres_desde_excel(archivo_excel))

    vistos = set()
    nombres = []
    for nombre in candidatos:
        clave = nombre.upper()
        if clave not in vistos:
            vistos.add(clave)
            nombres.append(nombre)
    return nombres



_SINONIMOS_CEDULA = ["IDENTIFICACION", "CEDULA", "ID", "DOCUMENTO", "IDENTIFICACION FISCAL", "NUMERO"]
_SINONIMOS_CLIENTE = ["CLIENTE", "CUENTA"]
_SINONIMOS_NOMBRE = ["NOMBRE", "NOMBRES", "EMPLEADO", "COLABORADOR", "NOMBRE COMPLETO"]
_SINONIMOS_NUMERO_ASEGURADO = [
    "NUMERO DE ASEGURADO", "NUMERO ASEGURADO", "ASEGURADO", "NUM ASEGURADO", "N ASEGURADO",
]


def _normalizar_encabezado(valor) -> str:
    texto = str(valor or "").strip().upper()
    sin_tildes = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in sin_tildes if not unicodedata.combining(c))


def _indice_por_sinonimos(encabezado, sinonimos: list[str]) -> int | None:
    """Recorre la lista de sinónimos en orden de prioridad, no las
    columnas de izquierda a derecha -así 'Cliente' gana sobre 'Empresa'
    aunque 'Empresa' venga primero en el Excel."""
    normalizados = [_normalizar_encabezado(valor) if valor else "" for valor in encabezado]
    for sinonimo in sinonimos:
        for i, valor in enumerate(normalizados):
            if valor == sinonimo:
                return i
    return None


def _extraer_cedula_limpia(valor) -> str:
    """Solo los dígitos, quitando el .0 que deja Excel cuando la columna
    quedó como número en vez de texto (ej. 303370238.0 -> '303370238')."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    texto = str(valor).strip()
    if "." in texto:
        partes = texto.split(".")
        if len(partes) == 2 and partes[1] == "0":
            texto = partes[0]
    return "".join(c for c in texto if c.isdigit())


def _registros_desde_excel(archivo):
    
    filas = _filas_desde_archivo(archivo)

    encabezado = next(filas, None)
    if encabezado is None:
        return None

    indice_cedula = _indice_por_sinonimos(encabezado, _SINONIMOS_CEDULA)
    indice_cliente = _indice_por_sinonimos(encabezado, _SINONIMOS_CLIENTE)
    indice_nombre = _indice_por_sinonimos(encabezado, _SINONIMOS_NOMBRE)
    indice_numero_asegurado = _indice_por_sinonimos(encabezado, _SINONIMOS_NUMERO_ASEGURADO)

    if indice_cedula is None or indice_cliente is None:
        columnas_con_datos = sum(1 for valor in encabezado if valor not in (None, ""))
        if columnas_con_datos > 1:
            raise ValueError(
                "El archivo no contiene una columna de Identificación/Cédula y Cliente válida."
            )
        return None

    registros = []
    for fila in filas:
        valor_cedula = fila[indice_cedula] if indice_cedula < len(fila) else None
        cliente = str(fila[indice_cliente]).strip() if indice_cliente < len(fila) and fila[indice_cliente] else ""
        cedula = _extraer_cedula_limpia(valor_cedula)
        if not cedula or not cliente:
            continue
        nombre = ""
        if indice_nombre is not None and indice_nombre < len(fila) and fila[indice_nombre]:
            nombre = str(fila[indice_nombre]).strip()
        numero_asegurado = ""
        if (
            indice_numero_asegurado is not None
            and indice_numero_asegurado < len(fila)
            and fila[indice_numero_asegurado]
        ):
            numero_asegurado = _extraer_cedula_limpia(fila[indice_numero_asegurado])
        registros.append(
            {
                "cedula": cedula,
                "cliente": cliente,
                "nombre": nombre,
                "numero_asegurado": numero_asegurado,
            }
        )

    return registros


@app.route("/api/detectar-modo-excel", methods=["POST"])
def detectar_modo_excel():
    """Preview para el frontend: qué modo se va a activar apenas se elige
    el Excel, sin esperar a subir los PDFs y procesar todo."""
    archivo_excel = request.files.get("excel")
    if not archivo_excel or not archivo_excel.filename:
        return jsonify(error="No se recibió ningún archivo Excel."), 400
    if not archivo_excel.filename.lower().endswith((".xlsx", ".xlsm", ".csv")):
        return jsonify(error="El archivo debe ser un Excel (.xlsx / .xlsm) o un archivo .csv."), 400

    try:
        registros = _registros_desde_excel(archivo_excel)
    except ValueError as error:
        return jsonify(error=str(error)), 400

    if registros:
        return jsonify(modo="cliente", total_registros=len(registros))
    return jsonify(modo="simple")


def _nombre_zip_sin_colision(nombre: str, nombres_usados: set) -> str:
    """Si el nombre ya se usó en este zip (dos PDFs de origen con el
    mismo nombre, por ejemplo), le agrega un sufijo numérico."""
    if nombre not in nombres_usados:
        nombres_usados.add(nombre)
        return nombre

    stem = Path(nombre).stem
    extension = Path(nombre).suffix
    contador = 1
    while True:
        candidato = f"{stem} ({contador}){extension}"
        if candidato not in nombres_usados:
            nombres_usados.add(candidato)
            return candidato
        contador += 1


def _procesar_modo_simple(nombres: list[str], archivos):
    reiniciar_contador_ocr()
    coincidencias_por_archivo: dict[str, dict] = {}
    errores_por_archivo: dict[str, str] = {}
    nombres_zip_usados: set[str] = set()
    fases: dict[str, float] = defaultdict(float)
    tiempos_lote: dict = {}
    bytes_subidos = 0
    inicio_total = time.perf_counter()
    carpeta_temporal = Path(tempfile.mkdtemp(prefix="resaltado_simple_"))
    archivo_zip_temporal = tempfile.NamedTemporaryFile(suffix=".zip", prefix="resaltado_simple_zip_", delete=False)
    ruta_zip = Path(archivo_zip_temporal.name)
    archivo_zip_temporal.close()

    try:
        with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, archivo in enumerate(archivos):
                nombre_archivo = _nombre_seguro(archivo.filename)
                if not _es_pdf(archivo):
                    errores_por_archivo[nombre_archivo] = "No es un archivo PDF."
                    continue
                ruta_entrada = carpeta_temporal / f"{i}_{nombre_archivo}"
                inicio = time.perf_counter()
                archivo.save(ruta_entrada)
                fases["guardar_subidas"] += time.perf_counter() - inicio
                bytes_subidos += ruta_entrada.stat().st_size

                nombre_salida = f"{Path(nombre_archivo).stem}_resaltado.pdf"
                ruta_salida = carpeta_temporal / f"{i}_{nombre_salida}"

                try:
                    inicio = time.perf_counter()
                    resultado = resaltar_nombres_en_pdf(str(ruta_entrada), nombres, str(ruta_salida))
                    fases["procesar"] += time.perf_counter() - inicio
                    _sumar_tiempos(tiempos_lote, resultado.tiempos)
                    arcname = _nombre_zip_sin_colision(nombre_salida, nombres_zip_usados)
                    coincidencias_por_archivo[arcname] = resultado.coincidencias_por_nombre
                    inicio = time.perf_counter()
                    zf.write(ruta_salida, arcname=arcname)
                    fases["zip"] += time.perf_counter() - inicio
                except Exception as error:
                    app.logger.warning("Fallo procesando %s: %s", nombre_archivo, error)
                    errores_por_archivo[nombre_archivo] = str(error)

            lineas_errores = [f"{archivo}: {mensaje}" for archivo, mensaje in errores_por_archivo.items()]
            lineas_coincidencias = [
                f"{archivo} — \"{nombre}\": {conteo} coincidencia(s)"
                for archivo, coincidencias in coincidencias_por_archivo.items()
                for nombre, conteo in coincidencias.items()
            ]
            if lineas_errores or lineas_coincidencias:
                resumen_pdf = generar_pdf_resumen(
                    "Resumen del procesamiento (Modo Simple)",
                    [
                        ("Archivos que no se pudieron procesar", lineas_errores),
                        ("Coincidencias por archivo", lineas_coincidencias),
                    ],
                )
                zf.writestr("Resumen_Modo_Simple.pdf", resumen_pdf)
    finally:
        shutil.rmtree(carpeta_temporal, ignore_errors=True)

    total_archivos_ok = len(coincidencias_por_archivo)
    total_errores = len(errores_por_archivo)
    total_coincidencias = sum(
        conteo for coincidencias in coincidencias_por_archivo.values() for conteo in coincidencias.values()
    )
    app.logger.info(
        "modo simple: %d/%d archivos OK, %d coincidencia(s)",
        total_archivos_ok, total_archivos_ok + total_errores, total_coincidencias,
    )
    fases["total"] = time.perf_counter() - inicio_total
    tiempos_lote.setdefault("conteos", {})["nombres"] = len(nombres)
    _registrar_tiempos("simple", fases, tiempos_lote, bytes_subidos, ruta_zip.stat().st_size)

    respuesta = _enviar_zip_y_borrar(ruta_zip, "pdfs_resaltados.zip")
    respuesta.headers["X-Modo"] = "simple"
    respuesta.headers["X-Total-Archivos"] = str(total_archivos_ok)
    respuesta.headers["X-Total-Errores"] = str(total_errores)
    respuesta.headers["X-Total-Coincidencias"] = str(total_coincidencias)
    respuesta.headers["X-Paginas-OCR"] = str(paginas_leidas_con_ocr())
    respuesta.headers["X-Paginas-Escaneadas-Sin-Leer"] = str(paginas_escaneadas_sin_leer())
    return respuesta


def _procesar_modo_simple_recortado(nombres: list[str], archivos, formato: str, resaltar_filas: bool):
    """Modo por nombre con recorte: un solo PDF con los renglones donde
    aparece cada nombre (de todas las planillas), con su encabezado y su
    pie de póliza -ver resaltar_nombres_y_recortar. Con "Solo resaltar"
    marcado se usa _procesar_modo_simple (el PDF completo, solo resaltado)."""
    reiniciar_contador_ocr()
    errores_por_archivo: dict[str, str] = {}
    fases: dict[str, float] = {}
    bytes_subidos = 0
    inicio_total = time.perf_counter()
    carpeta_temporal = Path(tempfile.mkdtemp(prefix="resaltado_simple_"))
    carpeta_entrada = carpeta_temporal / "entrada"
    carpeta_entrada.mkdir(parents=True, exist_ok=True)

    try:
        rutas_entrada = []
        # "3_planilla.pdf" -> "planilla.pdf" (o "planilla (1).pdf" si se
        # subieron dos con el mismo nombre, para que no se pisen en el resumen)
        nombre_original: dict[str, str] = {}
        nombres_vistos: set[str] = set()
        inicio = time.perf_counter()
        for i, archivo in enumerate(archivos):
            nombre_archivo = _nombre_seguro(archivo.filename)
            if not _es_pdf(archivo):
                errores_por_archivo[nombre_archivo] = "No es un archivo PDF."
                continue
            ruta = carpeta_entrada / f"{i}_{nombre_archivo}"
            archivo.save(ruta)
            rutas_entrada.append(str(ruta))
            nombre_original[ruta.name] = _nombre_zip_sin_colision(nombre_archivo, nombres_vistos)
            bytes_subidos += ruta.stat().st_size
        fases["guardar_subidas"] = time.perf_counter() - inicio

        if not rutas_entrada:
            return jsonify(error="Ninguno de los archivos subidos es un PDF válido."), 400

        inicio = time.perf_counter()
        resultado = resaltar_nombres_y_recortar(
            rutas_entrada, nombres, str(carpeta_temporal / "salida"), formato, resaltar_filas=resaltar_filas,
        )
        fases["procesar"] = time.perf_counter() - inicio

        for archivo, mensaje in resultado["errores_por_archivo"].items():
            errores_por_archivo[nombre_original.get(archivo, archivo)] = mensaje
        coincidencias_por_archivo = {
            nombre_original.get(archivo, archivo): conteo
            for archivo, conteo in resultado["coincidencias_por_archivo"].items()
        }

        archivo_zip_temporal = tempfile.NamedTemporaryFile(
            suffix=".zip", prefix="resaltado_simple_zip_", delete=False
        )
        ruta_zip = Path(archivo_zip_temporal.name)
        archivo_zip_temporal.close()

        inicio = time.perf_counter()
        # compresslevel=1: el PDF ya sale compactado (ver _procesar_modo_cliente)
        with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
            if resultado["ruta_pdf"]:
                zf.write(resultado["ruta_pdf"], arcname=f"{_NOMBRE_PDF_RECORTADO}.pdf")
            lineas_errores = [f"{archivo}: {mensaje}" for archivo, mensaje in errores_por_archivo.items()]
            lineas_coincidencias = [
                f"{archivo} — \"{nombre}\": {conteo} renglón(es)"
                for archivo, coincidencias in coincidencias_por_archivo.items()
                for nombre, conteo in coincidencias.items()
            ]
            no_encontrados = [
                nombre for nombre in nombres
                if not any(c.get(nombre, 0) for c in coincidencias_por_archivo.values())
            ]
            zf.writestr("Resumen_Modo_Simple.pdf", generar_pdf_resumen(
                "Resumen del procesamiento (Modo por nombre)",
                [
                    ("Archivos que no se pudieron procesar", lineas_errores),
                    ("Nombres que no aparecen en ninguna planilla", no_encontrados),
                    ("Renglones encontrados por archivo", lineas_coincidencias),
                ],
            ))
        fases["zip"] = time.perf_counter() - inicio
    finally:
        shutil.rmtree(carpeta_temporal, ignore_errors=True)

    total_archivos_ok = len(coincidencias_por_archivo)
    total_coincidencias = sum(
        conteo for coincidencias in coincidencias_por_archivo.values() for conteo in coincidencias.values()
    )
    app.logger.info(
        "modo simple recortado: %d/%d archivos OK, %d renglón(es), %d nombre(s) sin encontrar",
        total_archivos_ok, total_archivos_ok + len(errores_por_archivo), total_coincidencias, len(no_encontrados),
    )
    fases["total"] = time.perf_counter() - inicio_total
    tiempos_lote = resultado.get("tiempos", {})
    tiempos_lote.setdefault("conteos", {}).update(nombres=len(nombres), pdfs=len(rutas_entrada))
    _registrar_tiempos("simple_recortado", fases, tiempos_lote, bytes_subidos, ruta_zip.stat().st_size)

    respuesta = _enviar_zip_y_borrar(ruta_zip, "nombres_recortados.zip")
    respuesta.headers["X-Modo"] = "simple_recortado"
    respuesta.headers["X-Total-Archivos"] = str(total_archivos_ok)
    respuesta.headers["X-Total-Errores"] = str(len(errores_por_archivo))
    respuesta.headers["X-Total-Coincidencias"] = str(total_coincidencias)
    respuesta.headers["X-Paginas-OCR"] = str(paginas_leidas_con_ocr())
    respuesta.headers["X-Paginas-Escaneadas-Sin-Leer"] = str(paginas_escaneadas_sin_leer())
    return respuesta


_NOMBRE_PDF_RECORTADO = "Nombres_encontrados"

_FORMATOS_VALIDOS = {"auto", "ccss", "mnk", "ins"}

_LIMITE_BYTES_NO_ENCONTRADOS_HEADER = 4000


def _procesar_modo_cliente(
    registros: list[dict], archivos, formato: str, resaltar_filas: bool = True, solo_resaltar: bool = False,
):
    reiniciar_contador_ocr()
    fases: dict[str, float] = {}
    inicio_total = time.perf_counter()
    carpeta_temporal = Path(tempfile.mkdtemp(prefix="resaltado_cliente_"))
    carpeta_entrada = carpeta_temporal / "entrada"
    carpeta_entrada.mkdir(parents=True, exist_ok=True)
    carpeta_salida = carpeta_temporal / "salida_por_cliente"

    try:
        rutas_entrada = []
        pdfs_invalidos = []
        bytes_subidos = 0
        inicio = time.perf_counter()
        for i, archivo in enumerate(archivos):
            nombre_archivo = _nombre_seguro(archivo.filename)
            if not _es_pdf(archivo):
                pdfs_invalidos.append(nombre_archivo)
                continue

            ruta = carpeta_entrada / f"{i}_{nombre_archivo}"
            archivo.save(ruta)
            rutas_entrada.append(str(ruta))
            bytes_subidos += ruta.stat().st_size
            _guardar_copia_debug(ruta, nombre_archivo)
        fases["guardar_subidas"] = time.perf_counter() - inicio

        if not rutas_entrada:
            return jsonify(error="Ninguno de los archivos subidos es un PDF válido."), 400

        inicio = time.perf_counter()
        try:
            if solo_resaltar:
                resultado = resaltar_por_cedula_sin_recortar(rutas_entrada, registros, str(carpeta_salida))
                archivos_generados = resultado["archivos_resaltados"]
            else:
                resultado = resaltar_por_cedula_y_exportar_por_cliente(
                    rutas_entrada, registros, str(carpeta_salida), formato, resaltar_filas=resaltar_filas
                )
                archivos_generados = resultado["archivos_por_cliente"]
        except Exception as error:
            app.logger.error("modo cliente: no se pudo procesar el lote: %s", error)
            return jsonify(error=f"No se pudieron procesar los PDFs: {error}"), 500
        fases["procesar"] = time.perf_counter() - inicio

        no_encontrados = [
            f"{r['nombre'] or 'sin nombre'} (cédula {r['cedula']}, cliente {r['cliente']})"
            for r in resultado["no_encontrados"]
        ]
        errores_archivos = [
            f"{archivo}: {mensaje}" for archivo, mensaje in resultado["errores_por_archivo"].items()
        ]
        errores_archivos.extend(f"{nombre}: no es un archivo PDF." for nombre in pdfs_invalidos)

        # zip a disco, no a un io.BytesIO() en RAM (ver _procesar_modo_simple)
        archivo_zip_temporal = tempfile.NamedTemporaryFile(
            suffix=".zip", prefix="resaltado_cliente_zip_", delete=False
        )
        ruta_zip = Path(archivo_zip_temporal.name)
        archivo_zip_temporal.close()

        # compresslevel=1: los PDFs ya salen comprimidos (_compactar_pdf), el
        # nivel por defecto gastaba ~6 s por lote sin achicar casi nada
        with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
            inicio = time.perf_counter()
            for _clave, ruta_pdf in sorted(archivos_generados.items()):
                zf.write(ruta_pdf, arcname=Path(ruta_pdf).name)
            fases["zip"] = time.perf_counter() - inicio
            inicio = time.perf_counter()
            if no_encontrados or errores_archivos:
                resumen_pdf = generar_pdf_resumen(
                    "Resumen del procesamiento",
                    [
                        ("Archivos que no se pudieron procesar", errores_archivos),
                        ("Cédulas no encontradas en ningún PDF", no_encontrados),
                    ],
                )
                zf.writestr("Resumen.pdf", resumen_pdf)

            resumen_excel = generar_excel_resumen(resultado["detalle_registros"])
            zf.writestr("Resumen_Facturacion.xlsx", resumen_excel)
            fases["resumenes"] = time.perf_counter() - inicio
    finally:
        shutil.rmtree(carpeta_temporal, ignore_errors=True)

    total_archivos_generados = len(archivos_generados)
    app.logger.info(
        "modo cliente (solo_resaltar=%s): %d PDF(s) generados, %d error(es), %d cédula(s) no encontradas",
        solo_resaltar, total_archivos_generados, len(errores_archivos), len(no_encontrados),
    )
    fases["total"] = time.perf_counter() - inicio_total
    tiempos_lote = resultado.get("tiempos", {})
    tiempos_lote.setdefault("conteos", {}).update(
        pdfs=len(rutas_entrada), registros=len(registros), pdfs_generados=total_archivos_generados,
    )
    _registrar_tiempos(
        "cliente_sin_recorte" if solo_resaltar else "cliente",
        fases, tiempos_lote, bytes_subidos, ruta_zip.stat().st_size,
    )

    respuesta = _enviar_zip_y_borrar(ruta_zip, "pdfs_resaltados_sin_recorte.zip" if solo_resaltar else "pdfs_por_cliente.zip")
    respuesta.headers["X-Modo"] = "cliente_sin_recorte" if solo_resaltar else "cliente"
    respuesta.headers["X-Total-Clientes"] = str(total_archivos_generados)
    respuesta.headers["X-Total-Errores"] = str(len(errores_archivos))
    respuesta.headers["X-Total-No-Encontrados"] = str(len(no_encontrados))
    respuesta.headers["X-Paginas-OCR"] = str(paginas_leidas_con_ocr())
    respuesta.headers["X-Paginas-Escaneadas-Sin-Leer"] = str(paginas_escaneadas_sin_leer())

    if no_encontrados:
        no_encontrados_b64 = base64.b64encode(json.dumps(no_encontrados).encode("utf-8")).decode("ascii")
        if len(no_encontrados_b64) <= _LIMITE_BYTES_NO_ENCONTRADOS_HEADER:
            respuesta.headers["X-No-Encontrados-B64"] = no_encontrados_b64

    return respuesta


@app.route("/api/procesar", methods=["POST"])
def procesar():
    archivo_excel = request.files.get("excel")
    archivos = request.files.getlist("pdfs")

    app.logger.info(
        "solicitud recibida: %d PDF(s), excel=%s",
        len(archivos), _nombre_seguro(archivo_excel.filename) if archivo_excel else "no",
    )

    if not archivos:
        return jsonify(error="Seleccione al menos un archivo PDF."), 400

    if archivo_excel and archivo_excel.filename and not archivo_excel.filename.lower().endswith((".xlsx", ".xlsm", ".csv")):
        return jsonify(error="El archivo de nombres debe ser un Excel (.xlsx / .xlsm) o un .csv."), 400

    formato = request.form.get("formato", "auto").strip().lower()
    if formato not in _FORMATOS_VALIDOS:
        formato = "auto"
    resaltar_param = request.form.get("resaltar", "true").strip().lower()
    resaltar_filas = resaltar_param in ("true", "1", "on", "yes")
    solo_resaltar_param = request.form.get("solo_resaltar", "false").strip().lower()
    solo_resaltar = solo_resaltar_param in ("true", "1", "on", "yes")

    try:
        registros = None
        if archivo_excel and archivo_excel.filename:
            registros = _registros_desde_excel(archivo_excel)
            archivo_excel.seek(0)

        if registros:
            with _turno_pymupdf():
                return _procesar_modo_cliente(
                    registros, archivos, formato, resaltar_filas=resaltar_filas, solo_resaltar=solo_resaltar,
                )

        texto_nombres = request.form.get("nombres", "").strip()
        nombres = _combinar_nombres(texto_nombres, archivo_excel)
    except ValueError as error:
        app.logger.warning("solicitud rechazada: %s", error)
        return jsonify(error=str(error)), 400

    if not nombres:
        return jsonify(error="Escriba al menos un nombre o suba un Excel con la lista de nombres."), 400

    with _turno_pymupdf():
        if solo_resaltar:
            return _procesar_modo_simple(nombres, archivos)
        return _procesar_modo_simple_recortado(nombres, archivos, formato, resaltar_filas)


if __name__ == "__main__":
    app.run()  # para depurar en la PC: flask --app api.index run --debug (nunca en el servidor)

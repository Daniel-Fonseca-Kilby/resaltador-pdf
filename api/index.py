"""
api/index.py
"""

import base64
import csv
import io
import json
import logging
import os
import shutil
import sys
import tempfile
import time
import unicodedata
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
from flask import Flask, after_this_request, jsonify, render_template, request, send_file
from werkzeug.exceptions import HTTPException

from resaltado_pdf import (
    generar_excel_resumen,
    generar_pdf_resumen,
    resaltar_nombres_en_pdf,
    resaltar_por_cedula_y_exportar_por_cliente,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024  # 60 MB, de sobra para una planilla

# TEMPORAL: apagado por defecto. Solo para diagnosticar el bug del pie de
# página de MNK con un archivo real -activar con DEBUG_GUARDAR_SUBIDAS=1
# en el entorno del servidor, y desactivar (borrando lo guardado) apenas
# se termine de diagnosticar.
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
    """Borra carpetas temporales huérfanas (prefijo 'resaltado_') más
    viejas que segundos_vida. Si una solicitud se cae a mitad de camino
    su carpeta queda sin borrar, y con el tiempo eso llena el disco de
    Render. Devuelve cuántas se borraron."""
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


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.errorhandler(Exception)
def _manejar_error(error):
    """Cualquier error sale como JSON en español -si no, Flask devuelve
    HTML y el front no lo puede leer."""
    if isinstance(error, HTTPException) and error.code == 413:
        return jsonify(error="El archivo es demasiado grande (máximo 60 MB en total)."), 413
    if isinstance(error, HTTPException):
        return jsonify(error=error.description or "Ocurrió un error."), error.code
    app.logger.exception("Error no manejado al procesar la solicitud")
    return jsonify(error="Ocurrió un error inesperado al procesar la solicitud."), 500


def _abrir_libro(archivo):
    """Si el Excel está dañado o no es realmente un .xlsx, tira un mensaje claro."""
    try:
        return openpyxl.load_workbook(archivo, read_only=True, data_only=True)
    except Exception:
        raise ValueError(
            f"El archivo Excel '{archivo.filename}' no se pudo leer. "
            "Verifique que sea un .xlsx válido y no esté dañado."
        )


def _decodificar_csv(contenido_bytes: bytes) -> str:
    """Prueba UTF-8 con BOM, UTF-8 normal y Latin-1 (Windows-1252), que es
    lo que suelen exportar Softland/SAP/Exactus en Costa Rica."""
    for codificacion in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return contenido_bytes.decode(codificacion)
        except UnicodeDecodeError:
            continue
    return contenido_bytes.decode("latin-1", errors="replace")


def _detectar_delimitador(texto: str) -> str:
    """Excel en español exporta CSV con punto y coma, porque la coma
    queda reservada para los decimales."""
    primera_linea = texto.splitlines()[0] if texto else ""
    if primera_linea.count(";") > primera_linea.count(","):
        return ";"
    return ","


def _filas_desde_archivo(archivo):
    """Generador de filas, venga el archivo en .xlsx/.xlsm o en .csv, para
    que el resto del código (sinónimos de columnas, extracción de
    cédula/nombre) no tenga que preocuparse por el formato de origen."""
    nombre = Path(archivo.filename).name.lower()
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
    """Primera columna no vacía de cada fila. Ignora la primera fila si
    parece encabezado (ej. 'Nombre')."""
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
    """Junta los nombres escritos a mano con los del Excel, sin
    duplicados (comparando en mayúsculas, pero conservando el primer
    formato con el que apareció cada uno)."""
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


# sinónimos normalizados (sin tildes, mayúsculas), en orden de prioridad:
# si el Excel trae varias columnas que calzan, gana la más específica de
# la lista, sin importar cuál esté más a la izquierda.
#
# "EMPRESA" no cuenta como sinónimo de cliente a propósito: en las
# planillas de VMA esa columna es una unidad interna (Comer, Servicios,
# etc.) que no tiene que ver con a quién se le factura.
_SINONIMOS_CEDULA = ["IDENTIFICACION", "CEDULA", "ID", "DOCUMENTO", "IDENTIFICACION FISCAL", "NUMERO"]
_SINONIMOS_CLIENTE = ["CLIENTE", "CUENTA"]
_SINONIMOS_NOMBRE = ["NOMBRE", "NOMBRES", "EMPLEADO", "COLABORADOR", "NOMBRE COMPLETO"]
# columna opcional: en la CCSS, un extranjero con DIMEX sale impreso en la
# planilla bajo su número de asegurado de la Caja, no bajo el DIMEX que
# trae el Excel -si esta columna viene, solo debe tener valor para
# extranjeros (vacía para nacionales, que se buscan por su cédula normal)
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
    """Si trae columna de cédula y de cliente (con sus sinónimos), arma
    un registro por fila para Modo Cliente. Si es de una sola columna, se
    asume lista de nombres y devuelve None para que el llamador use Modo
    Simple. Si tiene varias columnas pero ninguna calza, mejor un error
    claro que degradar en silencio -seguramente el usuario quería Modo
    Cliente y algo no calzó.

    La columna de número de asegurado es opcional -si no viene, no pasa
    nada; si viene, solo debería tener valor para extranjeros (la CCSS los
    imprime en la planilla bajo ese número, no bajo el DIMEX del Excel)."""
    filas = _filas_desde_archivo(archivo)

    encabezado = next(filas, None)
    if encabezado is None:
        return None

    indice_cedula = _indice_por_sinonimos(encabezado, _SINONIMOS_CEDULA)
    indice_cliente = _indice_por_sinonimos(encabezado, _SINONIMOS_CLIENTE)
    indice_nombre = _indice_por_sinonimos(encabezado, _SINONIMOS_NOMBRE)
    # opcional -si no viene la columna, indice_numero_asegurado queda en
    # None y ningún registro trae ese dato, sin romper nada
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
    """Genera el zip de PDFs resaltados y lo manda por streaming
    (send_file) en vez de base64 -con planillas pesadas eso duplicaba el
    archivo en memoria. El detalle por archivo/nombre va dentro del zip,
    en Resumen_Modo_Simple.pdf; la respuesta solo trae los totales, en
    cabeceras.

    El zip se arma en un archivo temporal en disco, no en un
    io.BytesIO() -con lotes grandes ese buffer se sumaba a la memoria que
    ya estaban usando los PDFs generados. Se borra apenas termina de
    enviarse la respuesta.
    """
    coincidencias_por_archivo: dict[str, dict] = {}
    errores_por_archivo: dict[str, str] = {}
    nombres_zip_usados: set[str] = set()
    carpeta_temporal = Path(tempfile.mkdtemp(prefix="resaltado_simple_"))
    archivo_zip_temporal = tempfile.NamedTemporaryFile(suffix=".zip", prefix="resaltado_simple_zip_", delete=False)
    ruta_zip = Path(archivo_zip_temporal.name)
    archivo_zip_temporal.close()

    try:
        with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, archivo in enumerate(archivos):
                nombre_archivo = Path(archivo.filename).name
                if not nombre_archivo.lower().endswith(".pdf"):
                    errores_por_archivo[nombre_archivo] = "No es un archivo PDF."
                    continue

                # prefijo por índice: es común subir dos PDFs con el mismo
                # nombre (de portales distintos) y sin esto se pisarían
                ruta_entrada = carpeta_temporal / f"{i}_{nombre_archivo}"
                archivo.save(ruta_entrada)

                nombre_salida = f"{Path(nombre_archivo).stem}_resaltado.pdf"
                ruta_salida = carpeta_temporal / f"{i}_{nombre_salida}"

                try:
                    resultado = resaltar_nombres_en_pdf(str(ruta_entrada), nombres, str(ruta_salida))
                    arcname = _nombre_zip_sin_colision(nombre_salida, nombres_zip_usados)
                    coincidencias_por_archivo[arcname] = resultado.coincidencias_por_nombre
                    zf.write(ruta_salida, arcname=arcname)
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

    @after_this_request
    def _borrar_zip_temporal(response):
        ruta_zip.unlink(missing_ok=True)
        return response

    respuesta = send_file(
        str(ruta_zip),
        mimetype="application/zip",
        as_attachment=True,
        download_name="pdfs_resaltados.zip",
    )
    respuesta.headers["X-Modo"] = "simple"
    respuesta.headers["X-Total-Archivos"] = str(total_archivos_ok)
    respuesta.headers["X-Total-Errores"] = str(total_errores)
    respuesta.headers["X-Total-Coincidencias"] = str(total_coincidencias)
    return respuesta


_FORMATOS_VALIDOS = {"auto", "ccss", "mnk", "ins"}

# margen bajo el límite típico de ~8 KB por cabecera HTTP -con un lote
# grande de no encontrados, mejor omitir la cabecera que arriesgarse a
# que el proxy rechace toda la respuesta
_LIMITE_BYTES_NO_ENCONTRADOS_HEADER = 4000


def _procesar_modo_cliente(registros: list[dict], archivos, formato: str, resaltar_filas: bool = True):
    carpeta_temporal = Path(tempfile.mkdtemp(prefix="resaltado_cliente_"))
    carpeta_entrada = carpeta_temporal / "entrada"
    carpeta_entrada.mkdir(parents=True, exist_ok=True)
    carpeta_salida = carpeta_temporal / "salida_por_cliente"

    try:
        rutas_entrada = []
        pdfs_invalidos = []
        for i, archivo in enumerate(archivos):
            nombre_archivo = Path(archivo.filename).name
            if not nombre_archivo.lower().endswith(".pdf"):
                pdfs_invalidos.append(nombre_archivo)
                continue
            # prefijo por índice: es común descargar "Planilla.pdf" de
            # varios portales (CCSS, INS...) con el mismo nombre
            ruta = carpeta_entrada / f"{i}_{nombre_archivo}"
            archivo.save(ruta)
            rutas_entrada.append(str(ruta))
            _guardar_copia_debug(ruta, nombre_archivo)

        if not rutas_entrada:
            return jsonify(error="Ninguno de los archivos subidos es un PDF válido."), 400

        try:
            resultado = resaltar_por_cedula_y_exportar_por_cliente(
                rutas_entrada, registros, str(carpeta_salida), formato, resaltar_filas=resaltar_filas
            )
        except Exception as error:
            app.logger.error("modo cliente: no se pudo procesar el lote: %s", error)
            return jsonify(error=f"No se pudieron procesar los PDFs: {error}"), 500

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

        with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for _cliente, ruta_pdf in sorted(resultado["archivos_por_cliente"].items()):
                zf.write(ruta_pdf, arcname=Path(ruta_pdf).name)
            if no_encontrados or errores_archivos:
                resumen_pdf = generar_pdf_resumen(
                    "Resumen del procesamiento",
                    [
                        ("Archivos que no se pudieron procesar", errores_archivos),
                        ("Cédulas no encontradas en ningún PDF", no_encontrados),
                    ],
                )
                zf.writestr("Resumen.pdf", resumen_pdf)

            # para que Facturación no tenga que abrir cada PDF a contar oficiales
            resumen_excel = generar_excel_resumen(resultado["detalle_registros"])
            zf.writestr("Resumen_Facturacion.xlsx", resumen_excel)
    finally:
        shutil.rmtree(carpeta_temporal, ignore_errors=True)

    total_clientes = len(resultado["archivos_por_cliente"])
    app.logger.info(
        "modo cliente: %d PDF(s) generados, %d error(es), %d cédula(s) no encontradas",
        total_clientes, len(errores_archivos), len(no_encontrados),
    )

    @after_this_request
    def _borrar_zip_temporal(response):
        ruta_zip.unlink(missing_ok=True)
        return response

    respuesta = send_file(
        str(ruta_zip),
        mimetype="application/zip",
        as_attachment=True,
        download_name="pdfs_por_cliente.zip",
    )
    respuesta.headers["X-Modo"] = "cliente"
    respuesta.headers["X-Total-Clientes"] = str(total_clientes)
    respuesta.headers["X-Total-Errores"] = str(len(errores_archivos))
    respuesta.headers["X-Total-No-Encontrados"] = str(len(no_encontrados))

    # para que el navegador liste las cédulas sin abrir el zip -si el lote
    # es grande y no cabe en una cabecera, mejor omitirla (igual queda en
    # Resumen.pdf, dentro del zip)
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
        len(archivos), archivo_excel.filename if archivo_excel else "no",
    )

    if not archivos:
        return jsonify(error="Seleccione al menos un archivo PDF."), 400

    if archivo_excel and archivo_excel.filename and not archivo_excel.filename.lower().endswith((".xlsx", ".xlsm", ".csv")):
        return jsonify(error="El archivo de nombres debe ser un Excel (.xlsx / .xlsm) o un .csv."), 400

    try:
        registros = None
        if archivo_excel and archivo_excel.filename:
            registros = _registros_desde_excel(archivo_excel)
            archivo_excel.seek(0)

        if registros:
            formato = request.form.get("formato", "auto").strip().lower()
            if formato not in _FORMATOS_VALIDOS:
                formato = "auto"
            resaltar_param = request.form.get("resaltar", "true").strip().lower()
            resaltar_filas = resaltar_param in ("true", "1", "on", "yes")
            return _procesar_modo_cliente(registros, archivos, formato, resaltar_filas=resaltar_filas)

        texto_nombres = request.form.get("nombres", "").strip()
        nombres = _combinar_nombres(texto_nombres, archivo_excel)
    except ValueError as error:
        app.logger.warning("solicitud rechazada: %s", error)
        return jsonify(error=str(error)), 400

    if not nombres:
        return jsonify(error="Escriba al menos un nombre o suba un Excel con la lista de nombres."), 400

    return _procesar_modo_simple(nombres, archivos)


if __name__ == "__main__":
    app.run(debug=True)

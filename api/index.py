"""
api/index.py
"""

import base64
import io
import json
import logging
import shutil
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import HTTPException

from resaltado_pdf import (
    generar_pdf_resumen,
    resaltar_nombres_en_pdf,
    resaltar_por_cedula_y_exportar_por_cliente,
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024  # 60 MB, de sobra para una planilla


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.errorhandler(Exception)
def _manejar_error(error):
    """Cualquier error sale como JSON en español (si no, Flask devuelve
    HTML y el front no puede leer la respuesta)."""
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


def _nombres_desde_excel(archivo) -> list[str]:
    """Lee la primera columna no vacía de cada fila de la primera hoja.
    Ignora la primera fila si parece un encabezado (ej. 'Nombre')."""
    libro = _abrir_libro(archivo)
    hoja = libro.active
    nombres = []
    for i, fila in enumerate(hoja.iter_rows(values_only=True)):
        valor = next((c for c in fila if c not in (None, "")), None)
        if valor is None:
            continue
        texto = str(valor).strip()
        if i == 0 and texto.upper() in ("NOMBRE", "NOMBRES", "NOMBRE COMPLETO"):
            continue
        if texto:
            nombres.append(texto)
    libro.close()
    return nombres


def _combinar_nombres(texto_nombres: str, archivo_excel) -> list[str]:
    """Une los nombres escritos a mano con los del Excel (si se subió uno),
    sin duplicados (comparando en mayúsculas, conservando el primer formato
    con el que apareció cada nombre)."""
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


# sinónimos normalizados (sin tildes, mayúsculas) que puede traer cada columna
_SINONIMOS_CEDULA = {"IDENTIFICACION", "CEDULA", "ID", "DOCUMENTO", "IDENTIFICACION FISCAL", "NUMERO"}
_SINONIMOS_CLIENTE = {"CLIENTE", "EMPRESA", "CUENTA", "PROYECTO", "PUESTO", "CONTRATO"}
_SINONIMOS_NOMBRE = {"NOMBRE", "NOMBRES", "EMPLEADO", "COLABORADOR", "NOMBRE COMPLETO"}


def _normalizar_encabezado(valor) -> str:
    texto = str(valor or "").strip().upper()
    sin_tildes = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in sin_tildes if not unicodedata.combining(c))


def _indice_por_sinonimos(encabezado, sinonimos: set) -> int | None:
    for i, valor in enumerate(encabezado):
        if valor and _normalizar_encabezado(valor) in sinonimos:
            return i
    return None


def _extraer_cedula_limpia(valor) -> str:
    """Extrae únicamente los dígitos de una cédula, eliminando decimales
    espurios que deja Excel cuando la columna quedó como celda numérica en
    vez de texto (ej. una BUSCARV trae 303370238.0 -> '303370238')."""
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
    """Si el Excel trae una columna de cédula y una de cliente (aceptando
    los sinónimos de _SINONIMOS_CEDULA/_SINONIMOS_CLIENTE), devuelve un
    registro por fila. Si el Excel es de una sola columna, se asume lista
    simple de nombres y devuelve None para que el llamador use el Modo
    Simple. Si trae varias columnas pero ninguna calza con lo esperado,
    mejor avisarle al usuario con un error claro que degradar en silencio
    a Modo Simple -probablemente quiso usar Modo Cliente y algo no calzó."""
    libro = _abrir_libro(archivo)
    hoja = libro.active
    filas = hoja.iter_rows(values_only=True)

    encabezado = next(filas, None)
    if encabezado is None:
        libro.close()
        return None

    indice_cedula = _indice_por_sinonimos(encabezado, _SINONIMOS_CEDULA)
    indice_cliente = _indice_por_sinonimos(encabezado, _SINONIMOS_CLIENTE)
    indice_nombre = _indice_por_sinonimos(encabezado, _SINONIMOS_NOMBRE)

    if indice_cedula is None or indice_cliente is None:
        columnas_con_datos = sum(1 for valor in encabezado if valor not in (None, ""))
        libro.close()
        if columnas_con_datos > 1:
            raise ValueError(
                "El Excel no contiene una columna de Identificación/Cédula y Cliente válida."
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
        registros.append(
            {
                "cedula": cedula,
                "cliente": cliente,
                "nombre": nombre,
            }
        )

    libro.close()
    return registros


@app.route("/api/detectar-modo-excel", methods=["POST"])
def detectar_modo_excel():
    """Preview rápido para el frontend: le dice al usuario qué modo se va
    a activar apenas elige el Excel, sin tener que subir los PDFs y
    esperar el procesamiento completo para enterarse."""
    archivo_excel = request.files.get("excel")
    if not archivo_excel or not archivo_excel.filename:
        return jsonify(error="No se recibió ningún archivo Excel."), 400
    if not archivo_excel.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify(error="El archivo debe ser un Excel (.xlsx o .xlsm)."), 400

    try:
        registros = _registros_desde_excel(archivo_excel)
    except ValueError as error:
        return jsonify(error=str(error)), 400

    if registros:
        return jsonify(modo="cliente", total_registros=len(registros))
    return jsonify(modo="simple")


def _nombre_zip_sin_colision(nombre: str, nombres_usados: set) -> str:
    """Si 'nombre' ya se usó en este zip (ej. dos PDFs de origen que se
    llamaban igual), le agrega un sufijo numérico -'Planilla_resaltado
    (1).pdf'- para no pisar la entrada anterior."""
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
    """Igual que el modo cliente: el zip se manda directo por streaming
    (send_file), sin pasar por base64 -eso duplicaba el archivo en memoria
    (bytes + texto) y con planillas pesadas era lo que más pegaba contra
    los 512 MB de RAM de Render. El detalle por archivo/nombre que antes
    se mandaba en el JSON ahora va en Resumen_Modo_Simple.pdf, adentro del
    zip; en la respuesta solo quedan los conteos, en cabeceras."""
    coincidencias_por_archivo: dict[str, dict] = {}
    errores_por_archivo: dict[str, str] = {}
    nombres_zip_usados: set[str] = set()
    buffer_zip = io.BytesIO()
    carpeta_temporal = Path(tempfile.mkdtemp(prefix="resaltado_simple_"))  # aislado por request, se borra al final

    try:
        with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, archivo in enumerate(archivos):
                nombre_archivo = Path(archivo.filename).name
                if not nombre_archivo.lower().endswith(".pdf"):
                    errores_por_archivo[nombre_archivo] = "No es un archivo PDF."
                    continue

                # prefijo por índice en disco -es común subir varios PDFs
                # con el mismo nombre (ej. descargados de portales distintos)
                # y sin esto el segundo pisaría al primero antes de procesarlo
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
        buffer_zip.seek(0)
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

    respuesta = send_file(
        buffer_zip,
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

# margen bajo el límite típico de ~8 KB por cabecera HTTP de la mayoría de
# servidores/proxies -con un lote grande de no encontrados es mejor omitir
# la cabecera que arriesgarse a que el proxy rechace toda la respuesta
_LIMITE_BYTES_NO_ENCONTRADOS_HEADER = 4000


def _procesar_modo_cliente(registros: list[dict], archivos, formato: str):
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
            # varios portales (CCSS, INS...) con el mismo nombre -sin esto
            # el segundo archivo pisaría al primero antes de procesar nada
            ruta = carpeta_entrada / f"{i}_{nombre_archivo}"
            archivo.save(ruta)
            rutas_entrada.append(str(ruta))

        if not rutas_entrada:
            return jsonify(error="Ninguno de los archivos subidos es un PDF válido."), 400

        try:
            resultado = resaltar_por_cedula_y_exportar_por_cliente(
                rutas_entrada, registros, str(carpeta_salida), formato
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

        buffer_zip = io.BytesIO()
        with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as zf:
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
        buffer_zip.seek(0)
    finally:
        shutil.rmtree(carpeta_temporal, ignore_errors=True)

    total_clientes = len(resultado["archivos_por_cliente"])
    app.logger.info(
        "modo cliente: %d PDF(s) generados, %d error(es), %d cédula(s) no encontradas",
        total_clientes, len(errores_archivos), len(no_encontrados),
    )

    respuesta = send_file(
        buffer_zip,
        mimetype="application/zip",
        as_attachment=True,
        download_name="pdfs_por_cliente.zip",
    )
    respuesta.headers["X-Modo"] = "cliente"
    respuesta.headers["X-Total-Clientes"] = str(total_clientes)
    respuesta.headers["X-Total-Errores"] = str(len(errores_archivos))
    respuesta.headers["X-Total-No-Encontrados"] = str(len(no_encontrados))

    # para que el navegador pueda listar las cédulas sin abrir el zip -pero
    # si el lote es grande y no cabe en una cabecera HTTP, mejor omitirla
    # (igual queda el detalle completo en Resumen.pdf, dentro del zip)
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

    if archivo_excel and archivo_excel.filename and not archivo_excel.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify(error="El archivo de nombres debe ser un Excel (.xlsx o .xlsm)."), 400

    try:
        registros = None
        if archivo_excel and archivo_excel.filename:
            registros = _registros_desde_excel(archivo_excel)
            archivo_excel.seek(0)

        if registros:
            formato = request.form.get("formato", "auto").strip().lower()
            if formato not in _FORMATOS_VALIDOS:
                formato = "auto"
            return _procesar_modo_cliente(registros, archivos, formato)

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

"""
Benchmark de rendimiento del resaltador (fase 0 del plan de rendimiento).

Mide cuánto tarda cada fase del procesamiento (el mismo desglose que queda
en el log del servidor como "tiempos <modo>: {...}") y cuánta memoria usa
cada modo, para decidir qué optimizar con números y no a ojo.

Cada modo corre en un proceso aparte, así el pico de memoria de uno no
contamina al siguiente.

Uso, en el servidor (idealmente cuando nadie esté usando la herramienta,
el benchmark compite por los mismos 2 vCPU):

    cd /opt/resaltador-pdf
    venv/bin/python deploy/benchmark_rendimiento.py
    venv/bin/python deploy/benchmark_rendimiento.py --pdfs 20 --paginas 30 --clientes 80
    venv/bin/python deploy/benchmark_rendimiento.py --reales /ruta/a/pdfs --excel /ruta/a/lista.xlsx
    venv/bin/python deploy/benchmark_rendimiento.py --modos cliente --perfil

Sin --reales se genera un corpus sintético estilo MNK (encabezado,
filas con cédula, renglón de total y leyenda "CODIFICACIÓN" en la última
página) con cédulas y nombres inventados.
"""

import argparse
import cProfile
import io
import os
import pstats
import random
import sys
import tempfile
import time
import zipfile
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pymupdf as fitz

from resaltado_pdf import (
    resaltar_nombres_en_pdf,
    resaltar_por_cedula_sin_recortar,
    resaltar_por_cedula_y_exportar_por_cliente,
)

_NOMBRES = [
    "JUAN", "MARIA", "JOSE", "ANA", "LUIS", "CARLOS", "SOFIA", "DANIEL", "LAURA", "PEDRO",
    "ANDREA", "JORGE", "PATRICIA", "MIGUEL", "GABRIELA", "FERNANDO", "VALERIA", "RICARDO",
]
_APELLIDOS = [
    "PEREZ", "MORA", "SOLANO", "ZUNIGA", "RAMIREZ", "ROJAS", "VARGAS", "JIMENEZ", "CASTRO",
    "CHAVES", "SEGURA", "QUESADA", "ARAYA", "BRENES", "CALDERON", "MONGE", "VEGA", "UMANA",
]

_ANCHO, _ALTO = 595, 842
_Y_TITULOS = 100
_Y_PRIMERA_FILA = 130
_ALTO_FILA = 26


def _escribir_fila(pagina, y, celdas, fontsize=9):
    x = 36
    for texto, ancho in celdas:
        pagina.insert_text((x, y), texto, fontsize=fontsize)
        x += ancho


def _escribir_encabezado(pagina, numero_poliza, numero_pagina):
    pagina.insert_text((36, 40), f"EMPRESA DE PRUEBA S.A. - POLIZA {numero_poliza:04d}", fontsize=12)
    pagina.insert_text((36, 60), f"Planilla de prueba - pagina {numero_pagina}", fontsize=9)
    _escribir_fila(pagina, _Y_TITULOS, [
        ("IDENTIFICACIÓN", 90), ("NOMBRE", 80), ("APELLIDOS", 130), ("SALARIO", 80), ("OBSERVACIÓN", 90),
    ])


def generar_corpus(carpeta: Path, n_pdfs: int, paginas: int, filas_por_pagina: int,
                   n_clientes: int, semilla: int) -> tuple[list[str], list[dict]]:
    """PDFs sintéticos + registros del "Excel" (cédula, cliente, nombre).
    ~70 % de los empleados de las planillas están en el Excel, y además
    hay un ~5 % de registros que no aparecen en ningún PDF (no encontrados)."""
    rng = random.Random(semilla)
    clientes = [f"Cliente {i:03d}" for i in range(n_clientes)]
    registros: list[dict] = []
    rutas: list[str] = []
    siguiente_cedula = 100000001
    max_filas = (_ALTO - 60 - _Y_PRIMERA_FILA) // _ALTO_FILA

    for numero_poliza in range(n_pdfs):
        documento = fitz.open()
        total_empleados = 0
        for numero_pagina in range(1, paginas + 1):
            pagina = documento.new_page(width=_ANCHO, height=_ALTO)
            _escribir_encabezado(pagina, numero_poliza, numero_pagina)
            es_ultima = numero_pagina == paginas
            filas = min(filas_por_pagina, max_filas - (5 if es_ultima else 0))
            y = _Y_PRIMERA_FILA
            for _ in range(filas):
                cedula = str(siguiente_cedula)
                siguiente_cedula += rng.randint(1, 50)
                nombre = rng.choice(_NOMBRES)
                apellidos = f"{rng.choice(_APELLIDOS)} {rng.choice(_APELLIDOS)}"
                _escribir_fila(pagina, y, [
                    (cedula, 90), (nombre, 80), (apellidos, 130), (f"{rng.randint(300, 900)}000", 80), ("Ninguna", 90),
                ])
                if rng.random() < 0.7:
                    registros.append({
                        "cedula": cedula, "cliente": rng.choice(clientes),
                        "nombre": f"{nombre} {apellidos}", "numero_asegurado": "",
                    })
                y += _ALTO_FILA
                total_empleados += 1
            if es_ultima:
                _escribir_fila(pagina, y + 20, [
                    (f"TOTAL DE TRABAJADORES: {total_empleados}", 220), ("TOTAL DE SALARIO: 999.999.999", 200),
                ])
                pagina.insert_text((36, y + 70), "CODIFICACIÓN: 01 Incapacidad  02 Vacaciones  03 Permiso", fontsize=8)
                pagina.insert_text((36, y + 90), "Firma del patrono: ______________________", fontsize=8)
        ruta = carpeta / f"poliza_{numero_poliza:04d}.pdf"
        documento.save(str(ruta), garbage=3, deflate=True)
        documento.close()
        rutas.append(str(ruta))

    for _ in range(max(1, len(registros) // 20)):
        registros.append({
            "cedula": str(900000000 + rng.randint(0, 99999999)), "cliente": rng.choice(clientes),
            "nombre": "NO EXISTE EN PLANILLAS", "numero_asegurado": "",
        })
    return rutas, registros


def cargar_reales(carpeta_pdfs: str, ruta_excel: str) -> tuple[list[str], list[dict]]:
    """Usa el mismo lector de Excel/CSV que el servidor, para que las
    columnas se detecten igual que en producción."""
    from werkzeug.datastructures import FileStorage

    from api.index import _registros_desde_excel

    rutas = sorted(str(p) for p in Path(carpeta_pdfs).glob("*.pdf"))
    if not rutas:
        raise SystemExit(f"No hay PDFs en {carpeta_pdfs}")
    with open(ruta_excel, "rb") as f:
        archivo = FileStorage(stream=io.BytesIO(f.read()), filename=Path(ruta_excel).name)
    registros = _registros_desde_excel(archivo)
    if not registros:
        raise SystemExit("El Excel no tiene columnas de Identificación/Cédula y Cliente reconocibles.")
    return rutas, registros


def _rss_pico_mb() -> float | None:
    try:
        import resource
    except ImportError:
        return None
    return round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024, 1)


def _sumar_tiempos(acumulado: dict, tiempos: dict) -> None:
    for seccion in ("segundos", "conteos"):
        destino = acumulado.setdefault(seccion, {})
        for clave, valor in tiempos.get(seccion, {}).items():
            destino[clave] = destino.get(clave, 0) + valor


def _medir_zip(rutas_salida: list[str]) -> dict:
    """Compara zip comprimido vs sin comprimir sobre los PDFs de salida
    (punto 10 del plan: los PDFs ya vienen comprimidos por dentro)."""
    resultado = {}
    for etiqueta, modo_zip in (("deflated", zipfile.ZIP_DEFLATED), ("stored", zipfile.ZIP_STORED)):
        buffer = tempfile.TemporaryFile()
        inicio = time.perf_counter()
        with zipfile.ZipFile(buffer, "w", modo_zip) as zf:
            for ruta in rutas_salida:
                zf.write(ruta, arcname=Path(ruta).name)
        resultado[etiqueta] = {
            "segundos": round(time.perf_counter() - inicio, 3),
            "mb": round(buffer.tell() / 1024 / 1024, 2),
        }
        buffer.close()
    return resultado


def _desglose_tamano(rutas_salida: list[str]) -> dict:
    """En qué se van los MB de los PDFs de salida: imágenes (el pie de
    póliza va como PNG), Form XObjects (lo que copia show_pdf_page) u otros
    streams (fuentes, contenido de página, apariencia de resaltados).
    "sin_comprimir" es el subconjunto de todo lo anterior que no trae Filter."""
    megas = {"imagenes": 0, "xobjects_form": 0, "otros_streams": 0, "sin_comprimir": 0}
    imagenes = 0
    for ruta in rutas_salida:
        documento = fitz.open(ruta)
        for xref in range(1, documento.xref_length()):
            if not documento.xref_is_stream(xref):
                continue
            tamano = len(documento.xref_stream_raw(xref) or b"")
            subtipo = documento.xref_get_key(xref, "Subtype")[1]
            if subtipo == "/Image":
                megas["imagenes"] += tamano
                imagenes += 1
            elif subtipo == "/Form":
                megas["xobjects_form"] += tamano
            else:
                megas["otros_streams"] += tamano
            if documento.xref_get_key(xref, "Filter")[0] == "null":
                megas["sin_comprimir"] += tamano
        documento.close()
    resultado = {clave: round(valor / 1024 / 1024, 2) for clave, valor in megas.items()}
    resultado["cantidad_imagenes"] = imagenes
    return resultado


def correr_modo(modo: str, rutas: list[str], registros: list[dict], nombres: list[str], perfil: bool) -> dict:
    """Corre en un proceso hijo (ver main)."""
    carpeta_salida = Path(tempfile.mkdtemp(prefix=f"benchmark_{modo}_"))
    perfilador = cProfile.Profile() if perfil else None
    inicio = time.perf_counter()
    if perfilador:
        perfilador.enable()

    if modo == "cliente":
        resultado = resaltar_por_cedula_y_exportar_por_cliente(rutas, registros, str(carpeta_salida), "mnk")
        tiempos = resultado["tiempos"]
        rutas_salida = list(resultado["archivos_por_cliente"].values())
        errores = resultado["errores_por_archivo"]
    elif modo == "sin_recorte":
        resultado = resaltar_por_cedula_sin_recortar(rutas, registros, str(carpeta_salida))
        tiempos = resultado["tiempos"]
        rutas_salida = list(resultado["archivos_resaltados"].values())
        errores = resultado["errores_por_archivo"]
    elif modo == "simple":
        tiempos, rutas_salida, errores = {}, [], {}
        for ruta in rutas:
            ruta_salida = str(carpeta_salida / f"{Path(ruta).stem}_resaltado.pdf")
            try:
                r = resaltar_nombres_en_pdf(ruta, nombres, ruta_salida)
                _sumar_tiempos(tiempos, r.tiempos)
                rutas_salida.append(ruta_salida)
            except Exception as error:
                errores[Path(ruta).name] = str(error)
    else:
        raise ValueError(modo)

    if perfilador:
        perfilador.disable()
    segundos_total = time.perf_counter() - inicio

    texto_perfil = None
    if perfilador:
        salida = io.StringIO()
        pstats.Stats(perfilador, stream=salida).sort_stats("cumulative").print_stats(25)
        texto_perfil = salida.getvalue()

    return {
        "modo": modo,
        "segundos_total": round(segundos_total, 3),
        "tiempos": tiempos,
        "rss_pico_mb": _rss_pico_mb(),
        "archivos_salida": len(rutas_salida),
        "mb_salida": round(sum(os.path.getsize(r) for r in rutas_salida) / 1024 / 1024, 2),
        "zip": _medir_zip(rutas_salida) if rutas_salida else None,
        "desglose_mb": _desglose_tamano(rutas_salida) if rutas_salida else None,
        "errores": errores,
        "perfil": texto_perfil,
    }


def imprimir(resultado: dict) -> None:
    total = resultado["segundos_total"]
    print(f"\n=== modo {resultado['modo']} ===")
    print(f"total: {total:.2f} s | pico de memoria: {resultado['rss_pico_mb']} MB | "
          f"{resultado['archivos_salida']} PDF(s) de salida, {resultado['mb_salida']} MB")
    segundos = resultado["tiempos"].get("segundos", {})
    for fase, s in sorted(segundos.items(), key=lambda par: -par[1]):
        print(f"  {fase:<24} {s:>8.2f} s  {100 * s / total if total else 0:5.1f} %")
    print("  (copiar_filas incluye show_pdf_page, resaltar_fila y reabrir_pdf_cliente; la suma pasa de 100 % por eso)")
    conteos = resultado["tiempos"].get("conteos", {})
    if conteos:
        print("  conteos: " + ", ".join(f"{k}={v}" for k, v in sorted(conteos.items())))
    if resultado["zip"]:
        z = resultado["zip"]
        print(f"  zip deflated: {z['deflated']['segundos']:.2f} s, {z['deflated']['mb']} MB | "
              f"zip stored: {z['stored']['segundos']:.2f} s, {z['stored']['mb']} MB")
    if resultado["desglose_mb"]:
        print("  tamaño de salida (MB): " + ", ".join(f"{k}={v}" for k, v in resultado["desglose_mb"].items()))
    if resultado["errores"]:
        print(f"  errores: {resultado['errores']}")
    if resultado["perfil"]:
        print("\n  --- cProfile (top 25 por tiempo acumulado) ---")
        print(resultado["perfil"])


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--pdfs", type=int, default=10, help="PDFs sintéticos (default 10)")
    parser.add_argument("--paginas", type=int, default=20, help="páginas por PDF sintético (default 20)")
    parser.add_argument("--filas", type=int, default=25, help="filas por página (default 25)")
    parser.add_argument("--clientes", type=int, default=40, help="clientes distintos en el Excel sintético (default 40)")
    parser.add_argument("--nombres", type=int, default=200, help="nombres a buscar en modo simple (default 200)")
    parser.add_argument("--semilla", type=int, default=42)
    parser.add_argument("--reales", help="carpeta con PDFs reales (en vez del corpus sintético)")
    parser.add_argument("--excel", help="Excel/CSV real con Identificación y Cliente (junto con --reales)")
    parser.add_argument("--modos", default="cliente,sin_recorte,simple")
    parser.add_argument("--perfil", action="store_true", help="agrega cProfile de cada modo")
    args = parser.parse_args()

    carpeta_corpus = Path(tempfile.mkdtemp(prefix="benchmark_corpus_"))
    if args.reales:
        if not args.excel:
            parser.error("--reales necesita también --excel")
        rutas, registros = cargar_reales(args.reales, args.excel)
        print(f"Corpus real: {len(rutas)} PDF(s), {len(registros)} registro(s) en el Excel")
    else:
        inicio = time.perf_counter()
        rutas, registros = generar_corpus(
            carpeta_corpus, args.pdfs, args.paginas, args.filas, args.clientes, args.semilla,
        )
        print(f"Corpus sintético: {len(rutas)} PDF(s) x {args.paginas} pág., {len(registros)} registro(s), "
              f"{args.clientes} cliente(s) -generado en {time.perf_counter() - inicio:.1f} s en {carpeta_corpus}")

    mb_entrada = sum(os.path.getsize(r) for r in rutas) / 1024 / 1024
    print(f"Entrada: {mb_entrada:.2f} MB | CPUs: {os.cpu_count()} | PyMuPDF {fitz.VersionBind}")

    nombres = list(dict.fromkeys(r["nombre"] for r in registros if r.get("nombre")))[: args.nombres]

    for modo in [m.strip() for m in args.modos.split(",") if m.strip()]:
        with ProcessPoolExecutor(max_workers=1) as ejecutor:
            resultado = ejecutor.submit(correr_modo, modo, rutas, registros, nombres, args.perfil).result()
        imprimir(resultado)

    print(f"\nLos PDFs de salida quedaron en {tempfile.gettempdir()}/benchmark_* -borrarlos al terminar.")


if __name__ == "__main__":
    main()

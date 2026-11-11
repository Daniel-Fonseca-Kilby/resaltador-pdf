"""Arranca Planillas en Windows con waitress (gunicorn, el de Linux, no corre en Windows).

    python -m servidor.arrancar --puerto 8001 --registros C:\\Planillas\\logs

Lo usa servidor/iniciar.ps1 (la tarea programada «Planillas VMA»). Deja los registros en un archivo por día (se
guardan 60). Un solo proceso con varios hilos: mientras un lote se procesa, los otros hilos siguen sirviendo la
página y la vista previa del Excel; PyMuPDF no es seguro entre hilos y api/index.py ya lo protege con su candado
(_turno_pymupdf), así que los lotes van de uno en uno.
"""
import argparse
import logging
import logging.handlers
import os
from pathlib import Path

# El OCR de planillas escaneadas usa el Tesseract portable empaquetado en <app>/tesseract. Se fija acá, antes de
# importar la app, para que resaltado_pdf lo encuentre: TESSERACT_EXE (el binario) y TESSDATA_PREFIX (el idioma).
# setdefault respeta valores puestos a mano (por ejemplo en las pruebas de integración).
_TESSERACT = Path(__file__).resolve().parent.parent / "tesseract"
os.environ.setdefault("TESSERACT_EXE", str(_TESSERACT / "tesseract.exe"))
os.environ.setdefault("TESSDATA_PREFIX", str(_TESSERACT / "tessdata"))


def _hilos_por_defecto() -> int:
    """Hilos de waitress: lo que diga PLANILLAS_HILOS en el .env (iniciar.ps1 lo carga al entorno), si no 8. Más hilos
    mantienen la página, la vista previa y /salud respondiendo aunque un lote pesado esté corriendo -los lotes igual van
    de a uno por el candado de PyMuPDF, así que esto da RESPUESTA, no más lotes a la vez. En JUPITER (20 hilos) subir a
    12-16 no molesta; en una PC chica, 8 sobra."""
    try:
        hilos = int(os.environ.get("PLANILLAS_HILOS", "") or 8)
    except ValueError:
        hilos = 8
    return max(1, hilos)


def configurar_registros(carpeta: Path) -> None:
    """Antes de importar la app: su logging.basicConfig no hace nada si el registro ya está configurado."""
    carpeta.mkdir(parents=True, exist_ok=True)
    archivo = logging.handlers.TimedRotatingFileHandler(
        carpeta / "app.log", when="midnight", backupCount=60, encoding="utf-8")
    archivo.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    raiz = logging.getLogger()
    raiz.handlers[:] = [archivo]
    raiz.setLevel(logging.INFO)


def main() -> None:
    parser = argparse.ArgumentParser(description="Arranca Planillas con waitress.")
    parser.add_argument("--puerto", type=int, default=8001)
    parser.add_argument("--registros", type=Path, default=Path("..") / "logs")
    parser.add_argument("--hilos", type=int, default=_hilos_por_defecto())
    args = parser.parse_args()
    configurar_registros(args.registros)

    from waitress import serve
    from api.index import app

    logging.getLogger("planillas").info("Arranca Planillas en el puerto %s con %s hilos", args.puerto, args.hilos)
    # Todas las interfaces, porque el equipo entra desde la red. El firewall (instalar.ps1) solo deja la red interna.
    # channel_timeout: una subida lenta de PDFs grandes no se corta (gunicorn tenía --timeout 600).
    serve(app, host="0.0.0.0", port=args.puerto, threads=args.hilos, channel_timeout=600, ident="Planillas")  # nosec B104


if __name__ == "__main__":
    main()

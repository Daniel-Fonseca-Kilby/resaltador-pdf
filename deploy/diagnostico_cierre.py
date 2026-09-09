"""Diagnostico: para las ultimas paginas de un PDF real, muestra donde caen
las filas de empleado (ancla posicional) y que tan alto es el hueco entre la
ultima fila y el final de esa pagina, para poder ver exactamente que esta
agarrando el respaldo posicional de _pixmaps_pie_de_poliza. No depende de
search_for -por eso sirve incluso cuando el texto del total/codificacion no
es buscable en el PDF real.

Uso: venv/bin/python deploy/diagnostico_cierre.py /tmp/debug_subidas/Electro_MNK_082026.pdf
"""
import sys

sys.path.insert(0, "/opt/resaltador-pdf")

import pymupdf as fitz
from resaltado_pdf import _y0s_anclas_fila


def main():
    ruta = sys.argv[1]
    documento = fitz.open(ruta)
    total_paginas = documento.page_count
    print(f"PAGINAS TOTALES: {total_paginas}")

    n = min(8, total_paginas)
    print(f"\n=== Revisando las ultimas {n} paginas ===")
    for indice in range(total_paginas - n, total_paginas):
        pagina = documento[indice]
        y0s = _y0s_anclas_fila(pagina)
        alto_pagina = pagina.rect.height
        texto_len = len(pagina.get_text().strip())
        print(f"\n--- Pagina {indice + 1} (indice {indice}) ---")
        print(f"  alto de hoja: {alto_pagina:.1f}pt")
        print(f"  cantidad de texto extraible: {texto_len} caracteres")
        print(f"  anclas de fila (cedula) encontradas: {len(y0s)}")
        if y0s:
            print(f"  primera ancla y0={y0s[0]:.1f}  ultima ancla y0={y0s[-1]:.1f}")
            hueco = alto_pagina - y0s[-1]
            print(f"  hueco desde la ultima ancla hasta el fondo de la hoja: {hueco:.1f}pt")
        else:
            print("  (sin anclas de fila -> pagina sin empleados, seria 'pagina de cierre siguiente')")

        # texto crudo de la pagina, para verlo aunque search_for no encuentre
        # las frases exactas -a veces si aparece por get_text() aunque no por
        # search_for
        texto = pagina.get_text().strip()
        primeras_lineas = "\n    ".join(texto.splitlines()[:15])
        print(f"  primeras lineas de texto crudo:\n    {primeras_lineas}")

    documento.close()


if __name__ == "__main__":
    main()

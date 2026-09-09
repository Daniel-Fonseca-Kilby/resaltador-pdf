"""Diagnostico: corre las funciones REALES de recorte del pie de pagina
(las mismas que usa _pixmaps_pie_de_poliza) contra la ultima pagina de un
PDF real, e imprime los rectangulos exactos que calculan, mas el texto de
esa zona agrupado por renglon (una linea de salida por renglon del PDF,
no una por palabra) para que quepa en una sola pantalla sin necesitar
scroll.

Uso: venv/bin/python deploy/diagnostico_franjas.py /tmp/debug_subidas/Servicios_MNK_082026.pdf
"""
import sys

sys.path.insert(0, "/opt/resaltador-pdf")

import pymupdf as fitz
from resaltado_pdf import (
    _franja_leyenda_en_pagina,
    _franja_total_en_pagina,
    _recortar_leyenda_tras_total,
    _y0s_anclas_fila,
)


def agrupar_por_renglon(palabras, tolerancia=3):
    """Junta palabras cuyo y0 cae dentro de 'tolerancia' entre si, y arma
    una sola linea de texto por renglon (ordenadas por x)."""
    ordenadas = sorted(palabras, key=lambda w: (w[1], w[0]))
    renglones = []
    actual = []
    y_actual = None
    for w in ordenadas:
        if y_actual is None or abs(w[1] - y_actual) <= tolerancia:
            actual.append(w)
            y_actual = w[1] if y_actual is None else y_actual
        else:
            renglones.append(actual)
            actual = [w]
            y_actual = w[1]
    if actual:
        renglones.append(actual)
    return renglones


def main():
    ruta = sys.argv[1]
    documento = fitz.open(ruta)
    ultimo_indice = documento.page_count - 1
    num_a_revisar = min(6, documento.page_count)

    print(f"PAGINAS: {documento.page_count}")

    pagina_total = franja_total = None
    for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
        candidata = documento[indice]
        franja = _franja_total_en_pagina(candidata, "mnk")
        if franja is not None:
            pagina_total, franja_total = candidata, franja
            print(f"TOTAL en pag {indice + 1}: y0={franja.y0:.1f} y1={franja.y1:.1f}")
            break
    if pagina_total is None:
        print("TOTAL: no encontrada")

    pagina_leyenda = franja_leyenda = None
    for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
        candidata = documento[indice]
        franja = _franja_leyenda_en_pagina(candidata, "mnk")
        if franja is not None:
            pagina_leyenda, franja_leyenda = candidata, franja
            print(f"LEYENDA cruda en pag {indice + 1}: y0={franja.y0:.1f} y1={franja.y1:.1f}")
            break
    if pagina_leyenda is None:
        print("LEYENDA: no encontrada")

    if franja_leyenda is not None:
        misma = pagina_leyenda is pagina_total
        recortada = _recortar_leyenda_tras_total(franja_leyenda, franja_total, misma)
        print(f"misma_pagina={misma}  LEYENDA recortada: y0={recortada.y0:.1f} y1={recortada.y1:.1f}")

    if pagina_total is not None:
        y0s = _y0s_anclas_fila(pagina_total)
        print(f"ultima fila (ancla) en esa pagina: y0={y0s[-1] if y0s else None}")
        print(f"alto de la hoja: {pagina_total.rect.height:.1f}")

        print("\n--- renglones desde y=180 hasta el fondo (uno por linea) ---")
        palabras = [w for w in pagina_total.get_text("words") if w[1] > 180]
        for renglon in agrupar_por_renglon(palabras):
            y0 = renglon[0][1]
            texto = " ".join(w[4] for w in sorted(renglon, key=lambda w: w[0]))
            print(f"y0={y0:6.1f}  {texto[:100]}")

    documento.close()


if __name__ == "__main__":
    main()

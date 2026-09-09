"""Diagnostico: corre las funciones REALES de recorte del pie de pagina
(las mismas que usa _pixmaps_pie_de_poliza) contra la ultima pagina de un
PDF real, e imprime los rectangulos exactos que calculan -para no tener
que hacer la aritmetica a mano y arriesgarse a un error de calculo.

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


def main():
    ruta = sys.argv[1]
    documento = fitz.open(ruta)
    ultimo_indice = documento.page_count - 1
    num_a_revisar = min(6, documento.page_count)

    print(f"PAGINAS TOTALES: {documento.page_count}")

    pagina_total = franja_total = None
    for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
        candidata = documento[indice]
        franja = _franja_total_en_pagina(candidata, "mnk")
        if franja is not None:
            pagina_total, franja_total = candidata, franja
            print(f"franja_total encontrada en pagina {indice + 1}: {franja}")
            break
    if pagina_total is None:
        print("franja_total: NO encontrada en ninguna de las ultimas paginas")

    pagina_leyenda = franja_leyenda = None
    for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
        candidata = documento[indice]
        franja = _franja_leyenda_en_pagina(candidata, "mnk")
        if franja is not None:
            pagina_leyenda, franja_leyenda = candidata, franja
            print(f"franja_leyenda (antes de recortar) encontrada en pagina {indice + 1}: {franja}")
            break
    if pagina_leyenda is None:
        print("franja_leyenda: NO encontrada en ninguna de las ultimas paginas")

    if franja_leyenda is not None:
        misma_pagina = pagina_leyenda is pagina_total
        franja_leyenda_recortada = _recortar_leyenda_tras_total(franja_leyenda, franja_total, misma_pagina)
        print(f"misma_pagina (total y leyenda): {misma_pagina}")
        print(f"franja_leyenda (recortada): {franja_leyenda_recortada}")

    if pagina_total is not None:
        y0s = _y0s_anclas_fila(pagina_total)
        print(f"\nanclas de fila en la pagina del total: {y0s}")
        print(f"alto de esa hoja: {pagina_total.rect.height}")

    # texto de la pagina donde cae el total, para ver que hay exactamente
    # entre el final de franja_total y el inicio de franja_leyenda recortada
    if pagina_total is not None:
        print(f"\n--- texto completo de la pagina {pagina_total.number + 1} con posiciones Y ---")
        for w in pagina_total.get_text("words"):
            if w[1] > 180:  # solo lo que esta cerca o debajo del total, para no saturar
                print(f"  y0={w[1]:.1f} y1={w[3]:.1f} x0={w[0]:.1f}  '{w[4]}'")

    documento.close()


if __name__ == "__main__":
    main()

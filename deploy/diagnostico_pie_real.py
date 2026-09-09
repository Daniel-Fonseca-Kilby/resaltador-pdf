"""Reproduce EXACTAMENTE la logica de _pixmaps_pie_de_poliza (la funcion
real que usa la app en produccion, no una copia aparte) contra un PDF real,
e imprime cada paso -que pagina se eligio para el total, cual para la
leyenda, como se recortan y extienden entre si, y el tamano final de cada
recorte. Ademas guarda cada recorte como PNG en /tmp para verlo tal cual
saldria en el PDF del cliente.

Uso: venv/bin/python deploy/diagnostico_pie_real.py /tmp/debug_subidas/archivo.pdf [formato]
"""
import sys

sys.path.insert(0, "/opt/resaltador-pdf")

import pymupdf as fitz

from resaltado_pdf import (
    _extender_leyenda_para_incluir_total,
    _franja_leyenda_en_pagina,
    _franja_total_en_pagina,
    _recortar_total_antes_de_leyenda,
)


def main():
    ruta = sys.argv[1]
    formato = sys.argv[2] if len(sys.argv) > 2 else "mnk"

    documento = fitz.open(ruta)
    ultimo_indice = documento.page_count - 1
    num_a_revisar = min(6, documento.page_count)
    print(f"PAGINAS TOTALES: {documento.page_count} -revisando las ultimas {num_a_revisar}")

    pagina_total = franja_total = None
    for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
        candidata = documento[indice]
        franja = _franja_total_en_pagina(candidata, formato)
        if franja is not None:
            pagina_total, franja_total = candidata, franja
            break
    if pagina_total is not None:
        print(f"TOTAL cruda: pagina {pagina_total.number + 1} (indice {pagina_total.number})  franja={franja_total}  alto={franja_total.height:.1f}")
    else:
        print("TOTAL: no se encontro en ninguna de las ultimas paginas")

    pagina_leyenda = franja_leyenda = None
    for indice in range(ultimo_indice, ultimo_indice - num_a_revisar, -1):
        candidata = documento[indice]
        franja = _franja_leyenda_en_pagina(candidata, formato)
        if franja is not None:
            pagina_leyenda, franja_leyenda = candidata, franja
            break
    if pagina_leyenda is not None:
        print(f"LEYENDA cruda: pagina {pagina_leyenda.number + 1} (indice {pagina_leyenda.number})  franja={franja_leyenda}  alto={franja_leyenda.height:.1f}")
    else:
        print("LEYENDA: no se encontro en ninguna de las ultimas paginas")

    misma_pagina = pagina_total is not None and pagina_leyenda is not None and pagina_leyenda.number == pagina_total.number
    print(f"misma_pagina (comparando .number) = {misma_pagina}")

    franja_total_original = franja_total
    if pagina_total is not None:
        franja_total = _recortar_total_antes_de_leyenda(franja_total, franja_leyenda, misma_pagina)
        print(f"TOTAL recortada (para no tragarse la leyenda): {franja_total}  alto={franja_total.height:.1f}")
        for widget in pagina_total.widgets() or []:
            try:
                widget.update()
            except Exception:
                pass
        pixmap_total = pagina_total.get_pixmap(clip=franja_total, matrix=fitz.Matrix(2, 2))
        pixmap_total.save("/tmp/diagnostico_total.png")
        print(f"Guardado /tmp/diagnostico_total.png  ({pixmap_total.width}x{pixmap_total.height}px)")

    if pagina_leyenda is not None:
        franja_leyenda_extendida = _extender_leyenda_para_incluir_total(franja_leyenda, franja_total_original, misma_pagina)
        print(f"LEYENDA extendida (para incluir el total completo, aunque se duplique): {franja_leyenda_extendida}  alto={franja_leyenda_extendida.height:.1f}")
        for widget in pagina_leyenda.widgets() or []:
            try:
                widget.update()
            except Exception:
                pass
        pixmap = pagina_leyenda.get_pixmap(clip=franja_leyenda_extendida, matrix=fitz.Matrix(2, 2))
        pixmap.save("/tmp/diagnostico_leyenda.png")
        print(f"Guardado /tmp/diagnostico_leyenda.png  ({pixmap.width}x{pixmap.height}px)")

    print(f"\nAlto real de la hoja donde esta la leyenda: {(pagina_leyenda or pagina_total).rect.height:.1f}pt")
    print(f"Ancho real de esa hoja: {(pagina_leyenda or pagina_total).rect.width:.1f}pt")

    documento.close()


if __name__ == "__main__":
    main()

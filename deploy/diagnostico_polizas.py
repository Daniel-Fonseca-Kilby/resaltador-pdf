
import re
import sys

sys.path.insert(0, "/opt/resaltador-pdf")

import pymupdf as fitz

PATRON_POLIZA = re.compile(r"#\s*DE\s*P[OÓ]LIZA\s*([A-Z0-9\-]+)", re.IGNORECASE)


def main():
    ruta = sys.argv[1]
    documento = fitz.open(ruta)
    total_paginas = documento.page_count
    print(f"PAGINAS TOTALES: {total_paginas}")

    poliza_actual = None
    apariciones = []
    for indice, pagina in enumerate(documento):
        texto = pagina.get_text()
        m = PATRON_POLIZA.search(texto)
        if m:
            poliza = m.group(1)
            if poliza != poliza_actual:
                apariciones.append((indice, poliza))
                poliza_actual = poliza

    print(f"\nCantidad de polizas distintas detectadas: {len(apariciones)}")
    for indice, poliza in apariciones:
        print(f"  empieza en pagina {indice + 1} (indice {indice}): {poliza}")

    documento.close()


if __name__ == "__main__":
    main()

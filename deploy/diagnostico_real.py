"""Diagnostico usando las funciones reales del proyecto contra un PDF real.
Uso: venv/bin/python deploy/diagnostico_real.py /tmp/diagnostico.pdf
"""
import sys
sys.path.insert(0, "/opt/resaltador-pdf")

import pymupdf as fitz
from resaltado_pdf import (
    _techo_de_datos,
    _franja_total_en_pagina,
    _franja_leyenda_en_pagina,
    _PERFILES_ENCABEZADO,
    _PERFILES_PIE_PAGINA,
    _PERFILES_LEYENDA_PIE,
)


def main():
    ruta = sys.argv[1]
    documento = fitz.open(ruta)
    print(f"PAGINAS: {documento.page_count}")

    for indice, pagina in enumerate(documento):
        print(f"\n=== Pagina {indice + 1} ({pagina.rect.width:.0f}x{pagina.rect.height:.0f}pt) ===")

        techo = _techo_de_datos(pagina, "mnk")
        print(f"_techo_de_datos(formato=mnk) = {techo}")

        franja_total = _franja_total_en_pagina(pagina, "mnk")
        print(f"_franja_total_en_pagina = {franja_total}")

        franja_leyenda = _franja_leyenda_en_pagina(pagina, "mnk")
        print(f"_franja_leyenda_en_pagina = {franja_leyenda}")

        # cuantas veces aparece cada ancla de encabezado (por si se repite
        # varias veces dentro de la misma pagina gigante)
        for nombre_perfil, anclas in _PERFILES_ENCABEZADO.items():
            for ancla in anclas:
                n = len(pagina.search_for(ancla))
                if n:
                    print(f"  ancla encabezado '{ancla}' ({nombre_perfil}) aparece {n} vez(es)")

        widgets = list(pagina.widgets() or [])
        print(f"cantidad de widgets (campos de formulario) en esta pagina: {len(widgets)}")
        for w in widgets[:5]:
            print(f"  widget: field_name={w.field_name!r} value={w.field_value!r} rect={w.rect}")

    documento.close()


if __name__ == "__main__":
    main()

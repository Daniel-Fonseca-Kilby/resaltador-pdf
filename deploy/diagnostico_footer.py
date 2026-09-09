"""Diagnostico: imprime, para las ultimas paginas de un PDF, donde
aparecen exactamente los anclajes del pie de pagina (total y leyenda) y
las cedulas encontradas cerca de una fila de interes. Uso:
    venv/bin/python deploy/diagnostico_footer.py /tmp/diagnostico.pdf
"""
import sys

import pymupdf as fitz

ANCLAS_TOTAL = ["TOTAL DE TRABAJADORES", "TOTAL DE SALARIO"]
ANCLA_LEYENDA = "CODIFICACIÓN"


def normalizar(texto):
    import unicodedata
    texto = texto.strip().upper()
    sin_tildes = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in sin_tildes if not unicodedata.combining(c))


def main():
    ruta = sys.argv[1]
    documento = fitz.open(ruta)
    print(f"PAGINAS TOTALES: {documento.page_count}")

    n = min(5, documento.page_count)
    print(f"\n=== Revisando las ultimas {n} paginas ===")
    for indice in range(documento.page_count - n, documento.page_count):
        pagina = documento[indice]
        print(f"\n--- Pagina {indice + 1} (indice {indice}) ---")
        for ancla in ANCLAS_TOTAL:
            rects = pagina.search_for(ancla, quads=False)
            if not rects:
                rects = pagina.search_for(normalizar(ancla), quads=False)
            if rects:
                for r in rects:
                    print(f"  '{ancla}' encontrado en y0={r.y0:.1f} y1={r.y1:.1f}")
            else:
                print(f"  '{ancla}' NO aparece")
        rects_leyenda = pagina.search_for(ANCLA_LEYENDA, quads=False)
        if not rects_leyenda:
            rects_leyenda = pagina.search_for(normalizar(ANCLA_LEYENDA), quads=False)
        if rects_leyenda:
            for r in rects_leyenda:
                print(f"  '{ANCLA_LEYENDA}' encontrado en y0={r.y0:.1f} y1={r.y1:.1f}")
        else:
            print(f"  '{ANCLA_LEYENDA}' NO aparece")

    # busca la fila de Jehudy (cedula 701810913) y lista todas las
    # palabras cercanas en Y, para ver que se esta colando
    print("\n=== Buscando fila con cedula 701810913 (Jehudy) ===")
    for indice, pagina in enumerate(documento):
        rects = pagina.search_for("701810913")
        if not rects:
            continue
        for r in rects:
            print(f"\nEncontrado en pagina {indice + 1}, y0={r.y0:.1f}")
            palabras = pagina.get_text("words")
            cercanas = sorted(
                (w for w in palabras if abs(w[1] - r.y0) <= 20),
                key=lambda w: (w[1], w[0]),
            )
            print("Palabras dentro de 20pt en Y (ordenadas por fila, luego columna):")
            for w in cercanas:
                print(f"  y0={w[1]:.1f} y1={w[3]:.1f} x0={w[0]:.1f}  '{w[4]}'")

    documento.close()


if __name__ == "__main__":
    main()

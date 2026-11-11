"""Precisión en escaneados: cuando el OCR confunde la cédula de alguien con la de
OTRA persona (visto en escaneos densos), el cruce con el nombre del Excel evita
un resaltado -y una facturación- equivocados. El cruce SOLO corre en archivos
escaneados (verificar_nombre=True); en PDFs digitales el match exacto manda y no
se toca nada.
"""
import pymupdf as fitz

from resaltado_pdf import (
    _construir_mapas_cedula,
    _localizar_coincidencias_de_cedula,
    _nombre_de_fila_coincide,
    _recuperar_cedula_por_cercania,
    _y0s_anclas_fila,
)

AN, AL = 595, 842


def _pagina_una_fila(cedula, nombre):
    doc = fitz.open()
    p = doc.new_page(width=AN, height=AL)
    p.insert_text((40, 120), cedula, fontsize=9)
    p.insert_text((160, 120), nombre, fontsize=9)
    return doc, p


# ---- _nombre_de_fila_coincide -------------------------------------------------

def test_nombre_presente_en_la_fila_coincide():
    _doc, p = _pagina_una_fila("304560789", "CARLOS RAMIREZ ROJAS")
    palabras = p.get_text("words")
    y0 = palabras[0][1]
    assert _nombre_de_fila_coincide(palabras, y0, "Carlos Ramirez Rojas") is True


def test_nombre_distinto_en_la_fila_no_coincide():
    _doc, p = _pagina_una_fila("506780001", "LUIS CASTRO CHAVES")
    palabras = p.get_text("words")
    y0 = palabras[0][1]
    # el Excel esperaba a RICARDO en esa cédula; la fila trae a LUIS -> no coincide
    assert _nombre_de_fila_coincide(palabras, y0, "Ricardo Quesada Lopez") is False


def test_sin_nombre_util_no_filtra():
    _doc, p = _pagina_una_fila("506780001", "LUIS CASTRO CHAVES")
    palabras = p.get_text("words")
    y0 = palabras[0][1]
    assert _nombre_de_fila_coincide(palabras, y0, "") is True          # Excel sin nombre
    assert _nombre_de_fila_coincide(palabras, y0, "Luis") is True       # una sola palabra: no alcanza para filtrar


def test_fila_sin_nombre_claro_no_filtra():
    _doc, p = _pagina_una_fila("506780001", "")   # solo la cédula, sin nombre en la fila
    palabras = p.get_text("words")
    y0 = palabras[0][1]
    assert _nombre_de_fila_coincide(palabras, y0, "Ricardo Quesada Lopez") is True


def test_tolera_la_ene_rota():
    _doc, p = _pagina_una_fila("222222222", "MARIA MU OZ ARAYA")  # MUÑOZ con Ñ rota (espacio)
    palabras = p.get_text("words")
    y0 = palabras[0][1]
    assert _nombre_de_fila_coincide(palabras, y0, "Maria Munoz Araya") is True


# ---- cruce dentro del localizador (gated por verificar_nombre) ----------------

def _pagina_dos_filas():
    doc = fitz.open()
    p = doc.new_page(width=AN, height=AL)
    # fila A: cédula de RICARDO pero nombre de LUIS (el OCR confundió la cédula)
    p.insert_text((40, 120), "506780001", fontsize=9)
    p.insert_text((160, 120), "LUIS CASTRO CHAVES", fontsize=9)
    # fila B: CARLOS correcto
    p.insert_text((40, 160), "304560789", fontsize=9)
    p.insert_text((160, 160), "CARLOS RAMIREZ ROJAS", fontsize=9)
    return doc, p


def _mapas():
    registros = [
        {"cedula": "506780001", "cliente": "VMA", "nombre": "Ricardo Quesada Lopez"},
        {"cedula": "304560789", "cliente": "VMA", "nombre": "Carlos Ramirez Rojas"},
    ]
    registros_unicos, mapa_cedulas, mapa_id = _construir_mapas_cedula(registros)
    mapa_nombres = {}
    for (ced, _cli), datos in registros_unicos.items():
        mapa_nombres.setdefault(ced, datos.get("nombre", ""))
    return mapa_cedulas, mapa_id, mapa_nombres


def test_digital_sin_verificacion_acepta_ambas_filas():
    doc, p = _pagina_dos_filas()
    mapa_cedulas, mapa_id, mapa_nombres = _mapas()
    palabras = p.get_text("words")
    y0s = _y0s_anclas_fila(p)
    res = list(_localizar_coincidencias_de_cedula(
        p, palabras, y0s, mapa_cedulas, mapa_id, mapa_nombres=mapa_nombres, verificar_nombre=False,
    ))
    cedulas = {r[2] for r in res}
    assert cedulas == {"506780001", "304560789"}   # comportamiento digital: match exacto, sin cruce
    doc.close()


def test_escaneado_con_verificacion_descarta_la_fila_con_nombre_equivocado():
    doc, p = _pagina_dos_filas()
    mapa_cedulas, mapa_id, mapa_nombres = _mapas()
    palabras = p.get_text("words")
    y0s = _y0s_anclas_fila(p)
    res = list(_localizar_coincidencias_de_cedula(
        p, palabras, y0s, mapa_cedulas, mapa_id, mapa_nombres=mapa_nombres, verificar_nombre=True,
    ))
    cedulas = {r[2] for r in res}
    assert cedulas == {"304560789"}   # se descarta la cédula de RICARDO leída sobre la fila de LUIS
    doc.close()


# ---- recuperación por cercanía (fuzzy de 1 dígito, confirmada por el nombre) ---

def test_recuperar_cedula_unica_a_distancia_1():
    oficiales = frozenset({"506780901", "304560789"})
    assert _recuperar_cedula_por_cercania("506730901", oficiales) == "506780901"  # 1 dígito distinto


def test_no_recupera_si_es_ambiguo_o_lejano():
    oficiales = frozenset({"506780901", "506780902"})   # dos oficiales a distancia-1 del leído
    assert _recuperar_cedula_por_cercania("506780900", oficiales) is None          # ambiguo -> None
    assert _recuperar_cedula_por_cercania("123456789", frozenset({"506780901"})) is None  # lejano
    assert _recuperar_cedula_por_cercania("5067", frozenset({"506780901"})) is None        # muy corto


def test_escaneado_recupera_cedula_con_un_digito_mal_si_el_nombre_confirma():
    # el OCR leyó 506730901 (un dígito mal) en la fila de LUIS; con el nombre se recupera a LUIS
    doc, p = _pagina_una_fila("506730901", "LUIS CASTRO CHAVES")
    registros = [{"cedula": "506780901", "cliente": "VMA", "nombre": "Luis Castro Chaves"}]
    ru, mc, mid = _construir_mapas_cedula(registros)
    mn = {c: d.get("nombre", "") for (c, _cl), d in ru.items()}
    ofi = frozenset(mc)
    palabras = p.get_text("words"); y0s = _y0s_anclas_fila(p)
    # sin OCR (digital): no se recupera -match exacto manda
    assert list(_localizar_coincidencias_de_cedula(p, palabras, y0s, mc, mid,
               mapa_nombres=mn, verificar_nombre=False, cedulas_oficiales=ofi)) == []
    # escaneado + nombre confirma: se recupera a LUIS
    res = list(_localizar_coincidencias_de_cedula(p, palabras, y0s, mc, mid,
               mapa_nombres=mn, verificar_nombre=True, cedulas_oficiales=ofi))
    assert {r[2] for r in res} == {"506780901"}
    doc.close()


def test_escaneado_no_recupera_si_el_nombre_no_confirma():
    # 506730901 (cerca de LUIS) pero la fila es de otra persona -> NO se recupera (evita falso positivo)
    doc, p = _pagina_una_fila("506730901", "PEDRO ARAYA BRENES")
    registros = [{"cedula": "506780901", "cliente": "VMA", "nombre": "Luis Castro Chaves"}]
    ru, mc, mid = _construir_mapas_cedula(registros)
    mn = {c: d.get("nombre", "") for (c, _cl), d in ru.items()}
    ofi = frozenset(mc)
    palabras = p.get_text("words"); y0s = _y0s_anclas_fila(p)
    res = list(_localizar_coincidencias_de_cedula(p, palabras, y0s, mc, mid,
               mapa_nombres=mn, verificar_nombre=True, cedulas_oficiales=ofi))
    assert res == []
    doc.close()

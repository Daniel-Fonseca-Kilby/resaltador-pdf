"""
test_sesion_portal.py

Con SESION_VMA_SECRETO, Planillas solo deja entrar a quien inició sesión en el portal de servicios
(cookie vma_sesion), y en cada pedido le pregunta al portal quién es. Sin ella queda abierta, como siempre.
Aquí el portal es de mentira: PortalFalso hace de POST /api/sesion del portal.
"""
import json

import pytest

from api import sesion_vma
from api.index import _historial_solicitudes_por_ip, app

SECRETO = "p" * 40


class PortalFalso:
    """Lo que contesta el portal a POST /api/sesion: las sesiones abiertas, por identificador."""

    def __init__(self):
        self.sesiones, self.pedidos, self.caido = {}, [], False

    def __call__(self, cuerpo: bytes, clave: str):
        if self.caido:
            raise OSError("conexión rechazada")
        datos = json.loads(cuerpo)
        self.pedidos.append(datos)
        if clave != sesion_vma.clave_consulta(SECRETO):
            return 403, {"error": "solo desde este servidor"}
        persona = self.sesiones.get(sesion_vma.id_de_cookie(datos["cookie"], SECRETO))
        return (200, persona) if persona else (401, {"error": "sin sesion"})


@pytest.fixture
def portal():
    return PortalFalso()


@pytest.fixture
def cliente(monkeypatch, portal):
    _historial_solicitudes_por_ip.clear()
    monkeypatch.setitem(app.config, "SESION_VMA_SECRETO", SECRETO)
    monkeypatch.setitem(app.config, "PORTAL_PUERTO", 8080)
    monkeypatch.setitem(app.config, "CONSULTA_PORTAL", sesion_vma.ConsultaPortal(8080, SECRETO, enviar=portal))
    cliente = app.test_client()
    cliente.portal_falso = portal
    return cliente


def _con_sesion(cliente, rol="equipo", secreto=SECRETO, servicios=("planillas",)):
    """Una sesión abierta en el portal (de mentira) y su cookie en el navegador."""
    id_ = sesion_vma.nuevo_id()
    cliente.portal_falso.sesiones[id_] = {"usuario": "ana.mora", "nombre": "Ana Mora", "rol": rol,
                                          "servicios": list(servicios)}
    cliente.set_cookie(sesion_vma.COOKIE, sesion_vma.valor_cookie(id_, secreto), domain="10.0.0.5")
    return cliente


def test_sin_sesion_la_pagina_manda_al_portal_y_vuelve_aqui(cliente):
    respuesta = cliente.get("/?x=1", base_url="http://10.0.0.5:8001")
    assert respuesta.status_code == 302
    assert respuesta.headers["Location"] == "http://10.0.0.5:8080/entrar?volver=http%3A%2F%2F10.0.0.5%3A8001%2F%3Fx%3D1"


def test_sin_sesion_la_api_responde_401(cliente):
    respuesta = cliente.post("/api/procesar", base_url="http://10.0.0.5:8001")
    assert respuesta.status_code == 401 and "portal" in respuesta.get_json()["error"]


def test_salud_y_estilos_no_piden_sesion(cliente):
    assert cliente.get("/salud").status_code == 200
    assert cliente.get("/static/vma/css/vma-tokens.css").status_code == 200


def test_con_sesion_muestra_quien_entro_y_salir(cliente):
    respuesta = _con_sesion(cliente).get("/", base_url="http://10.0.0.5:8001")
    pagina = respuesta.get_data(as_text=True)
    assert respuesta.status_code == 200 and "Ana Mora" in pagina
    assert 'href="http://10.0.0.5:8080/salir"' in pagina and 'href="http://10.0.0.5:8080/"' in pagina


def test_sin_acceso_a_planillas_vuelve_al_portal(cliente):
    _con_sesion(cliente, servicios=["licitaciones"])
    pagina = cliente.get("/", base_url="http://10.0.0.5:8001")
    assert pagina.status_code == 302 and pagina.headers["Location"] == "http://10.0.0.5:8080/?sin_acceso=planillas"
    assert cliente.post("/api/procesar", base_url="http://10.0.0.5:8001").status_code == 403
    _con_sesion(cliente, rol="admin", servicios=[])
    assert cliente.get("/", base_url="http://10.0.0.5:8001").status_code == 200  # los administradores entran a todo


def test_sesion_firmada_con_otra_clave_no_sirve(cliente):
    assert _con_sesion(cliente, secreto="q" * 40).get("/", base_url="http://10.0.0.5:8001").status_code == 302


def test_host_raro_no_sale_en_la_redireccion(cliente):
    respuesta = cliente.get("/", headers={"Host": "otro<sitio>"})
    assert respuesta.headers["Location"].startswith("http://localhost:8080/entrar")


def test_sin_secreto_queda_abierta(monkeypatch):
    monkeypatch.setitem(app.config, "SESION_VMA_SECRETO", None)
    pagina = app.test_client().get("/")
    assert pagina.status_code == 200 and "/salir" not in pagina.get_data(as_text=True)


def test_cerrada_en_el_portal_deja_de_servir(cliente):
    _con_sesion(cliente)
    assert cliente.get("/", base_url="http://10.0.0.5:8001").status_code == 200
    cliente.portal_falso.sesiones.clear()  # el portal la cerró (salió, la desactivaron, venció)
    app.config["CONSULTA_PORTAL"]._recordadas.clear()  # pasaron los 30 segundos
    assert cliente.get("/", base_url="http://10.0.0.5:8001").status_code == 302


def test_pedidos_automaticos_no_alargan_la_sesion(cliente):
    _con_sesion(cliente)
    cliente.get("/salud")  # sin sesión: ni pregunta
    assert cliente.portal_falso.pedidos == []
    cliente.post("/api/procesar", base_url="http://10.0.0.5:8001", headers={"X-Sin-Actividad": "1"})
    assert cliente.portal_falso.pedidos[-1]["actividad"] is False


def test_portal_caido(cliente):
    _con_sesion(cliente)
    cliente.portal_falso.caido = True
    pagina = cliente.get("/", base_url="http://10.0.0.5:8001")
    assert pagina.status_code == 503 and "no responde" in pagina.get_data(as_text=True)
    api = cliente.post("/api/procesar", base_url="http://10.0.0.5:8001")
    assert api.status_code == 503 and "portal" in api.get_json()["error"]


# --- Lo que miraría un auditor ------------------------------------------------------------------------------------
def test_ninguna_ruta_responde_sin_sesion(cliente):
    """Si alguien agrega una ruta, queda protegida sola; esto lo vigila."""
    rutas = [r for r in app.url_map.iter_rules() if not r.rule.startswith(("/salud", "/static"))]
    assert len(rutas) >= 3
    for regla in rutas:
        for metodo in regla.methods - {"HEAD", "OPTIONS"}:
            respuesta = cliente.open(regla.rule, method=metodo, base_url="http://10.0.0.5:8001")
            assert respuesta.status_code in (302, 401), f"{metodo} {regla.rule} responde sin sesión"


@pytest.mark.parametrize("encabezados", [{"Origin": "http://otro-sitio.com"}, {"Origin": "http://10.0.0.5:8000"},
                                         {"Origin": "null"}, {"Referer": "http://otro-sitio.com/x"}])
@pytest.mark.parametrize("ruta", ["/api/procesar", "/api/detectar-modo-excel"])
def test_no_se_procesa_nada_desde_otro_origen(cliente, ruta, encabezados):
    _con_sesion(cliente)
    respuesta = cliente.post(ruta, base_url="http://10.0.0.5:8001", headers=encabezados)
    assert respuesta.status_code == 403 and "otro sitio" in respuesta.get_json()["error"]


def test_desde_la_propia_app_si_se_procesa(cliente):
    _con_sesion(cliente)
    respuesta = cliente.post("/api/procesar", base_url="http://10.0.0.5:8001", headers={"Origin": "http://10.0.0.5:8001"})
    assert respuesta.status_code == 400  # llegó a la app: le falta el PDF


def test_en_el_servidor_no_arranca_sin_inicio_de_sesion(monkeypatch):
    """Si alguien borra SESION_VMA_SECRETO del .env, Planillas no queda abierta: no arranca."""
    from api import index
    monkeypatch.setenv("EN_SERVIDOR", "1")
    monkeypatch.delenv("PERMITIR_SIN_LOGIN", raising=False)
    monkeypatch.setitem(app.config, "SESION_VMA_SECRETO", None)
    with pytest.raises(RuntimeError, match="sin inicio de sesión"):
        index.revisar_configuracion()


def test_con_https_las_direcciones_son_https(cliente, monkeypatch):
    monkeypatch.setenv("VMA_HTTPS", "1")
    respuesta = cliente.get("/?x=1", base_url="http://10.0.0.5:8001")
    assert respuesta.headers["Location"] == "https://10.0.0.5:8080/entrar?volver=https%3A%2F%2F10.0.0.5%3A8001%2F%3Fx%3D1"
    pagina = _con_sesion(cliente).get("/", base_url="http://10.0.0.5:8001")
    assert 'href="https://10.0.0.5:8080/"' in pagina.get_data(as_text=True)
    assert pagina.headers["Strict-Transport-Security"] == "max-age=31536000"

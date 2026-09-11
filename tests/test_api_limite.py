
import api.index as api_index

_cliente = api_index.app.test_client()


def _post(ip: str, ruta: str = "/api/detectar-modo-excel"):
    """POST simulando que la solicitud viene de 'ip' -así se puede probar
    el límite sin depender de la IP real de quien corre las pruebas."""
    return _cliente.post(ruta, environ_base={"REMOTE_ADDR": ip})


def _get(ip: str, ruta: str = "/"):
    return _cliente.get(ruta, environ_base={"REMOTE_ADDR": ip})


def setup_function(_funcion):
    # cada prueba arranca con el historial limpio -si no, el orden en que
    # pytest corre los tests contaminaría el conteo entre uno y otro
    api_index._historial_solicitudes_por_ip.clear()


def test_dentro_del_limite_deja_pasar_todas_las_solicitudes():
    for _ in range(api_index._LIMITE_SOLICITUDES_POR_IP):
        respuesta = _post("10.0.0.1")
        assert respuesta.status_code != 429


def test_pasarse_del_limite_bloquea_con_429():
    for _ in range(api_index._LIMITE_SOLICITUDES_POR_IP):
        _post("10.0.0.2")

    respuesta = _post("10.0.0.2")

    assert respuesta.status_code == 429
    assert "error" in respuesta.get_json()


def test_ip_distinta_tiene_su_propio_limite():
    """Una IP que ya se pasó del límite no debe afectar a otra distinta."""
    for _ in range(api_index._LIMITE_SOLICITUDES_POR_IP + 1):
        _post("10.0.0.3")

    respuesta = _post("10.0.0.4")

    assert respuesta.status_code != 429


def test_pasado_el_tiempo_de_la_ventana_se_libera_de_nuevo(monkeypatch):
    ahora = 1_000_000.0
    monkeypatch.setattr(api_index.time, "time", lambda: ahora)

    for _ in range(api_index._LIMITE_SOLICITUDES_POR_IP):
        _post("10.0.0.5")
    assert _post("10.0.0.5").status_code == 429

    # avanza el reloj más allá de la ventana -las solicitudes viejas ya no cuentan
    ahora += api_index._VENTANA_LIMITE_SEGUNDOS + 1
    monkeypatch.setattr(api_index.time, "time", lambda: ahora)

    assert _post("10.0.0.5").status_code != 429


def test_el_limite_no_afecta_la_pagina_principal():
    """El límite solo aplica a /api/* -la página principal (sin login
    todavía) debe seguir cargando siempre, aunque una IP se pase del
    límite en la API."""
    for _ in range(api_index._LIMITE_SOLICITUDES_POR_IP + 5):
        _post("10.0.0.6")

    respuesta = _get("10.0.0.6")

    assert respuesta.status_code == 200

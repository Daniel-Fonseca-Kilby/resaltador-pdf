"""Sesión compartida de las apps internas de Grupo VMA (portal de servicios, Licitaciones, Planillas).

El portal es la puerta de entrada: cuando alguien inicia sesión, guarda la sesión en su base (tabla «sesion») y deja en
el navegador la cookie «vma_sesion», que solo lleva un identificador al azar firmado con una clave secreta que conocen
las tres apps (variable de entorno SESION_VMA_SECRETO, la misma en las tres). Como las tres están en el mismo
servidor, el navegador manda esa cookie a todas (las cookies distinguen el nombre del servidor, no el puerto).

Licitaciones y Planillas no tienen la base del portal: le preguntan al portal, por 127.0.0.1, quién es esa sesión
(ConsultaPortal), y recuerdan la respuesta 30 segundos. Así una cuenta desactivada, una sesión cerrada por un
administrador o 30 minutos sin uso cortan el acceso a las tres apps en menos de un minuto.

La cookie no lleva la contraseña ni nada que sirva para sacarla. Si alguien cambia el identificador, la firma deja de
coincidir y se rechaza sin consultar la base.

firmar() y leer() siguen sirviendo para datos firmados de corta vida (el «pase» del cambio de contraseña del portal).

Este archivo se edita solo en Proyecto Licitaciones/compartido/ y se copia a cada app con marca/sincronizar.py.
Solo usa la librería estándar de Python.
"""
import base64
import hashlib
import hmac
import json
import os
import secrets
import threading
import time
import urllib.error
import urllib.request
from http.cookies import CookieError, SimpleCookie
from urllib.parse import quote

COOKIE = "vma_sesion"
HORAS = 12
ROLES = ("admin", "equipo")
TODAS = "*"  # en la lista de herramientas: todas
ENCABEZADO_CLAVE = "X-Clave-Portal"            # la llevan las consultas de las apps al portal
ENCABEZADO_SIN_ACTIVIDAD = "X-Sin-Actividad"   # «1» en pedidos automáticos: no alargan la sesión
_LARGO_MINIMO_SECRETO = 32


def _b64(datos: bytes) -> str:
    return base64.urlsafe_b64encode(datos).rstrip(b"=").decode("ascii")


def _de_b64(texto: str) -> bytes:
    return base64.urlsafe_b64decode(texto + "=" * (-len(texto) % 4))


def _firma(cuerpo: str, secreto: str) -> str:
    return _b64(hmac.new(secreto.encode("utf-8"), cuerpo.encode("ascii"), hashlib.sha256).digest())


def revisar_secreto(secreto: str | None) -> str:
    if not secreto or len(secreto) < _LARGO_MINIMO_SECRETO:
        raise ValueError(f"SESION_VMA_SECRETO debe tener al menos {_LARGO_MINIMO_SECRETO} caracteres al azar.")
    return secreto


def firmar(usuario: str, nombre: str, rol: str, secreto: str, ahora: float | None = None, horas: int = HORAS,
           servicios=()) -> str:
    """El valor de la cookie: datos y firma, separados por un punto. `servicios`: ids de las herramientas a las que
    puede entrar («*» para todas)."""
    if rol not in ROLES:
        raise ValueError(f"Rol no válido: {rol}")
    ahora = time.time() if ahora is None else ahora
    datos = {"u": usuario, "n": nombre, "r": rol, "s": [str(s) for s in servicios], "exp": int(ahora + horas * 3600)}
    cuerpo = _b64(json.dumps(datos, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))
    return f"{cuerpo}.{_firma(cuerpo, revisar_secreto(secreto))}"


def leer(valor: str | None, secreto: str, ahora: float | None = None) -> dict | None:
    """{"usuario", "nombre", "rol", "servicios"} si la cookie es válida y no venció; si no, None (nunca lanza error)."""
    if not valor or "." not in valor or not secreto:
        return None
    cuerpo, firma = valor.rsplit(".", 1)
    if not hmac.compare_digest(firma, _firma(cuerpo, secreto)):
        return None
    try:
        datos = json.loads(_de_b64(cuerpo))
    except (ValueError, UnicodeDecodeError):
        return None
    ahora = time.time() if ahora is None else ahora
    if not isinstance(datos, dict) or datos.get("r") not in ROLES or int(datos.get("exp", 0)) <= ahora:
        return None
    servicios = datos.get("s")
    servicios = [str(s) for s in servicios] if isinstance(servicios, list) else []
    return {"usuario": str(datos["u"]), "nombre": str(datos["n"]), "rol": datos["r"], "servicios": servicios}


def puede_entrar(datos: dict, servicio: str) -> bool:
    """Si la sesión (lo que devuelve leer) da acceso a la herramienta `servicio` (su id en servicios.json del portal).
    Los administradores entran a todas."""
    return datos["rol"] == "admin" or TODAS in datos["servicios"] or servicio in datos["servicios"]


def valor_de_encabezado(encabezado_cookie: str | None) -> str | None:
    """El valor de vma_sesion dentro del encabezado «Cookie» completo del pedido."""
    if not encabezado_cookie:
        return None
    try:
        galletas = SimpleCookie(encabezado_cookie)
    except CookieError:
        return None
    galleta = galletas.get(COOKIE)
    return galleta.value if galleta else None


def leer_de_encabezado(encabezado_cookie: str | None, secreto: str, ahora: float | None = None) -> dict | None:
    """leer(), a partir del encabezado «Cookie» completo del pedido."""
    return leer(valor_de_encabezado(encabezado_cookie), secreto, ahora)


# Sesiones guardadas en el portal
def nuevo_id() -> str:
    """Identificador de una sesión nueva: 32 bytes al azar."""
    return secrets.token_urlsafe(32)


def valor_cookie(id_: str, secreto: str) -> str:
    """El valor de vma_sesion: el identificador y su firma, separados por un punto."""
    return f"{id_}.{_firma(id_, revisar_secreto(secreto))}"


def id_de_cookie(valor: str | None, secreto: str) -> str | None:
    """El identificador, si la firma coincide; si no, None (nunca lanza error)."""
    if not valor or not secreto or valor.count(".") != 1:
        return None
    id_, firma = valor.split(".")
    if not id_ or not id_.isascii() or not hmac.compare_digest(firma.encode(), _firma(id_, secreto).encode()):
        return None
    return id_


def hash_id(id_: str) -> str:
    """Lo que guarda el portal en vez del identificador: con una copia de la base no se pueden usar las sesiones."""
    return hashlib.sha256(id_.encode("ascii")).hexdigest()


def clave_consulta(secreto: str) -> str:
    """La clave con que las apps le preguntan al portal por una sesión (sale del secreto, no lo revela)."""
    return _b64(hmac.new(secreto.encode("utf-8"), b"consulta-sesion", hashlib.sha256).digest())


class PortalNoResponde(Exception):
    """El portal no contestó: la app no deja pasar a nadie (y tampoco lo manda a entrar, que sería un bucle)."""


_SIN_PROXY = urllib.request.build_opener(urllib.request.ProxyHandler({}))  # 127.0.0.1 nunca va por un proxy


class ConsultaPortal:
    """Le pregunta al portal (POST http://127.0.0.1:<puerto>/api/sesion) quién es la sesión de una cookie.

    Recuerda cada respuesta `segundos_recordar` (también «no vale»): una app con mucho tráfico consulta como mucho una
    vez cada 30 segundos por persona. `enviar(cuerpo, clave) -> (código HTTP, JSON)` se cambia en las pruebas."""

    _MAXIMO_RECORDADAS = 5000

    def __init__(self, puerto: int, secreto: str, segundos_recordar: float = 30, espera: float = 3.0,
                 reloj=time.monotonic, enviar=None):
        self.secreto = revisar_secreto(secreto)
        self.url = f"http://127.0.0.1:{int(puerto)}/api/sesion"
        self.segundos_recordar, self.espera, self.reloj = segundos_recordar, espera, reloj
        self.enviar = enviar or self._enviar
        self._recordadas: dict[str, tuple[float, dict | None]] = {}
        self._candado = threading.Lock()

    def _enviar(self, cuerpo: bytes, clave: str) -> tuple[int, dict]:
        pedido = urllib.request.Request(self.url, data=cuerpo, method="POST", headers={
            "Content-Type": "application/json", ENCABEZADO_CLAVE: clave})
        try:
            with _SIN_PROXY.open(pedido, timeout=self.espera) as respuesta:
                return respuesta.status, json.loads(respuesta.read())
        except urllib.error.HTTPError as error:
            try:
                return error.code, json.loads(error.read() or b"{}")
            except ValueError:
                return error.code, {}

    def persona(self, valor: str | None, actividad: bool = True) -> dict | None:
        """{"usuario", "nombre", "rol", "servicios"} si la sesión vale; None si no. PortalNoResponde si no contesta."""
        if id_de_cookie(valor, self.secreto) is None:
            return None
        ahora = self.reloj()
        with self._candado:
            recordada = self._recordadas.get(valor)
        if recordada and ahora - recordada[0] < self.segundos_recordar:
            return recordada[1]
        cuerpo = json.dumps({"cookie": valor, "actividad": bool(actividad)}).encode("utf-8")
        try:
            codigo, datos = self.enviar(cuerpo, clave_consulta(self.secreto))
        except (OSError, ValueError) as error:  # URLError y los tiempos de espera son OSError
            raise PortalNoResponde(str(error)) from error
        if codigo == 401:
            resultado = None
        elif codigo == 200 and isinstance(datos, dict) and datos.get("rol") in ROLES:
            servicios = datos.get("servicios")
            resultado = {"usuario": str(datos.get("usuario", "")), "nombre": str(datos.get("nombre", "")),
                         "rol": datos["rol"],
                         "servicios": [str(x) for x in servicios] if isinstance(servicios, list) else []}
        else:
            raise PortalNoResponde(f"el portal respondió {codigo}")
        with self._candado:
            if len(self._recordadas) >= self._MAXIMO_RECORDADAS:
                self._recordadas.clear()
            self._recordadas[valor] = (ahora, resultado)
        return resultado

    def persona_de_encabezado(self, encabezado_cookie: str | None, actividad: bool = True) -> dict | None:
        return self.persona(valor_de_encabezado(encabezado_cookie), actividad)


def esquema() -> str:
    """https con VMA_HTTPS=1 en el .env de las tres apps (cuando TI dé el certificado); si no, http."""
    return "https" if os.environ.get("VMA_HTTPS") == "1" else "http"


def _con_https() -> str:
    return "; Secure" if esquema() == "https" else ""


def puerto_para_consultar(puerto_portal: int) -> int:
    """Dónde preguntarle al portal: su puerto o, con un proxy de HTTPS delante, PORTAL_PUERTO_INTERNO."""
    return int(os.environ.get("PORTAL_PUERTO_INTERNO") or puerto_portal)


def encabezado_guardar(valor: str, segundos: int = HORAS * 3600) -> str:
    """Set-Cookie para guardar la sesión. Sin Domain: vale para este servidor, en cualquier puerto. HttpOnly: el
    JavaScript de las páginas no la puede leer. SameSite=Lax: otro sitio no puede usarla para mandar formularios.
    Secure, con HTTPS: el navegador nunca la manda sin cifrar."""
    return f"{COOKIE}={valor}; Path=/; Max-Age={int(segundos)}; HttpOnly; SameSite=Lax{_con_https()}"


def encabezado_borrar() -> str:
    return f"{COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax{_con_https()}"


def url_entrar(host: str, puerto_portal: int, volver: str = "") -> str:
    """Adónde mandar a quien llega sin sesión: la página de entrada del portal, que lo devuelve después."""
    destino = f"{esquema()}://{host}:{int(puerto_portal)}/entrar"
    return f"{destino}?volver={quote(volver, safe='')}" if volver else destino


def url_sin_acceso(host: str, puerto_portal: int, servicio: str) -> str:
    """El portal, con el aviso de que esta cuenta no tiene acceso a la herramienta."""
    return f"{esquema()}://{host}:{int(puerto_portal)}/?sin_acceso={quote(servicio, safe='')}"


def url_salir(host: str, puerto_portal: int) -> str:
    return f"{esquema()}://{host}:{int(puerto_portal)}/salir"


def url_portal(host: str, puerto_portal: int) -> str:
    return f"{esquema()}://{host}:{int(puerto_portal)}/"

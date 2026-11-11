"""Encabezados de seguridad comunes de las apps internas de Grupo VMA (portal, Licitaciones, Planillas).

Cada app los agrega a sus respuestas a su manera (WSGI, ASGI o Flask), pero la lista y la política de contenido (CSP)
salen de aquí, para que las tres digan lo mismo:

- que el navegador no adivine tipos de archivo, que ninguna otra página meta la app en un marco (clickjacking), que
  a otros sitios no les llegue la dirección de la app y que la página no pida cámara, micrófono, ubicación ni pagos;
- la CSP solo deja correr scripts de la propia app: archivos del mismo servidor y los scripts escritos dentro del
  HTML cuyo hash va en la política. El hash se calcula del HTML que se sirve (csp(html)), así que editar un script
  no rompe nada en silencio. Un script inyectado (XSS) no tiene su hash y el navegador no lo corre. Los estilos en
  línea se permiten: las apps los usan y su riesgo es mucho menor.

Este archivo se edita solo en Proyecto Licitaciones/compartido/ y se copia a cada app con marca/sincronizar.py.
Solo usa la librería estándar de Python.
"""
import base64
import hashlib
import os
import re
from urllib.parse import urlsplit

ENCABEZADOS = [
    ("X-Content-Type-Options", "nosniff"),
    ("X-Frame-Options", "DENY"),
    ("Referrer-Policy", "same-origin"),
    ("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()"),
]
# Solo con HTTPS. Sobre HTTP el navegador IGNORA estos dos y, además, deja un error en la consola; no aportan nada sin
# cifrado, así que no se mandan (COOP necesita un origen seguro; HSTS obliga a HTTPS).
_SOLO_HTTPS = [
    ("Cross-Origin-Opener-Policy", "same-origin"),
    ("Strict-Transport-Security", "max-age=31536000"),
]

# <script ...>texto</script>, sin importar mayúsculas; el texto tal cual (el navegador calcula el hash sobre él).
_SCRIPT = re.compile(r"<script\b([^>]*)>(.*?)</script\s*>", re.IGNORECASE | re.DOTALL)
_CON_SRC = re.compile(r"\bsrc\s*=", re.IGNORECASE)
_ESTILO = re.compile(r"<style\b[^>]*>(.*?)</style\s*>", re.IGNORECASE | re.DOTALL)
_ATRIBUTO_STYLE = re.compile(r"<[a-zA-Z][^<>]*\sstyle\s*=", re.IGNORECASE)


def _hash(texto: str) -> str:
    return "'sha256-" + base64.b64encode(hashlib.sha256(texto.encode("utf-8")).digest()).decode("ascii") + "'"


def encabezados() -> list[tuple[str, str]]:
    """Los comunes siempre; con HTTPS (VMA_HTTPS=1) además COOP y HSTS (sobre HTTP el navegador los ignora y ensucia
    la consola)."""
    return ENCABEZADOS + (_SOLO_HTTPS if os.environ.get("VMA_HTTPS") == "1" else [])


def host_permitido(host: str, por_defecto: str = "localhost") -> str:
    """El host con el que la app arma los enlaces y las redirecciones al portal. Sin VMA_HOSTS (desarrollo) se acepta el
    que vino en el pedido. Con VMA_HOSTS (una lista de nombres o IP separados por comas, en el .env), un Host que no
    esté en la lista —posible envenenamiento para redirigir a un sitio falso— se reemplaza por el primero de la lista.
    Hay que listar TODOS los nombres e IP con que el equipo entra, o quien use uno que falte terminará en el primero."""
    permitidos = [h.strip() for h in os.environ.get("VMA_HOSTS", "").split(",") if h.strip()]
    if not permitidos:
        return host
    return host if host.lower() in {h.lower() for h in permitidos} else permitidos[0]


def revisar_login_en_servidor(con_login: bool) -> None:
    """En el servidor (iniciar.ps1 pone EN_SERVIDOR=1), una app sin inicio de sesión no arranca: si alguien borra la
    clave del .env, es mejor que no responda a que quede abierta. Solo la primera instalación puede arrancar sin él, a
    propósito, con PERMITIR_SIN_LOGIN=1 (lo pone instalar.ps1 -SinLogin y lo quita preparar_sesion.ps1)."""
    if not con_login and os.environ.get("EN_SERVIDOR") == "1" and os.environ.get("PERMITIR_SIN_LOGIN") != "1":
        raise RuntimeError("No se arranca sin inicio de sesión en el servidor: revisa el .env de esta app "
                           "(LOGIN y SESION_VMA_SECRETO; ver README).")


def hashes_de_scripts(html: str) -> list[str]:
    """«'sha256-…'» de cada script escrito dentro del HTML (los que tienen src no cuentan: son archivos)."""
    return [_hash(texto) for atributos, texto in _SCRIPT.findall(html or "") if not _CON_SRC.search(atributos)]


def csp(html: str = "", form_action: str = "'self'", estilos_en_linea: bool = True) -> str:
    """La política para una página (con los hashes de sus scripts) o, sin HTML, para cualquier otra respuesta.
    `estilos_en_linea=False`: tampoco estilos inyectados, solo los <style> de la página, por su hash. Licitaciones no
    puede: las librerías de su interfaz (Radix, sonner) crean sus propios <style> al correr."""
    scripts = " ".join(["'self'", *dict.fromkeys(hashes_de_scripts(html))])
    estilos = "'self' 'unsafe-inline'"
    if not estilos_en_linea:
        estilos = " ".join(["'self'", *dict.fromkeys(_hash(texto) for texto in _ESTILO.findall(html or ""))])
    return ("default-src 'self'; "
            f"script-src {scripts}; "
            f"style-src {estilos}; "
            "img-src 'self' data: blob:; "
            "font-src 'self' data:; "
            "connect-src 'self'; "
            "frame-ancestors 'none'; "
            "object-src 'none'; "
            "base-uri 'self'; "
            f"form-action {form_action}")


def mismo_origen(origen: str | None, host: str | None) -> bool:
    """Si un pedido que cambia datos viene de la propia app. `origen`: su encabezado Origin o, si no trae, Referer;
    `host`: su encabezado Host. Otra página no puede falsificar esos encabezados, y el navegador los manda siempre en
    un POST entre páginas. SameSite no alcanza: no distingue puertos, y en este servidor hay otros programas."""
    if not origen:
        return True
    try:
        return bool(host) and urlsplit(origen).netloc.lower() == host.lower()
    except ValueError:
        return False


def estilos_cubiertos(html: str, politica: str) -> bool:
    """Si los <style> de la página tienen su hash y no hay atributos style (para las pruebas de cada app)."""
    return not _ATRIBUTO_STYLE.search(html) and all(_hash(texto) in politica for texto in _ESTILO.findall(html))


def scripts_cubiertos(html: str, politica: str) -> bool:
    """Si todos los scripts en línea del HTML tienen su hash en la política (para las pruebas de cada app)."""
    return all(h in politica for h in hashes_de_scripts(html))

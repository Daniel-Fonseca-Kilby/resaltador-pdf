# Resaltador de Planillas en PDF

Aplicación web (Flask) que abre uno o varios PDFs de planilla, busca
nombres o cédulas, y genera copias resaltadas en amarillo — sin modificar
el PDF original ni perder su encabezado o formato. Se usa desde el
navegador, no requiere instalar nada en la computadora del usuario final.

## Modos de uso

La página (`api/templates/index.html`) tiene dos flujos, y el sistema
elige uno automáticamente según lo que se suba:

### 1. Modo simple — nombres escritos a mano
- Se escriben uno o varios nombres separados por coma (o se sube un Excel
  con una sola columna de nombres).
- Se seleccionan uno o varios PDFs.
- Por cada PDF se genera una copia con esos nombres resaltados, tal cual
  el documento original (mismas páginas, mismo formato).
- El resultado indica, por archivo, qué nombres se encontraron y cuántas
  veces.

### 2. Modo por cliente — Excel con Identificación y Cliente
- Se sube un Excel con columnas **Identificación** y **Cliente** (y
  opcionalmente **Nombre**) — el listado completo de empleados por
  cliente/empresa.
- Se seleccionan los PDFs de la planilla (pueden ser varias páginas,
  varios clientes mezclados).
- El sistema busca cada cédula del Excel dentro de los PDFs, recorta
  **solo esa fila** (con su apariencia exacta) y arma un PDF nuevo por
  cliente, con las filas de ese cliente apiladas una debajo de otra y
  resaltadas — sin exponer filas de empleados de otros clientes.
- Al inicio de cada PDF de cliente se antepone un recorte del encabezado
  de la página de origen (empresa, período, títulos de columna).
- El campo **Tipo de planilla** (Automático / CCSS / MNK / INS) ayuda a
  ubicar con más precisión dónde termina el encabezado de cada formato
  conocido; en modo automático el sistema igual intenta reconocer los
  formatos conocidos antes de recurrir a detección genérica.
- El resultado es un `.zip` con un PDF por cliente, más el listado de
  cédulas del Excel que no se encontraron en ningún PDF.

En ambos modos, la descarga es un `.zip` que se genera en el navegador
al terminar de procesar.

## Estructura del proyecto

- `api/` — aplicación web Flask: rutas (`index.py`, decide el modo, arma
  el `.zip` de respuesta, es el punto de entrada para Render/gunicorn) y
  la plantilla (`templates/index.html`, HTML + CSS + JS en un solo
  archivo).
- `api/static/vma/` — **librería visual de Grupo VMA** (libro de marca de
  ene. 2025): Poppins, colores y logos. Es una copia: no se edita aquí. La
  original está en el repositorio de Licitaciones (`marca/vma/`, con su
  catálogo en `marca/catalogo.html`) y se copia con
  `python marca/sincronizar.py` desde allá. La plantilla traduce sus
  propias variables (`--azul`, `--dorado`…) a las de la librería, así que
  un cambio de marca llega solo con sincronizar.
- `resaltado_pdf.py` — núcleo de la lógica de negocio: búsqueda de texto,
  resaltado, recorte y armado de PDFs por cliente con PyMuPDF. Sin
  dependencias de Flask, para poder reutilizarlo o probarlo aparte.
- `tests/` — pruebas automatizadas con pytest (ver sección de abajo).
- `Procfile` — comando de arranque para Render (`gunicorn`, con
  `--workers 1 --threads 4` para no bloquear a otros usuarios mientras se
  procesa una planilla pesada, dado que el plan de Render solo tiene 1 CPU).

## Cómo correrlo en local (para desarrollo)

1. Instalar las dependencias:
   ```
   pip install -r requirements-dev.txt
   ```
2. Ejecutar el servidor de desarrollo:
   ```
   python -m flask --app api.index run --debug
   ```
   (o `python api/index.py`, que arranca Flask en modo debug en el puerto
   por defecto).
3. Abrir `http://127.0.0.1:5000` en el navegador.

## En un servidor Windows (JUPITER, junto a Licitaciones)

En Windows no corre gunicorn: la app se sirve con **waitress** (`servidor/arrancar.py`, un proceso con 4 hilos;
los lotes de PDF van de uno en uno por el candado de PyMuPDF) como tarea programada «Planillas VMA», que arranca
con Windows y se vuelve a levantar si se cae. Mismo esquema que Licitaciones.

1. **En la PC:** `powershell -ExecutionPolicy Bypass -File servidor\empaquetar.ps1` deja
   `paquetes\planillas-<fecha>.zip` (con `VERSION.txt`: commit y fecha).
2. **En el servidor** (Python 3.14 «para todos los usuarios»): copiar a `C:\Planillas` el ZIP y
   `servidor\instalar.ps1`, y en PowerShell como administrador:
   ```
   cd C:\Planillas
   powershell -ExecutionPolicy Bypass -File .\instalar.ps1 -Paquete C:\Planillas\planillas-<fecha>.zip
   ```
   Instala las librerías en `C:\Planillas\venv`, comprueba que la app carga, crea la tarea y la regla de firewall
   del puerto 8001 (solo red interna) y espera a que responda en `http://<servidor>:8001`. Para actualizar, lo
   mismo con el ZIP nuevo.
3. **Día a día:** registros en `C:\Planillas\logs` (`app.log` uno por día, `arranque.log`, `errores.log`);
   temporales de cada lote en `C:\Planillas\temp`; reiniciar con
   `powershell -ExecutionPolicy Bypass -File C:\Planillas\app\servidor\reiniciar.ps1` (detener solo la tarea deja
   Python ocupando el puerto).

En el servidor se instala `requirements-servidor.lock` (versiones exactas y hashes; se genera de `requirements.txt` con
`pip-tools`, ver `servidor/revisar_dependencias.py` de Proyecto Licitaciones). Las directas van con versión exacta en `requirements.txt` (`pymupdf` 1.28.2: sin eso, cada instalación nueva
bajaría la última, y ya hay 2.x). Las revisa ese mismo script (pip-audit y Bandit).

**Inicio de sesión (portal de servicios).** Con `SESION_VMA_SECRETO` en `C:\Planillas\.env` (la misma clave del
portal, en el puerto `PORTAL_PUERTO`, 8080 por defecto), Planillas solo deja entrar a quien inició sesión en el
portal y tiene acceso a «Planillas» (o es administrador): sin sesión manda a la página de entrada del portal, sin
acceso vuelve al portal con un aviso, y la API responde 401/403. El encabezado muestra quién entró, un botón al
portal y «Salir». `/salud` y `/static/` quedan abiertos (revisiones del portal y del instalador). Sin
`SESION_VMA_SECRETO` (Vercel, desarrollo) todo sigue abierto como antes. La clave la escribe el script del portal
`preparar_sesion.ps1 -ActivarApps`. En cada pedido Planillas le pregunta al portal (por `127.0.0.1`) quién es esa
sesión y recuerda la respuesta 30 segundos (`api/sesion_vma.py`, copia del archivo compartido que vive en Proyecto
Licitaciones, `compartido/sesion_vma.py`, se copia con `marca/sincronizar.py`): una cuenta desactivada, una sesión
cerrada o 30 minutos sin uso cortan el acceso en menos de un minuto. Si el portal no responde, responde 503.

**En el servidor no arranca sin `SESION_VMA_SECRETO`.** La primera instalación puede, a propósito, con
`instalar.ps1 -SinLogin`. Los PDF se cortan a las 5000 páginas, y un lote a los 15 minutos.

**Archivos subidos.** Los nombres se limpian antes de tocar el disco (`../`, flujos de NTFS, `NUL.pdf`). Un PDF se
reconoce por el nombre y por su contenido. Un Excel que se infle de más al descomprimirlo o traiga DOCTYPE se rechaza. En
el resumen en Excel, lo que empiece con `=` queda como texto, no como fórmula. Con sesión, lo que manda datos solo se
acepta desde la propia app (CSRF), y `tests/test_sesion_portal.py` recorre las rutas para que ninguna quede sin sesión.

**Seguridad.** Los encabezados comunes de las apps de Grupo VMA (`api/seguridad_web.py`, otra copia compartida) y una
CSP que solo deja correr los scripts de la página (sus hashes salen del HTML que se sirve). La página y la API van sin
caché. JSZip 3.10.1 se sirve desde `api/static/vendor/` (con su licencia), no desde un CDN de internet. Si
`DEBUG_GUARDAR_SUBIDAS=1` queda encendido, avisa al arrancar: guarda copias de las planillas (datos de personas).

## Pruebas automatizadas

`tests/` tiene pruebas de `resaltado_pdf.py` (búsqueda de nombres,
tolerancia a la Ñ rota de MNK, exportación por cliente, detección de
encabezado). El PDF de ejemplo se genera en memoria con nombres y
cédulas **inventados** (ver `tests/conftest.py`) — no se necesita ni se
sube ningún dato real para probar.

```
pip install -r requirements-dev.txt
pytest
```

ue local no actualiza el sitio en línea.

Como no hay persistencia, cada solicitud procesa sus PDFs en una carpeta
temporal aislada que se borra al terminar. Si una solicitud se cancela o
se cae a mitad de camino esa carpeta puede quedar huérfana; al arrancar
el proceso (`api/index.py`, `_limpiar_temporales_antiguos`) se purgan
automáticamente las que lleven más de una hora sin tocarse, para no
saturar el disco efímero de Render.

## Notas de formato de PDF

Algunos generadores de reportes (visto en MNK) no exportan bien el
glifo de la Ñ/ñ en el texto interno del PDF — puede quedar como un
espacio en blanco aunque visualmente se vea la Ñ. `resaltado_pdf.py`
tolera esto probando variantes de búsqueda automáticamente.

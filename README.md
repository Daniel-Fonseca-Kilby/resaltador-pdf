# Resaltador de Planillas en PDF

Aplicación web (Flask) que abre uno o varios PDFs de planilla, busca
nombres o cédulas, y genera copias resaltadas en amarillo — sin modificar
el PDF original ni perder su encabezado o formato. Se usa desde el
navegador, no requiere instalar nada en la computadora del usuario final.

Desplegada en **Render** (ver `Procfile`). El proyecto vivió antes una
versión de escritorio (Tkinter) para probar la idea rápido; ese enfoque
se descartó a favor de la web por la comodidad de no tener que instalar
Python ni dependencias en cada máquina.

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

- `api/index.py` — servidor Flask: recibe la subida, decide el modo,
  arma el `.zip` de respuesta. Punto de entrada para Render/gunicorn.
- `api/templates/index.html` — la página completa (HTML + CSS + JS en un
  solo archivo).
- `resaltado_pdf.py` — toda la lógica de negocio: búsqueda de texto,
  resaltado, recorte y armado de PDFs por cliente. Sin dependencias de
  Flask ni de la interfaz, para poder reutilizarla o probarla aparte.
- `main.py`, `interfaz.py` — prototipo de escritorio (Tkinter) anterior
  a la versión web. Ya no es el foco del proyecto; se mantiene sin
  actualizar mientras se decide si se retira.
- `Procfile` — comando de arranque para Render (`gunicorn`).
- `vercel.json` — configuración de un despliegue anterior en Vercel; ya
  no se usa (el proyecto se movió a Render).

## Cómo correrlo en local (para desarrollo)

1. Instalar las dependencias:
   ```
   pip install -r requirements.txt
   ```
2. Ejecutar el servidor de desarrollo:
   ```
   python -m flask --app api.index run --debug
   ```
   (o `python api/index.py`, que arranca Flask en modo debug en el puerto
   por defecto).
3. Abrir `http://127.0.0.1:5000` en el navegador.

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

## Despliegue

Render ejecuta `Procfile` (`gunicorn api.index:app`) sobre la rama
conectada del repositorio. Cualquier cambio debe subirse a GitHub para
que Render lo tome — un despliegue local no actualiza el sitio en línea.

## Notas de formato de PDF

Algunos generadores de reportes (visto en MNK) no exportan bien el
glifo de la Ñ/ñ en el texto interno del PDF — puede quedar como un
espacio en blanco aunque visualmente se vea la Ñ. `resaltado_pdf.py`
tolera esto probando variantes de búsqueda automáticamente.

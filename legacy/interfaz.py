"""
interfaz.py
disclaimer: todo lo que el usuario VE y CLICKEA.
No contiene lógica de negocio (buscar/resaltar texto) eso vive en
resaltado_pdf.py. Esta separación es importante porque si mañana se
quiere cambiar esta ventana por, por ejemplo, una versión web, no hay
que tocar la lógica de resaltado en absoluto.
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from resaltado_pdf import resaltar_nombres_en_pdf


class VentanaResaltado(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Resaltador de Nombres en PDF")
        self.geometry("560x480")
        self.resizable(False, False)

        self.rutas_pdfs_seleccionados: list[str] = []

        self._construir_interfaz()

  #--Construcion de la interfaz
    def _construir_interfaz(self):
        padding = {"padx": 16, "pady": 8}

        # --- Paso 1: nombre(s) a buscar ---
        tk.Label(
            self,
            text="Paso 1: Escriba el/los nombre(s) a resaltar",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **padding)

        tk.Label(
            self,
            text="(Si son varios, sepárelos con una coma. Ej: Juan Perez, Maria Lopez)",
            fg="gray30",
        ).pack(anchor="w", padx=16)

        self.entrada_nombres = tk.Entry(self, width=60)
        self.entrada_nombres.pack(padx=16, pady=(4, 12), fill="x")

        # --- Paso 2: seleccionar PDFs ---
        tk.Label(
            self,
            text="Paso 2: Seleccione los archivos PDF",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **padding)

        frame_seleccion = tk.Frame(self)
        frame_seleccion.pack(anchor="w", padx=16, fill="x")

        tk.Button(
            frame_seleccion,
            text="📄 Seleccionar PDF(s)...",
            command=self._seleccionar_pdfs,
        ).pack(side="left")

        self.etiqueta_archivos = tk.Label(frame_seleccion, text="Ningún archivo seleccionado", fg="gray30")
        self.etiqueta_archivos.pack(side="left", padx=10)

        # --- Paso 3: procesar ---
        tk.Label(
            self,
            text="Paso 3: Procesar",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", **padding)

        tk.Button(
            self,
            text="✅ Resaltar nombres",
            bg="#2e7d32",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            command=self._procesar,
        ).pack(padx=16, pady=(0, 12))

        # --- Resultado ---
        tk.Label(self, text="Resultado:", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=16)

        self.caja_resultado = scrolledtext.ScrolledText(self, width=64, height=12, state="disabled")
        self.caja_resultado.pack(padx=16, pady=(4, 16))

    # ------------------------------------------------------------------
    # Acciones del usuario
    # ------------------------------------------------------------------
    def _seleccionar_pdfs(self):
        rutas = filedialog.askopenfilenames(
            title="Seleccione uno o varios PDFs",
            filetypes=[("Archivos PDF", "*.pdf")],
        )
        if rutas:
            self.rutas_pdfs_seleccionados = list(rutas)
            self.etiqueta_archivos.config(
                text=f"{len(rutas)} archivo(s) seleccionado(s)", fg="black"
            )

    def _procesar(self):
        texto_nombres = self.entrada_nombres.get().strip()
        nombres = [n.strip() for n in texto_nombres.split(",") if n.strip()]

        # Validaciones pensadas para un usuario no técnico: mensajes
        # claros, en español, sin jerga ni errores de Python en pantalla.
        if not nombres:
            messagebox.showwarning("Falta información", "Escriba al menos un nombre para buscar.")
            return

        if not self.rutas_pdfs_seleccionados:
            messagebox.showwarning("Falta información", "Seleccione al menos un archivo PDF.")
            return

        carpeta_salida = Path(self.rutas_pdfs_seleccionados[0]).parent / "pdfs_resaltados"
        carpeta_salida.mkdir(exist_ok=True)

        self._limpiar_resultado()

        total_no_encontrados = []

        for ruta_pdf in self.rutas_pdfs_seleccionados:
            nombre_archivo = Path(ruta_pdf).stem
            ruta_salida = carpeta_salida / f"{nombre_archivo}_resaltado.pdf"

            try:
                resultado = resaltar_nombres_en_pdf(ruta_pdf, nombres, str(ruta_salida))
            except Exception as error:
                self._agregar_resultado(f"❌ No se pudo procesar '{Path(ruta_pdf).name}': {error}\n")
                continue

            self._agregar_resultado(f"📄 {resultado.archivo}")
            for nombre, veces in resultado.coincidencias_por_nombre.items():
                if veces > 0:
                    self._agregar_resultado(f"    ✅ '{nombre}' resaltado {veces} vez/veces")
                else:
                    self._agregar_resultado(f"    ⚠️  '{nombre}' NO se encontró en este archivo")
                    total_no_encontrados.append((resultado.archivo, nombre))
            self._agregar_resultado("")

        self._agregar_resultado(f"Listo. Los archivos resaltados están en:\n{carpeta_salida}")

        if total_no_encontrados:
            messagebox.showwarning(
                "Revisar nombres no encontrados",
                "Algunos nombres no se encontraron en algunos archivos. "
                "Revise el detalle en la ventana de resultados (puede ser "
                "que el nombre esté escrito distinto en el PDF).",
            )
        else:
            messagebox.showinfo("Proceso terminado", "Todos los nombres fueron resaltados correctamente.")

    # ------------------------------------------------------------------
    # Utilidades de la caja de resultado (de solo lectura para el usuario)
    # ------------------------------------------------------------------
    def _limpiar_resultado(self):
        self.caja_resultado.config(state="normal")
        self.caja_resultado.delete("1.0", tk.END)
        self.caja_resultado.config(state="disabled")

    def _agregar_resultado(self, texto: str):
        self.caja_resultado.config(state="normal")
        self.caja_resultado.insert(tk.END, texto + "\n")
        self.caja_resultado.config(state="disabled")
        self.caja_resultado.see(tk.END)


if __name__ == "__main__":
    app = VentanaResaltado()
    app.mainloop()
    

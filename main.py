"""
main.py

Punto de entrada de la aplicación. Su única responsabilidad es arrancar
la interfaz gráfica. Así, si en el futuro se agrega, por ejemplo, un modo
de línea de comandos (para automatizar sin GUI), se puede crear sin tocar
main.py.
"""

from interfaz import VentanaResaltado

if __name__ == "__main__":
    app = VentanaResaltado()
    app.mainloop()

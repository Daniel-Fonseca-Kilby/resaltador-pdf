"""Hilos de waitress configurables por PLANILLAS_HILOS (.env). Importar el módulo solo fija variables de entorno de
Tesseract (setdefault), no arranca el servidor ni importa la app."""
import servidor.arrancar as arrancar


def test_hilos_por_defecto_sin_env(monkeypatch):
    monkeypatch.delenv("PLANILLAS_HILOS", raising=False)
    assert arrancar._hilos_por_defecto() == 8


def test_hilos_por_defecto_desde_env(monkeypatch):
    monkeypatch.setenv("PLANILLAS_HILOS", "12")
    assert arrancar._hilos_por_defecto() == 12


def test_hilos_por_defecto_valor_invalido_cae_al_default(monkeypatch):
    monkeypatch.setenv("PLANILLAS_HILOS", "no-es-numero")
    assert arrancar._hilos_por_defecto() == 8


def test_hilos_por_defecto_nunca_menos_de_uno(monkeypatch):
    monkeypatch.setenv("PLANILLAS_HILOS", "0")
    assert arrancar._hilos_por_defecto() == 1

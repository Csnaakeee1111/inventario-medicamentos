#!/usr/bin/env bash
# Abre el programa en Linux con una base de PRUEBA (no la real).
cd "$(dirname "$0")"
export INVENTARIO_DATOS="${INVENTARIO_DATOS:-$HOME/.local/share/InventarioMedicamentos-prueba}"
exec .venv/bin/python main.py

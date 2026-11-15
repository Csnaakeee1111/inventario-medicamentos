"""Puente de actualización con GitHub Releases.

Flujo:
1. GitHub Actions arma `InventarioMedicamentos.exe` cuando se sube una etiqueta v1.2.3
   y lo publica como Release.
2. El programa pregunta a la API de GitHub cuál es la última versión.
3. Si es más nueva, baja el .exe a una carpeta temporal y deja un .bat que espera a que
   el programa se cierre, reemplaza el .exe y lo vuelve a abrir.

Los datos (inventario.db) viven en %APPDATA%, así que actualizar no los toca.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from .version import NOMBRE_EXE, REPO_GITHUB, VERSION


@dataclass
class Novedad:
    version: str
    notas: str
    url_exe: str


def _tupla(v: str) -> tuple[int, ...]:
    partes = []
    for p in v.lstrip("vV").split("."):
        num = "".join(ch for ch in p if ch.isdigit())
        partes.append(int(num or 0))
    return tuple(partes)


def configurado() -> bool:
    return bool(REPO_GITHUB) and "/" in REPO_GITHUB


def es_exe() -> bool:
    return bool(getattr(sys, "frozen", False))


def buscar(timeout: float = 8) -> Novedad | None:
    """Devuelve la versión nueva si existe; None si ya está al día. Lanza OSError sin internet."""
    if not configurado():
        return None
    url = f"https://api.github.com/repos/{REPO_GITHUB}/releases/latest"
    req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json",
                                               "User-Agent": "InventarioMedicamentos"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        datos = json.load(r)
    etiqueta = datos.get("tag_name", "")
    if _tupla(etiqueta) <= _tupla(VERSION):
        return None
    for a in datos.get("assets", []):
        if a.get("name", "").lower() == NOMBRE_EXE.lower():
            return Novedad(etiqueta.lstrip("vV"), datos.get("body") or "", a["browser_download_url"])
    return None


def descargar(novedad: Novedad, progreso=None) -> Path:
    destino = Path(tempfile.gettempdir()) / f"InventarioMedicamentos-{novedad.version}.exe"
    req = urllib.request.Request(novedad.url_exe, headers={"User-Agent": "InventarioMedicamentos"})
    with urllib.request.urlopen(req, timeout=60) as r, open(destino, "wb") as f:
        total = int(r.headers.get("Content-Length") or 0)
        hecho = 0
        while bloque := r.read(256 * 1024):
            f.write(bloque)
            hecho += len(bloque)
            if progreso and total:
                progreso(int(hecho * 100 / total))
    return destino


def instalar_y_reiniciar(nuevo: Path) -> None:
    """Solo en Windows con el .exe: reemplaza el programa al cerrarse y lo reabre."""
    actual = Path(sys.executable)
    bat = Path(tempfile.gettempdir()) / "actualizar_inventario.bat"
    bat.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        "echo Actualizando Inventario de Medicamentos...\r\n"
        f":esperar\r\n"
        f"tasklist /FI \"PID eq {os.getpid()}\" | find \"{os.getpid()}\" >nul\r\n"
        "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto esperar)\r\n"
        f"copy /Y \"{actual}\" \"{actual}.anterior\" >nul\r\n"
        f"copy /Y \"{nuevo}\" \"{actual}\" >nul\r\n"
        f"if errorlevel 1 (echo No se pudo reemplazar el programa. & pause & exit /b 1)\r\n"
        f"start \"\" \"{actual}\"\r\n"
        "del \"%~f0\"\r\n",
        encoding="utf-8")
    subprocess.Popen(["cmd", "/c", str(bat)], creationflags=0x08000000)  # CREATE_NO_WINDOW

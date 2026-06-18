"""Carga de las fuentes Manrope e IBM Plex Mono empaquetadas en assets/fonts."""

import ctypes
import os
import platform
import sys

ASSETS_FONTS = os.path.join(os.path.dirname(__file__), "assets", "fonts")

_FONT_FILES = [
    "Manrope-Regular.ttf",
    "Manrope-Medium.ttf",
    "Manrope-SemiBold.ttf",
    "Manrope-Bold.ttf",
    "Manrope-ExtraBold.ttf",
    "IBMPlexMono-Regular.ttf",
    "IBMPlexMono-Medium.ttf",
    "IBMPlexMono-SemiBold.ttf",
]


_fuentes_cargadas = []  # mantiene referencias vivas para que Tk no las recolecte


def cargar_fuentes_embebidas(root=None):
    """Registra las fuentes .ttf empaquetadas para que Tk pueda usarlas.

    En Windows se registran en el proceso via AddFontResourceEx (no requiere
    instalacion ni permisos de administrador). En Linux/macOS se usa
    tkextrafont si esta disponible; si no, se asume que el sistema ya tiene
    Manrope / IBM Plex Mono instaladas (p. ej. via Google Fonts).
    """
    sistema = platform.system()
    if sistema == "Windows":
        FR_PRIVATE = 0x10
        for nombre in _FONT_FILES:
            ruta = os.path.join(ASSETS_FONTS, nombre)
            if os.path.isfile(ruta):
                ctypes.windll.gdi32.AddFontResourceExW(ruta, FR_PRIVATE, 0)
        return

    try:
        from tkextrafont import Font as ExtraFont

        for nombre in _FONT_FILES:
            ruta = os.path.join(ASSETS_FONTS, nombre)
            if os.path.isfile(ruta):
                _fuentes_cargadas.append(ExtraFont(file=ruta, master=root))
    except ImportError:
        pass

# Lepere Tools

Workbench de escritorio en Python con interfaz gráfica (CustomTkinter). Primera
herramienta: **Excel → CSV** — une todas las hojas de un libro Excel en un
único CSV codificado en UTF-8 con BOM.

El diseño replica el handoff `design_handoff_lepere_tools`: title bar y
sidebar azul oscuro, lienzo gris cálido, tarjetas blancas y acento azul, con
tipografía **Manrope** / **IBM Plex Mono** e iconos **Phosphor Icons**.

## Instalación

```bash
pip install -r requirements.txt
```

## Ejecutar

```bash
python main.py
```

## Estructura

```
main.py                          punto de entrada
src/lepere_tools/
  app.py                         interfaz grafica (CustomTkinter)
  conversion.py                  logica de conversion Excel -> CSV
  fonts.py                       carga de fuentes embebidas
  assets/
    fonts/                       Manrope e IBM Plex Mono (Google Fonts, OFL)
    icons/                       iconos Phosphor Icons rasterizados y tenidos
    logo/                        logo de Lepere Tools
```

## Créditos

- Iconos: [Phosphor Icons](https://phosphoricons.com) ([repositorio](https://github.com/phosphor-icons/homepage)), licencia MIT.
- Fuentes: [Manrope](https://fonts.google.com/specimen/Manrope) e
  [IBM Plex Mono](https://fonts.google.com/specimen/IBM+Plex+Mono), de
  [Google Fonts](https://github.com/google/fonts), licencia OFL (ver
  `src/lepere_tools/assets/fonts/OFL-*.txt`).

## Próximas herramientas

"Unir PDFs" y "Limpiar datos" aparecen en la barra lateral marcadas como
**PRONTO**, listas para implementarse siguiendo el mismo lenguaje visual.

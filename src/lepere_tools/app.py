"""
Lepere Tools - Workbench de escritorio.
Replica del diseno "design_handoff_lepere_tools" (CustomTkinter).

Herramienta activa: Excel -> CSV.
"""

import os
import platform
import queue
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _DND_OK = True
except ImportError:
    _DND_OK = False

from PIL import Image

from . import conversion
from .fonts import cargar_fuentes_embebidas

ASSETS = os.path.join(os.path.dirname(__file__), "assets")
ICONS = os.path.join(ASSETS, "icons")
LOGO_PATH = os.path.join(ASSETS, "logo", "lepere-logo.png")
LOGO_ICO_PATH = os.path.join(ASSETS, "logo", "lepere-logo.ico")

# ---------------------------------------------------------------------------
# Tokens de diseno (ver README del handoff)
# ---------------------------------------------------------------------------

C = {
    "azul_oscuro": "#16283F",
    "azul_borde": "#1F3450",
    "azul_borde_2": "#22384F",
    "activo_bg": "#213E62",
    "acento": "#2E86E0",
    "acento_hover": "#1D6FC4",
    "acento_claro": "#5AA2F0",
    "acento_tint": "#E3EEFB",
    "lienzo": "#EEEAE3",
    "tarjeta": "#FFFFFF",
    "tarjeta_borde": "#DDD8CD",
    "drop_borde": "#C4BDAE",
    "texto": "#1B1A16",
    "texto_muted": "#6F6A60",
    "texto_faint": "#9A948A",
    "side_titulo": "#F0F4FA",
    "side_muted": "#9FB1C8",
    "side_muted_2": "#8197B0",
    "side_label": "#62799A",
    "side_icono": "#7287A3",
    "pronto_borde": "#2D4A68",
    "pronto_texto": "#7E93AD",
    "exito": "#6BA03A",
    "exito_tint": "#E7F1DC",
    "error": "#C0392B",
    "btn_disabled": "#A9C7EC",
    "hover_item_inactivo": "#1F3A5C",
}

FUENTE_UI = "Manrope"
FUENTE_MONO = "IBM Plex Mono"


def _icon(nombre, size=None):
    ruta = os.path.join(ICONS, nombre)
    img = Image.open(ruta)
    if size:
        return ctk.CTkImage(light_image=img, dark_image=img, size=(size, size))
    w, h = img.size
    return ctk.CTkImage(light_image=img, dark_image=img, size=(w // 2, h // 2))


class _BaseTk(ctk.CTk):
    pass


if _DND_OK:
    class _DndCTk(TkinterDnD.DnDWrapper, ctk.CTk):
        def __init__(self, *args, **kwargs):
            ctk.CTk.__init__(self, *args, **kwargs)
            self.TkdndVersion = TkinterDnD._require(self)

    _BaseTk = _DndCTk


class App(_BaseTk):
    def __init__(self):
        super().__init__()
        cargar_fuentes_embebidas(self)
        ctk.set_appearance_mode("light")

        self.title("Lepere Tools")
        self.geometry("860x660")
        self.minsize(780, 600)
        self.configure(fg_color=C["tarjeta_borde"])
        self._set_icono_app()

        self.archivo = None
        self.cola = queue.Queue()
        self.worker = None
        self.worker_meta = None
        self._aviso_pronto_job = None

        self._build_ui()
        self.after(80, self._procesar_cola)

    def _set_icono_app(self):
        """Icono de la ventana y de la barra de tareas (logo Lepere)."""
        try:
            if platform.system() == "Windows" and os.path.isfile(LOGO_ICO_PATH):
                self.iconbitmap(LOGO_ICO_PATH)
            else:
                from PIL import ImageTk

                self._logo_icon_photo = ImageTk.PhotoImage(Image.open(LOGO_PATH))
                self.iconphoto(True, self._logo_icon_photo)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def _build_ui(self):
        marco = ctk.CTkFrame(
            self, fg_color=C["tarjeta"], corner_radius=4,
            border_width=1, border_color=C["tarjeta_borde"],
        )
        marco.pack(fill="both", expand=True, padx=1, pady=1)
        marco.grid_rowconfigure(1, weight=1)
        marco.grid_columnconfigure(0, weight=1)

        self._build_titlebar(marco)

        cuerpo = ctk.CTkFrame(marco, fg_color="transparent", corner_radius=0)
        cuerpo.grid(row=1, column=0, sticky="nsew")
        cuerpo.grid_columnconfigure(1, weight=1)
        cuerpo.grid_rowconfigure(0, weight=1)

        self._build_sidebar(cuerpo)
        self._build_contenido(cuerpo)

    def _build_titlebar(self, parent):
        bar = ctk.CTkFrame(
            parent, fg_color=C["azul_oscuro"], corner_radius=0, height=40,
            border_width=0,
        )
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_propagate(False)

        izq = ctk.CTkFrame(bar, fg_color="transparent")
        izq.pack(side="left", padx=14)
        logo_titlebar = Image.open(LOGO_PATH)
        logo_titlebar_ctk = ctk.CTkImage(
            light_image=logo_titlebar, dark_image=logo_titlebar, size=(18, 18)
        )
        ctk.CTkLabel(izq, image=logo_titlebar_ctk, text="").pack(
            side="left", padx=(0, 8)
        )
        ctk.CTkLabel(
            izq, text="Lepere Tools", text_color=C["side_muted"],
            font=ctk.CTkFont(FUENTE_UI, size=12, weight="bold"),
        ).pack(side="left")
        # Sin controles propios de minimizar/cerrar: se usan los nativos del
        # sistema operativo para evitar tener dos juegos de controles.

    def _build_sidebar(self, parent):
        side = ctk.CTkFrame(
            parent, fg_color=C["azul_oscuro"], corner_radius=0, width=234,
            border_width=0,
        )
        side.grid(row=0, column=0, sticky="nsw")
        side.grid_propagate(False)
        side.grid_columnconfigure(0, weight=1)
        side.grid_rowconfigure(3, weight=1)

        # Marca
        marca = ctk.CTkFrame(side, fg_color="transparent")
        marca.grid(row=0, column=0, sticky="ew", padx=14, pady=(22, 16))
        sep = ctk.CTkFrame(side, fg_color=C["azul_borde_2"], height=1)
        sep.place(in_=marca, relx=0, rely=1.0, relwidth=1.0, y=20)

        logo_img = Image.open(LOGO_PATH)
        logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(40, 40))
        ctk.CTkLabel(marca, image=logo_ctk, text="").grid(row=0, column=0, rowspan=2, padx=(0, 10))
        ctk.CTkLabel(
            marca, text="Lepere", text_color=C["side_titulo"],
            font=ctk.CTkFont(FUENTE_UI, size=15, weight="bold"),
        ).grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(
            marca, text="Tools", text_color=C["side_muted"],
            font=ctk.CTkFont(FUENTE_UI, size=11, weight="bold"),
        ).grid(row=1, column=1, sticky="w")

        # Etiqueta HERRAMIENTAS
        ctk.CTkLabel(
            side, text="HERRAMIENTAS", text_color=C["side_label"],
            font=ctk.CTkFont(FUENTE_UI, size=10, weight="bold"),
        ).grid(row=1, column=0, sticky="w", padx=22, pady=(20, 10))

        herramientas = ctk.CTkFrame(side, fg_color="transparent")
        herramientas.grid(row=2, column=0, sticky="ew", padx=14)
        herramientas.grid_columnconfigure(0, weight=1)

        self._item_activo(herramientas, "file-csv-active.png", "Excel → CSV").grid(
            row=0, column=0, sticky="ew", pady=(0, 4)
        )
        self._item_inactivo(
            herramientas, "file-pdf-inactive.png", "Unir PDFs"
        ).grid(row=1, column=0, sticky="ew", pady=(0, 4))
        self._item_inactivo(
            herramientas, "broom-inactive.png", "Limpiar datos"
        ).grid(row=2, column=0, sticky="ew")

    def _item_activo(self, parent, icono, texto):
        fila = ctk.CTkFrame(parent, fg_color=C["activo_bg"], corner_radius=10)
        fila.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(fila, image=_icon(icono, 19), text="").grid(
            row=0, column=0, padx=(12, 11), pady=11
        )
        ctk.CTkLabel(
            fila, text=texto, text_color=C["side_titulo"],
            font=ctk.CTkFont(FUENTE_UI, size=13, weight="bold"), anchor="w",
        ).grid(row=0, column=1, sticky="w", pady=11)
        return fila

    def _item_inactivo(self, parent, icono, texto):
        fila = ctk.CTkFrame(parent, fg_color="transparent", corner_radius=10, cursor="arrow")
        fila.grid_columnconfigure(1, weight=1)

        def on_enter(_):
            fila.configure(fg_color=C["hover_item_inactivo"])

        def on_leave(_):
            fila.configure(fg_color="transparent")

        def on_click(_):
            self._avisar_pronto(texto)

        widgets = [fila]
        widgets[0].bind("<Enter>", on_enter)
        widgets[0].bind("<Leave>", on_leave)
        widgets[0].bind("<Button-1>", on_click)

        icono_lbl = ctk.CTkLabel(fila, image=_icon(icono, 19), text="", cursor="arrow")
        icono_lbl.grid(row=0, column=0, padx=(12, 11), pady=11)
        texto_lbl = ctk.CTkLabel(
            fila, text=texto, text_color=C["side_muted"], cursor="arrow",
            font=ctk.CTkFont(FUENTE_UI, size=13, weight="normal"), anchor="w",
        )
        texto_lbl.grid(row=0, column=1, sticky="w", pady=11)
        ctk.CTkLabel(
            fila, text="PRONTO", text_color=C["pronto_texto"],
            font=ctk.CTkFont(FUENTE_UI, size=9, weight="bold"),
            corner_radius=5, fg_color="transparent",
        ).grid(row=0, column=2, padx=(0, 12))

        # El icono y el texto tambien deben reaccionar al hover/click, no solo el fondo
        for w in (icono_lbl, texto_lbl):
            w.bind("<Enter>", on_enter)
            w.bind("<Leave>", on_leave)
            w.bind("<Button-1>", on_click)

        return fila

    def _avisar_pronto(self, nombre_herramienta):
        if self._aviso_pronto_job is not None:
            self.after_cancel(self._aviso_pronto_job)
        self._set_estado(C["pronto_texto"], f"'{nombre_herramienta}' estará disponible próximamente")
        self._aviso_pronto_job = self.after(2500, self._restaurar_estado_archivo)

    def _restaurar_estado_archivo(self):
        self._aviso_pronto_job = None
        if self.archivo:
            self._set_estado(C["exito"], "Listo para convertir")
        else:
            self._set_estado(C["texto_muted"], "Selecciona un archivo")

    def _build_contenido(self, parent):
        cont = ctk.CTkFrame(parent, fg_color=C["lienzo"], corner_radius=0)
        cont.grid(row=0, column=1, sticky="nsew")
        cont.grid_columnconfigure(0, weight=1)

        wrap = ctk.CTkFrame(cont, fg_color="transparent")
        wrap.pack(fill="x", padx=40, pady=(36, 0))

        ctk.CTkLabel(
            wrap, text="CONVERSOR", text_color=C["acento"],
            font=ctk.CTkFont(FUENTE_UI, size=11, weight="bold"),
        ).pack(anchor="w", pady=(0, 9))

        titulo = ctk.CTkFrame(wrap, fg_color="transparent")
        titulo.pack(anchor="w")
        ctk.CTkLabel(
            titulo, text="Excel ", text_color=C["texto"],
            font=ctk.CTkFont(FUENTE_UI, size=32, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            titulo, text="→", text_color=C["acento"],
            font=ctk.CTkFont(FUENTE_UI, size=32, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            titulo, text=" CSV", text_color=C["texto"],
            font=ctk.CTkFont(FUENTE_UI, size=32, weight="bold"),
        ).pack(side="left")

        ctk.CTkLabel(
            wrap,
            text="Une todas las hojas de tu libro en un solo archivo CSV\ncodificado en UTF-8 con BOM.",
            text_color=C["texto_muted"], justify="left",
            font=ctk.CTkFont(FUENTE_UI, size=14),
        ).pack(anchor="w", pady=(9, 28))

        # Drop zone
        self.drop_zone = ctk.CTkFrame(
            wrap, fg_color=C["tarjeta"], corner_radius=14,
            border_width=2, border_color=C["drop_borde"],
        )
        self.drop_zone.pack(fill="x", pady=(0, 16))

        interior = ctk.CTkFrame(self.drop_zone, fg_color="transparent")
        interior.pack(pady=32)

        caja_icono = ctk.CTkFrame(
            interior, fg_color=C["acento_tint"], corner_radius=13, width=54, height=54,
        )
        caja_icono.pack(pady=(0, 13))
        caja_icono.pack_propagate(False)
        ctk.CTkLabel(caja_icono, image=_icon("cloud-arrow-up.png", 29), text="").place(
            relx=0.5, rely=0.5, anchor="center"
        )

        lbl_titulo_drop = ctk.CTkLabel(
            interior, text="Arrastra tu archivo Excel aquí", text_color=C["texto"],
            font=ctk.CTkFont(FUENTE_UI, size=15, weight="bold"),
        )
        lbl_titulo_drop.pack()
        lbl_sub_drop = ctk.CTkLabel(
            interior, text="o haz clic para buscar — .xlsx, .xls",
            text_color=C["texto_muted"], font=ctk.CTkFont(FUENTE_UI, size=12),
        )
        lbl_sub_drop.pack()

        drop_widgets = (
            self.drop_zone, interior, caja_icono, lbl_titulo_drop, lbl_sub_drop,
        )
        for widget in drop_widgets:
            widget.bind("<Button-1>", lambda e: self._seleccionar_archivo())
            if _DND_OK:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
                widget.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                widget.dnd_bind("<<DragLeave>>", self._on_drag_leave)

        # Tarjeta de archivo seleccionado (oculta hasta elegir archivo)
        self.archivo_card = ctk.CTkFrame(
            wrap, fg_color=C["tarjeta"], corner_radius=11,
            border_width=1, border_color=C["tarjeta_borde"],
        )
        self.archivo_card.grid_columnconfigure(1, weight=1)

        self.check_box = ctk.CTkFrame(
            self.archivo_card, fg_color=C["exito_tint"], corner_radius=8, width=31, height=31,
        )
        self.check_box.grid(row=0, column=0, padx=(15, 12), pady=13)
        self.check_box.grid_propagate(False)
        ctk.CTkLabel(self.check_box, image=_icon("check.png", 16), text="").place(
            relx=0.5, rely=0.5, anchor="center"
        )

        info = ctk.CTkFrame(self.archivo_card, fg_color="transparent")
        info.grid(row=0, column=1, sticky="ew", pady=13)
        self.lbl_nombre_archivo = ctk.CTkLabel(
            info, text="", text_color=C["texto"], anchor="w",
            font=ctk.CTkFont(FUENTE_MONO, size=12),
        )
        self.lbl_nombre_archivo.pack(anchor="w", fill="x")
        self.lbl_meta_archivo = ctk.CTkLabel(
            info, text="", text_color=C["texto_faint"], anchor="w",
            font=ctk.CTkFont(FUENTE_UI, size=11),
        )
        self.lbl_meta_archivo.pack(anchor="w", fill="x")

        quitar = ctk.CTkLabel(
            self.archivo_card, image=_icon("remove-file.png", 16), text="", cursor="hand2",
        )
        quitar.grid(row=0, column=2, padx=15)
        quitar.bind("<Button-1>", lambda e: self._quitar_archivo())

        # Boton convertir
        self.btn_convertir = ctk.CTkButton(
            wrap, text="→  Convertir a CSV", height=50, corner_radius=12,
            fg_color=C["btn_disabled"], hover_color=C["acento_hover"],
            text_color="white", font=ctk.CTkFont(FUENTE_UI, size=15, weight="bold"),
            command=self._iniciar_conversion, state="disabled",
        )
        self.btn_convertir.pack(fill="x", pady=(0, 0))

        # Barra de progreso (oculta hasta iniciar la conversion)
        self.progress = ctk.CTkProgressBar(
            wrap, height=6, corner_radius=3,
            fg_color=C["tarjeta_borde"], progress_color=C["acento"],
        )
        self.progress.set(0)

        # Estado
        self.estado_fila = ctk.CTkFrame(wrap, fg_color="transparent")
        self.estado_fila.pack(anchor="w", pady=(20, 0))
        self.estado_punto = ctk.CTkLabel(
            self.estado_fila, text="●", text_color=C["exito"],
            font=ctk.CTkFont(FUENTE_MONO, size=12),
        )
        self.estado_punto.pack(side="left", padx=(0, 6))
        self.lbl_estado = ctk.CTkLabel(
            self.estado_fila, text="Selecciona un archivo", text_color=C["texto_muted"],
            font=ctk.CTkFont(FUENTE_MONO, size=11),
        )
        self.lbl_estado.pack(side="left")

    # ------------------------------------------------------------------
    # Drag & drop / seleccion
    # ------------------------------------------------------------------

    def _on_drag_enter(self, event):
        self.drop_zone.configure(border_color=C["acento"], fg_color=C["acento_tint"])

    def _on_drag_leave(self, event):
        self.drop_zone.configure(border_color=C["drop_borde"], fg_color=C["tarjeta"])

    def _on_drop(self, event):
        self.drop_zone.configure(border_color=C["drop_borde"], fg_color=C["tarjeta"])
        rutas = self.tk.splitlist(event.data)
        if rutas:
            self._cargar_archivo(rutas[0])

    def _seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona un archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
        )
        if ruta:
            self._cargar_archivo(ruta)

    def _cargar_archivo(self, ruta):
        ext = os.path.splitext(ruta)[1].lower()
        if ext not in (".xlsx", ".xls"):
            self._set_estado(C["error"], f"Formato no soportado: {ext}")
            return

        self.archivo = ruta
        self.progress.pack_forget()
        self.progress.set(0)

        self.lbl_nombre_archivo.configure(text=os.path.basename(ruta))
        self.lbl_meta_archivo.configure(text="Leyendo información del archivo…")
        self.archivo_card.pack(fill="x", pady=(0, 22), before=self.btn_convertir)

        self.btn_convertir.configure(
            state="disabled", fg_color=C["btn_disabled"], text="→  Convertir a CSV",
        )
        self._set_estado(C["acento"], "Analizando archivo…")

        # La lectura de metadatos (nº de hojas) abre el libro completo; se hace
        # en un hilo aparte para no bloquear la ventana con archivos grandes.
        self.worker_meta = threading.Thread(
            target=self._tarea_metadatos, args=(ruta,), daemon=True
        )
        self.worker_meta.start()

    def _tarea_metadatos(self, ruta):
        try:
            n_hojas, tamano = conversion.contar_hojas_y_tamano(ruta)
            meta = f"{n_hojas} hoja{'s' if n_hojas != 1 else ''} detectadas · {tamano / 1024:.0f} KB"
            self.cola.put(("meta_ok", (ruta, meta)))
        except Exception as e:
            self.cola.put(("meta_error", (ruta, e)))

    def _quitar_archivo(self):
        self.archivo = None
        self.archivo_card.pack_forget()
        self.progress.pack_forget()
        self.btn_convertir.configure(state="disabled", fg_color=C["btn_disabled"])
        self._set_estado(C["texto_muted"], "Selecciona un archivo")

    def _set_estado(self, color, texto):
        texto_color = C["texto_muted"] if color in (C["exito"], C["acento"]) else color
        self.estado_punto.configure(text_color=color)
        self.lbl_estado.configure(text=texto, text_color=texto_color)

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------

    def _iniciar_conversion(self):
        if not self.archivo or (self.worker and self.worker.is_alive()):
            return
        if self.btn_convertir.cget("state") == "disabled":
            return

        self.btn_convertir.configure(state="disabled", text="Convirtiendo…")
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(16, 0), before=self.estado_fila)
        self._set_estado(C["acento"], "Procesando hojas…")

        self.worker = threading.Thread(
            target=self._tarea_conversion, args=(self.archivo,), daemon=True
        )
        self.worker.start()

    def _tarea_conversion(self, ruta):
        def cb(**kwargs):
            self.cola.put(("progreso", kwargs))

        try:
            salida = conversion.convertir(ruta, cb)
            self.cola.put(("ok", salida))
        except Exception as e:
            self.cola.put(("error", e))

    def _procesar_cola(self):
        try:
            while True:
                tipo, payload = self.cola.get_nowait()
                if tipo == "progreso":
                    self._actualizar_progreso(payload)
                elif tipo == "ok":
                    self._finalizar_ok(payload)
                elif tipo == "error":
                    self._finalizar_error(payload)
                elif tipo == "meta_ok":
                    self._aplicar_metadatos(*payload)
                elif tipo == "meta_error":
                    self._aplicar_metadatos_error(*payload)
        except queue.Empty:
            pass
        self.after(80, self._procesar_cola)

    def _aplicar_metadatos(self, ruta, meta):
        if ruta != self.archivo:
            return  # el usuario cambio de archivo mientras se leia este
        self.lbl_meta_archivo.configure(text=meta)
        self.btn_convertir.configure(state="normal", fg_color=C["acento"])
        self._set_estado(C["exito"], "Listo para convertir")

    def _aplicar_metadatos_error(self, ruta, err):
        if ruta != self.archivo:
            return
        self.lbl_meta_archivo.configure(text="No se pudo leer el archivo")
        self._set_estado(C["error"], f"Error: {err}")

    def _actualizar_progreso(self, info):
        fase = info.get("fase")
        hoja = info.get("hoja_actual", 0)
        total = info.get("total_hojas", 1) or 1
        nombre = info.get("nombre_hoja", "")
        filas = info.get("filas", 0)

        if fase == "hoja":
            self.progress.set((hoja - 1) / total)
            self._set_estado(C["acento"], f"Procesando hoja {hoja}/{total}: '{nombre}'…")
        elif fase == "filas":
            self._set_estado(C["acento"], f"Hoja {hoja}/{total} '{nombre}' · {filas:,} filas…")
        elif fase == "hoja_fin":
            self.progress.set(hoja / total)
            self._set_estado(C["acento"], f"Hoja {hoja}/{total} '{nombre}' lista ({filas:,} filas)")

    def _finalizar_ok(self, salida):
        self.progress.set(1)
        self.btn_convertir.configure(
            state="disabled", fg_color=C["exito"], text="✓  Convertido",
        )
        self._set_estado(C["exito"], f"✓ Convertido · guardado en {salida}")
        messagebox.showinfo(
            "Conversión completada",
            f"El archivo se convirtió correctamente.\n\nGuardado en:\n{salida}",
        )

    def _finalizar_error(self, err):
        self.progress.pack_forget()
        self.btn_convertir.configure(state="normal", text="→  Convertir a CSV")
        self._set_estado(C["error"], f"Error: {err}")
        messagebox.showerror("Error en la conversión", str(err))


def main():
    App().mainloop()


if __name__ == "__main__":
    main()

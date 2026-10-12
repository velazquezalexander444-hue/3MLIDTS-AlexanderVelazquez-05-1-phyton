"""Actividad 04 - Formulario de Registro (versión mejorada).

Mejoras principales:
  * Interfaz con ttk, paleta de colores consistente y diseño en dos paneles.
  * Validación en vivo con mensajes de error debajo de cada campo.
  * Restricción de teclas (teléfono solo dígitos, estatura solo números).
  * Rangos realistas (edad 1-120, estatura 0.50-2.50 m).
  * Nombres con acentos, ñ, apóstrofos y guiones.
  * Almacenamiento en CSV (compatible con Excel) con fecha y hora.
  * Tabla con los registros guardados y opción de eliminar.
  * Navegación con Enter y barra de estado.
"""

import csv
import re
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

ARCHIVO = Path(__file__).with_name("registros_3MLIDTS.csv")
ENCABEZADOS = ["Fecha", "Nombre", "Apellidos", "Edad", "Estatura (m)", "Teléfono", "Género"]

# Paleta
COLOR_FONDO = "#F3F4F8"
COLOR_TARJETA = "#FFFFFF"
COLOR_PRIMARIO = "#4F46E5"
COLOR_PRIMARIO_OSCURO = "#4338CA"
COLOR_TEXTO = "#1F2937"
COLOR_SECUNDARIO = "#6B7280"
COLOR_ERROR = "#DC2626"
COLOR_OK = "#16A34A"

LETRAS = "A-Za-zÁÉÍÓÚÜÑáéíóúüñ"
PATRON_TEXTO = re.compile(rf"^[{LETRAS}]+(?:[ '\-][{LETRAS}]+)*$")


# --------------------------------------------------------------------------
# Validaciones: devuelven None si es válido, o un texto con el error
# --------------------------------------------------------------------------
def validar_texto(valor, etiqueta):
    valor = valor.strip()
    if not valor:
        return f"{etiqueta} es obligatorio"
    if len(valor) < 2:
        return f"{etiqueta} es demasiado corto"
    if not PATRON_TEXTO.match(valor):
        return "Solo letras, espacios, guiones y apóstrofos"
    return None


def validar_edad(valor):
    valor = valor.strip()
    if not valor:
        return "La edad es obligatoria"
    if not valor.isdigit():
        return "Ingrese solo números enteros"
    if not 1 <= int(valor) <= 120:
        return "La edad debe estar entre 1 y 120"
    return None


def validar_estatura(valor):
    valor = valor.strip().replace(",", ".")
    if not valor:
        return "La estatura es obligatoria"
    try:
        numero = float(valor)
    except ValueError:
        return "Ingrese un número válido (ej. 1.70)"
    if not 0.50 <= numero <= 2.50:
        return "La estatura debe estar entre 0.50 y 2.50 m"
    return None


def validar_telefono(valor):
    valor = valor.strip()
    if not valor:
        return "El teléfono es obligatorio"
    if not valor.isdigit() or len(valor) != 10:
        return "Debe tener exactamente 10 dígitos"
    return None


# --------------------------------------------------------------------------
# Aplicación
# --------------------------------------------------------------------------
class FormularioRegistro(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Actividad 04 - Formulario de Registro")
        self.geometry("1000x600")
        self.minsize(900, 560)
        self.configure(bg=COLOR_FONDO)

        self.genero = tk.StringVar(value="")
        self.entradas = {}
        self.errores = {}
        self.registros = []

        self._configurar_estilos()
        self._crear_encabezado()
        self._crear_contenido()
        self._crear_barra_estado()
        self._cargar_registros()
        self.entradas["nombre"].focus_set()

    # ---------------------------------------------------------------- estilos
    def _configurar_estilos(self):
        estilo = ttk.Style(self)
        estilo.theme_use("clam")

        estilo.configure("Tarjeta.TFrame", background=COLOR_TARJETA)
        estilo.configure("Fondo.TFrame", background=COLOR_FONDO)
        estilo.configure("Etiqueta.TLabel", background=COLOR_TARJETA,
                         foreground=COLOR_TEXTO, font=("Segoe UI", 10, "bold"))
        estilo.configure("Error.TLabel", background=COLOR_TARJETA,
                         foreground=COLOR_ERROR, font=("Segoe UI", 8))
        estilo.configure("Titulo.TLabel", background=COLOR_TARJETA,
                         foreground=COLOR_TEXTO, font=("Segoe UI", 14, "bold"))
        estilo.configure("TRadiobutton", background=COLOR_TARJETA,
                         foreground=COLOR_TEXTO, font=("Segoe UI", 10))
        estilo.map("TRadiobutton", background=[("active", COLOR_TARJETA)])

        estilo.configure("TEntry", padding=6, fieldbackground="#FFFFFF")

        estilo.configure("Primario.TButton", background=COLOR_PRIMARIO,
                         foreground="white", font=("Segoe UI", 10, "bold"),
                         padding=(14, 8), borderwidth=0)
        estilo.map("Primario.TButton",
                   background=[("active", COLOR_PRIMARIO_OSCURO)])

        estilo.configure("Secundario.TButton", background="#E5E7EB",
                         foreground=COLOR_TEXTO, font=("Segoe UI", 10),
                         padding=(14, 8), borderwidth=0)
        estilo.map("Secundario.TButton", background=[("active", "#D1D5DB")])

        estilo.configure("Treeview", font=("Segoe UI", 9), rowheight=26,
                         background="#FFFFFF", fieldbackground="#FFFFFF")
        estilo.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                         background="#E5E7EB", foreground=COLOR_TEXTO)
        estilo.map("Treeview", background=[("selected", COLOR_PRIMARIO)],
                   foreground=[("selected", "white")])

    # ------------------------------------------------------------ componentes
    def _crear_encabezado(self):
        barra = tk.Frame(self, bg=COLOR_PRIMARIO, height=64)
        barra.pack(fill="x")
        barra.pack_propagate(False)
        tk.Label(barra, text="Formulario de Registro", bg=COLOR_PRIMARIO,
                 fg="white", font=("Segoe UI", 18, "bold")).pack(
            side="left", padx=24)
        tk.Label(barra, text="Actividad 04  ·  3° LIDTS", bg=COLOR_PRIMARIO,
                 fg="#C7D2FE", font=("Segoe UI", 10)).pack(side="right", padx=24)

    def _crear_contenido(self):
        contenedor = ttk.Frame(self, style="Fondo.TFrame")
        contenedor.pack(fill="both", expand=True, padx=20, pady=16)
        contenedor.columnconfigure(0, weight=0)
        contenedor.columnconfigure(1, weight=1)
        contenedor.rowconfigure(0, weight=1)

        self._crear_formulario(contenedor)
        self._crear_tabla(contenedor)

    def _crear_formulario(self, padre):
        tarjeta = ttk.Frame(padre, style="Tarjeta.TFrame", padding=20)
        tarjeta.grid(row=0, column=0, sticky="ns", padx=(0, 16))

        ttk.Label(tarjeta, text="Datos personales",
                  style="Titulo.TLabel").grid(row=0, column=0, columnspan=2,
                                              sticky="w", pady=(0, 10))

        # (clave, etiqueta, validador, ancho, restricción de teclado)
        campos = [
            ("nombre", "Nombre", lambda v: validar_texto(v, "El nombre"), None),
            ("apellidos", "Apellidos",
             lambda v: validar_texto(v, "Los apellidos"), None),
            ("edad", "Edad", validar_edad, self._solo_digitos(3)),
            ("estatura", "Estatura (m)", validar_estatura,
             self._solo_decimal()),
            ("telefono", "Teléfono", validar_telefono, self._solo_digitos(10)),
        ]

        fila = 1
        for clave, etiqueta, validador, restriccion in campos:
            ttk.Label(tarjeta, text=etiqueta, style="Etiqueta.TLabel").grid(
                row=fila, column=0, columnspan=2, sticky="w")
            entrada = ttk.Entry(tarjeta, width=34, font=("Segoe UI", 10))
            if restriccion:
                entrada.configure(validate="key", validatecommand=restriccion)
            entrada.grid(row=fila + 1, column=0, columnspan=2, sticky="ew")

            error = ttk.Label(tarjeta, text=" ", style="Error.TLabel")
            error.grid(row=fila + 2, column=0, columnspan=2, sticky="w",
                       pady=(0, 4))

            entrada.bind("<FocusOut>",
                         lambda e, k=clave, f=validador: self._validar_campo(k, f))
            entrada.bind("<KeyRelease>",
                         lambda e, k=clave: self._limpiar_error(k))
            entrada.bind("<Return>", lambda e: self._siguiente(e.widget))

            self.entradas[clave] = entrada
            self.errores[clave] = error
            self.validadores = getattr(self, "validadores", {})
            self.validadores[clave] = validador
            fila += 3

        # Género
        ttk.Label(tarjeta, text="Género", style="Etiqueta.TLabel").grid(
            row=fila, column=0, columnspan=2, sticky="w")
        marco_genero = ttk.Frame(tarjeta, style="Tarjeta.TFrame")
        marco_genero.grid(row=fila + 1, column=0, columnspan=2, sticky="w",
                          pady=(2, 0))
        for i, texto in enumerate(("Masculino", "Femenino", "No definido")):
            ttk.Radiobutton(marco_genero, text=texto, value=texto,
                            variable=self.genero,
                            command=lambda: self._limpiar_error("genero")
                            ).grid(row=0, column=i, padx=(0, 12))
        self.errores["genero"] = ttk.Label(tarjeta, text=" ",
                                           style="Error.TLabel")
        self.errores["genero"].grid(row=fila + 2, column=0, columnspan=2,
                                    sticky="w", pady=(0, 8))

        # Botones
        botones = ttk.Frame(tarjeta, style="Tarjeta.TFrame")
        botones.grid(row=fila + 3, column=0, columnspan=2, sticky="ew",
                     pady=(8, 0))
        botones.columnconfigure((0, 1), weight=1)
        ttk.Button(botones, text="Guardar", style="Primario.TButton",
                   command=self.guardar).grid(row=0, column=0, sticky="ew",
                                              padx=(0, 6))
        ttk.Button(botones, text="Limpiar", style="Secundario.TButton",
                   command=self.limpiar).grid(row=0, column=1, sticky="ew",
                                              padx=(6, 0))

    def _crear_tabla(self, padre):
        tarjeta = ttk.Frame(padre, style="Tarjeta.TFrame", padding=20)
        tarjeta.grid(row=0, column=1, sticky="nsew")
        tarjeta.rowconfigure(1, weight=1)
        tarjeta.columnconfigure(0, weight=1)

        cabecera = ttk.Frame(tarjeta, style="Tarjeta.TFrame")
        cabecera.grid(row=0, column=0, columnspan=2, sticky="ew",
                      pady=(0, 10))
        cabecera.columnconfigure(0, weight=1)
        ttk.Label(cabecera, text="Registros guardados",
                  style="Titulo.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(cabecera, text="Eliminar seleccionado",
                   style="Secundario.TButton",
                   command=self.eliminar).grid(row=0, column=1)

        anchos = {"Fecha": 120, "Nombre": 100, "Apellidos": 110, "Edad": 50,
                  "Estatura (m)": 85, "Teléfono": 95, "Género": 90}
        self.tabla = ttk.Treeview(tarjeta, columns=ENCABEZADOS,
                                  show="headings", selectmode="browse")
        for col in ENCABEZADOS:
            self.tabla.heading(col, text=col)
            self.tabla.column(col, width=anchos[col], anchor="center")
        self.tabla.grid(row=1, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(tarjeta, orient="vertical",
                               command=self.tabla.yview)
        scroll.grid(row=1, column=1, sticky="ns")
        self.tabla.configure(yscrollcommand=scroll.set)

    def _crear_barra_estado(self):
        self.estado = tk.Label(self, text="Listo", anchor="w", bg="#E5E7EB",
                               fg=COLOR_SECUNDARIO, font=("Segoe UI", 9),
                               padx=12, pady=4)
        self.estado.pack(fill="x", side="bottom")

    # ---------------------------------------------------- restricción de teclas
    def _solo_digitos(self, maximo):
        return (self.register(
            lambda p: p == "" or (p.isdigit() and len(p) <= maximo)), "%P")

    def _solo_decimal(self):
        return (self.register(
            lambda p: re.fullmatch(r"\d?[.,]?\d{0,2}", p) is not None), "%P")

    # ------------------------------------------------------------- validación
    def _validar_campo(self, clave, validador):
        mensaje = validador(self.entradas[clave].get())
        self.errores[clave].configure(text=mensaje or " ")
        return mensaje is None

    def _limpiar_error(self, clave):
        self.errores[clave].configure(text=" ")

    def _validar_todo(self):
        primer_invalido = None
        for clave, validador in self.validadores.items():
            if not self._validar_campo(clave, validador) and primer_invalido is None:
                primer_invalido = clave
        if not self.genero.get():
            self.errores["genero"].configure(text="Seleccione un género")
            if primer_invalido is None:
                primer_invalido = "genero"
        if primer_invalido and primer_invalido != "genero":
            self.entradas[primer_invalido].focus_set()
        return primer_invalido is None

    def _siguiente(self, widget):
        orden = list(self.entradas.values())
        indice = orden.index(widget)
        if indice < len(orden) - 1:
            orden[indice + 1].focus_set()
        else:
            self.guardar()

    # -------------------------------------------------------------- acciones
    def guardar(self):
        if not self._validar_todo():
            self._mensaje_estado("Hay campos con errores", COLOR_ERROR)
            return

        registro = [
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            self.entradas["nombre"].get().strip().title(),
            self.entradas["apellidos"].get().strip().title(),
            int(self.entradas["edad"].get()),
            f"{float(self.entradas['estatura'].get().replace(',', '.')):.2f}",
            self.entradas["telefono"].get().strip(),
            self.genero.get(),
        ]

        try:
            nuevo = not ARCHIVO.exists()
            with open(ARCHIVO, "a", newline="", encoding="utf-8-sig") as f:
                escritor = csv.writer(f)
                if nuevo:
                    escritor.writerow(ENCABEZADOS)
                escritor.writerow(registro)
        except OSError as error:
            messagebox.showerror("Error al guardar",
                                 f"No se pudo escribir el archivo:\n{error}")
            return

        self.registros.append(registro)
        self._refrescar_tabla()
        self.limpiar(mensaje=False)
        self._mensaje_estado("Registro guardado correctamente", COLOR_OK)

    def limpiar(self, mensaje=True):
        for entrada in self.entradas.values():
            entrada.delete(0, tk.END)
        for etiqueta in self.errores.values():
            etiqueta.configure(text=" ")
        self.genero.set("")
        self.entradas["nombre"].focus_set()
        if mensaje:
            self._mensaje_estado("Formulario limpiado", COLOR_SECUNDARIO)

    def eliminar(self):
        seleccion = self.tabla.selection()
        if not seleccion:
            messagebox.showinfo("Eliminar", "Seleccione un registro de la tabla.")
            return
        if not messagebox.askyesno("Confirmar",
                                   "¿Eliminar el registro seleccionado?"):
            return
        del self.registros[int(seleccion[0])]
        self._reescribir_archivo()
        self._refrescar_tabla()
        self._mensaje_estado("Registro eliminado", COLOR_SECUNDARIO)

    # ------------------------------------------------------------ persistencia
    def _cargar_registros(self):
        if ARCHIVO.exists():
            try:
                with open(ARCHIVO, newline="", encoding="utf-8-sig") as f:
                    lector = csv.reader(f)
                    next(lector, None)  # encabezado
                    self.registros = [fila for fila in lector if fila]
            except OSError as error:
                messagebox.showwarning("Aviso",
                                       f"No se pudo leer el archivo:\n{error}")
        self._refrescar_tabla()

    def _reescribir_archivo(self):
        with open(ARCHIVO, "w", newline="", encoding="utf-8-sig") as f:
            escritor = csv.writer(f)
            escritor.writerow(ENCABEZADOS)
            escritor.writerows(self.registros)

    def _refrescar_tabla(self):
        self.tabla.delete(*self.tabla.get_children())
        for i, fila in enumerate(self.registros):
            self.tabla.insert("", "end", iid=str(i), values=fila)
        if self.registros:
            self.tabla.see(str(len(self.registros) - 1))

    def _mensaje_estado(self, texto, color):
        self.estado.configure(text=texto, fg=color)


if __name__ == "__main__":
    FormularioRegistro().mainloop()

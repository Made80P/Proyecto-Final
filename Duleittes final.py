# =========================================================
# UNIVERSIDAD AUTÓNOMA DE OCCIDENTE
# ASIGNATURA: ALGORITMIA Y PROGRAMACIÓN
# PROYECTO FINAL DULEITTES - INTERFAZ BOUTIQUE PASTELERÍA DE AUTOR
# Integrantes: Isabella Moreno Gómez y Madeleine Palacios Baltan
# =========================================================

import os
import shutil
import re
import hashlib
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
from PIL import Image, ImageTk, ImageOps

# ==========================================
# CONSTANTES, ARCHIVOS Y PALETA PASTELERÍA DE AUTOR
# (Azul Arándano, Café Tipo Madera, Mostaza Claro & Crema)
# ==========================================
DIR_BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CATALOGO = os.path.join(DIR_BASE, "catalogo_duleittes.xlsx")
ARCHIVO_CLIENTES = os.path.join(DIR_BASE, "clientes.txt")
CARPETA_IMAGENES = os.path.join(DIR_BASE, "imagenes")
CLAVE_ADMIN_HASH = hashlib.sha256("admin123".encode("utf-8")).hexdigest()

# Paleta Duleittes Pastelería de Autor
COLOR_AZUL_ARANDANO = "#1C3144"   # Azul elegante corporativo (Navbar, títulos destacados, CTA)
COLOR_CAFE_MADERA   = "#4A2E1B"   # Café tipo madera artesanal (Sidebar headers, top bar, acentos)
COLOR_MOSTAZA_CLARO = "#E5B25D"   # Mostaza suave / Champán (Estrellas, badges, hovers, cupones)
COLOR_CREMA_FONDO   = "#FAF6F0"   # Fondo vainilla repostería suave
COLOR_BEIGE_PASTEL  = "#F3E9DC"   # Beige cálido para tarjetas y paneles laterales
COLOR_BLANCO_PURO   = "#FFFFFF"   # Blanco puro para tarjetas e insumos
COLOR_TEXTO_MAIN    = "#2A1E17"   # Texto principal chocolate oscuro
COLOR_TEXTO_MUTED   = "#6C5B51"   # Texto secundario sutil
COLOR_ALERTA_BG     = "#FADBD8"   # Fondo suave stock crítico
COLOR_ALERTA_FG     = "#78281F"   # Texto vino alerta
COLOR_VERDE_EXITO   = "#235B36"   # Verde esmeralda disponible

# Tipografías Pastelería de Autor
FONT_BRAND          = ("Georgia", 24, "bold")
FONT_SLOGAN         = ("Georgia", 11, "italic")
FONT_TITULO_SECTION = ("Georgia", 16, "bold")
FONT_SUBTITULO_SERIF= ("Georgia", 13, "bold")
FONT_TITULO_SERIF   = ("Georgia", 14, "bold")
FONT_CARD_TITULO    = ("Georgia", 12, "bold")
FONT_CARD_PRECIO    = ("Georgia", 13, "bold")
FONT_BODY           = ("Segoe UI", 10)
FONT_BODY_BOLD      = ("Segoe UI", 10, "bold")
FONT_SMALL          = ("Segoe UI", 8)
FONT_SMALL_BOLD     = ("Segoe UI", 8, "bold")


def hash_password(password):
    """Genera un hash SHA-256 seguro para almacenar y verificar contraseñas."""
    return hashlib.sha256(str(password).encode("utf-8")).hexdigest()


def safe_int(val, default=0):
    """Convierte de forma segura cualquier valor a entero."""
    try:
        if pd.isna(val):
            return default
        m = re.search(r"(\d+)", str(val))
        return int(m.group(1)) if m else default
    except Exception:
        return default


def safe_float(val, default=0.0):
    """Convierte de forma segura cualquier valor a flotante."""
    try:
        if pd.isna(val):
            return default
        m = re.search(r"(\d+\.\d+|\d+)", str(val))
        return float(m.group(1)) if m else default
    except Exception:
        return default


def formatear_cop(monto):
    """Formatea valores numéricos al estándar de Pesos Colombianos (COP)."""
    try:
        val = safe_float(monto, 0.0)
        return f"${val:,.0f} COP".replace(",", ".")
    except Exception:
        return "$0 COP"


# ==========================================
# INICIALIZACIÓN Y BASE DE DATOS EXCEL
# ==========================================
def inicializar_sistema():
    if not os.path.exists(CARPETA_IMAGENES):
        os.makedirs(CARPETA_IMAGENES)

    if not os.path.exists(ARCHIVO_CATALOGO):
        ejemplos = [
            {
                "codigo": "D001",
                "producto": "Torta Marquesa de Chocolate Valrhona",
                "categoria": "Tortas",
                "subcategoria": "Mundo de chocolates",
                "stock": 8,
                "precio": 24000,
                "imagen": "D001.png",
                "ingredientes": "Cacao 70% de origen, galleta crujiente, crema Chantilly y avellanas tostadas.",
                "rating": 4.9,
                "badge": "✦ Receta de Autor ✦"
            },
            {
                "codigo": "D002",
                "producto": "Cheesecake de Frutos del Bosque & Miel",
                "categoria": "Postres",
                "subcategoria": "Cheesecakes",
                "stock": 5,
                "precio": 14000,
                "imagen": "D002.png",
                "ingredientes": "Queso crema artesanal, reducción de frambuesas silvestres y base de sablée.",
                "rating": 4.8,
                "badge": "Recomendado"
            },
            {
                "codigo": "D003",
                "producto": "Galletas NyC Red Velvet & Choco Blanco",
                "categoria": "Galletas",
                "subcategoria": "NyC",
                "stock": 14,
                "precio": 8500,
                "imagen": "D003.png",
                "ingredientes": "Masa suave estilo Nueva York, trozos de chocolate blanco belga y vainilla de Papantla.",
                "rating": 5.0,
                "badge": "✦ Más Vendido ✦"
            },
            {
                "codigo": "D004",
                "producto": "Postre Tradicional de Natas & Canela Ceylon",
                "categoria": "Postres",
                "subcategoria": "Tradicionales",
                "stock": 2,
                "precio": 12000,
                "imagen": "D004.png",
                "ingredientes": "Receta familiar con leche fresca de finca, infusión de canela en rama y pasas al ron.",
                "rating": 4.9,
                "badge": "Receta de Casa"
            },
            {
                "codigo": "D005",
                "producto": "Galletas Danesas de Mantequilla & Almendras",
                "categoria": "Galletas",
                "subcategoria": "Danesas",
                "stock": 10,
                "precio": 6500,
                "imagen": "D005.png",
                "ingredientes": "Mantequilla clarificada pura, harina de almendra y azúcar cristalizada.",
                "rating": 4.7,
                "badge": "Clásico Artesanal"
            },
            {
                "codigo": "D006",
                "producto": "Torta Fina Decorada Pistacho & Frambuesa",
                "categoria": "Tortas",
                "subcategoria": "Decoradas",
                "stock": 4,
                "precio": 28000,
                "imagen": "D006.png",
                "ingredientes": "Bizcocho de pistachos italianos, mousse de frutos rojos y acabado satinado.",
                "rating": 5.0,
                "badge": "✦ Alta Repostería ✦"
            },
            {
                "codigo": "D007",
                "producto": "Vasito Marquesa de Nutella & Avellanas",
                "categoria": "Postres",
                "subcategoria": "Marquesas & Vasos",
                "stock": 12,
                "precio": 11000,
                "imagen": "D007.png",
                "ingredientes": "Capas alternadas de crema de avellanas, cacao puro y migas de galleta sablée.",
                "rating": 4.9,
                "badge": "Nuevo Antojo"
            },
            {
                "codigo": "D008",
                "producto": "Galletas Rellenas de Avena, Miel & Chía",
                "categoria": "Galletas",
                "subcategoria": "Rellenas & Avena",
                "stock": 15,
                "precio": 7500,
                "imagen": "D008.png",
                "ingredientes": "Avena integral horneada con miel de abejas pura y relleno cremoso de yogur.",
                "rating": 4.8,
                "badge": "Saludable & Dulce"
            },
            {
                "codigo": "D009",
                "producto": "Torta Familiar Tres Leches & Merengue",
                "categoria": "Tortas",
                "subcategoria": "Para compartir",
                "stock": 6,
                "precio": 22000,
                "imagen": "D009.png",
                "ingredientes": "Biscochuelo ultra suave bañado en infusión de tres leches y merengue suizo tostado.",
                "rating": 4.9,
                "badge": "Para Compartir"
            }
        ]
        df_inicial = pd.DataFrame(ejemplos)
        df_inicial.to_excel(ARCHIVO_CATALOGO, index=False)
    else:
        df = pd.read_excel(ARCHIVO_CATALOGO)
        cols_necesarias = {
            "categoria": "Postres",
            "subcategoria": "Tradicionales",
            "imagen": "",
            "ingredientes": "Ingredientes artesanales de selección especial.",
            "rating": 4.9,
            "badge": "Artesanal"
        }
        modificado = False
        for col, val in cols_necesarias.items():
            if col not in df.columns:
                df[col] = val
                modificado = True
        if modificado:
            df.to_excel(ARCHIVO_CATALOGO, index=False)

    if not os.path.exists(ARCHIVO_CLIENTES):
        with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
            pass

    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    archivo_pedidos_hoy = os.path.join(DIR_BASE, f"pedidos_{fecha_hoy}.txt")
    if not os.path.exists(archivo_pedidos_hoy):
        with open(archivo_pedidos_hoy, "w", encoding="utf-8") as f:
            pass


# ==========================================
# CLASE PRINCIPAL DULEITTES · PASTELERÍA DE AUTOR
# ==========================================
class DuleittesApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Duleittes  · Repostería Artesanal & Postres de Autor")
        self.geometry("1200x820")
        self.minsize(1040, 740)
        self.configure(bg=COLOR_CREMA_FONDO)

        inicializar_sistema()

        self.cliente_activo = None
        self.es_admin = False
        # Estructura: {key_cart: {"codigo_base": str, "nombre": str, "precio": float, "cantidad": int, "stock_max": int}}
        self.carrito = {}
        self.categoria_filtro = "Todos"
        self.subcategoria_filtro = "Todas"
        self.texto_busqueda = ""
        self.card_images = {}
        self.sidebar_visible = True
        self.right_panel = None

        self.configurar_estilos_ttk()
        self.crear_top_announcement_bar()
        self.crear_encabezado_principal()

        self.contenedor = tk.Frame(self, bg=COLOR_CREMA_FONDO)
        self.contenedor.pack(fill="both", expand=True, padx=20, pady=10)

        self.mostrar_vista_cliente_publico()

    def configurar_estilos_ttk(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.style.configure(
            "Treeview",
            background=COLOR_BLANCO_PURO,
            foreground=COLOR_TEXTO_MAIN,
            rowheight=32,
            fieldbackground=COLOR_BLANCO_PURO,
            font=FONT_BODY,
            borderwidth=1,
            relief="solid"
        )
        self.style.configure(
            "Treeview.Heading",
            background=COLOR_AZUL_ARANDANO,
            foreground=COLOR_MOSTAZA_CLARO,
            font=FONT_BODY_BOLD,
            padding=6
        )
        self.style.map("Treeview.Heading", background=[('active', COLOR_AZUL_ARANDANO)])
        self.style.map("Treeview", background=[('selected', COLOR_BEIGE_PASTEL)], foreground=[('selected', COLOR_AZUL_ARANDANO)])

        self.style.configure("TNotebook", background=COLOR_CREMA_FONDO, borderwidth=0)
        self.style.configure(
            "TNotebook.Tab",
            background=COLOR_BEIGE_PASTEL,
            foreground=COLOR_TEXTO_MAIN,
            padding=[20, 8],
            font=FONT_BODY_BOLD
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", COLOR_AZUL_ARANDANO)],
            foreground=[("selected", COLOR_MOSTAZA_CLARO)]
        )

    def crear_top_announcement_bar(self):
        top_bar = tk.Frame(self, bg=COLOR_CAFE_MADERA, height=30)
        top_bar.pack(fill="x", side="top")
        top_bar.pack_propagate(False)

        lbl_promo = tk.Label(
            top_bar,
            text="✦ DULEITTES · PASTELERÍA DE AUTOR · Postres artesanales al alcance de todos · Envíos rápidos a domicilio ✦",
            font=FONT_SMALL_BOLD,
            fg=COLOR_MOSTAZA_CLARO,
            bg=COLOR_CAFE_MADERA
        )
        lbl_promo.pack(expand=True)

    def crear_encabezado_principal(self):
        header_frame = tk.Frame(self, bg=COLOR_CREMA_FONDO, height=85, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
        header_frame.pack(fill="x", side="top", padx=20, pady=(10, 0))

        brand_box = tk.Frame(header_frame, bg=COLOR_CREMA_FONDO)
        brand_box.pack(side="left", padx=20, pady=10)

        lbl_logo = tk.Label(
            brand_box,
            text="DULEITTES",
            font=FONT_BRAND,
            fg=COLOR_AZUL_ARANDANO,
            bg=COLOR_CREMA_FONDO
        )
        lbl_logo.pack(anchor="w")

        lbl_slogan = tk.Label(
            brand_box,
            text="• Duleittes · Pastelería de Autor",
            font=FONT_SLOGAN,
            fg=COLOR_CAFE_MADERA,
            bg=COLOR_CREMA_FONDO
        )
        lbl_slogan.pack(anchor="w")

        self.lbl_usuario_status = tk.Label(
            header_frame,
            text="👤 Modo: Visitante",
            font=FONT_BODY_BOLD,
            fg=COLOR_TEXTO_MUTED,
            bg=COLOR_CREMA_FONDO
        )
        self.lbl_usuario_status.pack(side="right", padx=20)

    def limpiar_contenedor(self):
        for widget in self.contenedor.winfo_children():
            widget.destroy()

    # ==========================================
    # VISTAS PRINCIPALES (CLIENTE & NAVEGACIÓN)
    # ==========================================
    def mostrar_vista_cliente_publico(self):
        self.limpiar_contenedor()
        self.es_admin = False
        self.lbl_usuario_status.config(text="👤 Modo: Visitante")

        nav_bar = tk.Frame(self.contenedor, bg=COLOR_AZUL_ARANDANO, height=50)
        nav_bar.pack(fill="x", pady=(0, 15))

        btn_cat = tk.Button(
            nav_bar,
            text="✦ Ver Catálogo",
            bg=COLOR_AZUL_ARANDANO,
            fg=COLOR_MOSTAZA_CLARO,
            font=FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            command=self.mostrar_vista_cliente_publico
        )
        btn_cat.pack(side="left", padx=15, pady=10)

        btn_login = tk.Button(
            nav_bar,
            text="Iniciar Sesión",
            bg=COLOR_AZUL_ARANDANO,
            fg=COLOR_BLANCO_PURO,
            font=FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            command=self.modal_iniciar_sesion
        )
        btn_login.pack(side="right", padx=15, pady=10)

        btn_reg = tk.Button(
            nav_bar,
            text="Crear Cuenta",
            bg=COLOR_AZUL_ARANDANO,
            fg=COLOR_BLANCO_PURO,
            font=FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            command=self.modal_registrarse
        )
        btn_reg.pack(side="right", padx=5, pady=10)

        btn_admin = tk.Button(
            nav_bar,
            text="Módulo Empresa",
            bg=COLOR_CAFE_MADERA,
            fg=COLOR_MOSTAZA_CLARO,
            font=FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            padx=10,
            command=self.modal_login_admin
        )
        btn_admin.pack(side="right", padx=15, pady=10)

        self.render_layout_tienda(self.contenedor)

    def mostrar_vista_cliente_logueado(self):
        self.limpiar_contenedor()
        self.es_admin = False
        self.lbl_usuario_status.config(text=f"👤 Cliente: {self.cliente_activo['nombre']}")

        notebook = ttk.Notebook(self.contenedor)
        notebook.pack(fill="both", expand=True)

        tab_catalogo = tk.Frame(notebook, bg=COLOR_CREMA_FONDO)
        notebook.add(tab_catalogo, text="  🍰 Tienda Duleittes  ")
        self.render_layout_tienda(tab_catalogo)

        tab_mis_compras = tk.Frame(notebook, bg=COLOR_CREMA_FONDO)
        notebook.add(tab_mis_compras, text="  🛍️ Mis Compras  ")
        self.render_mis_compras(tab_mis_compras)

        frame_usuario_bar = tk.Frame(self.contenedor, bg=COLOR_CREMA_FONDO)
        frame_usuario_bar.pack(fill="x", pady=10)

        btn_eliminar_cuenta = tk.Button(
            frame_usuario_bar,
            text="Eliminar Mi Cuenta",
            bg="#C0392B",
            fg=COLOR_BLANCO_PURO,
            font=FONT_BODY_BOLD,
            relief="flat",
            command=self.eliminar_cuenta_cliente
        )
        btn_eliminar_cuenta.pack(side="left")

        btn_logout = tk.Button(
            frame_usuario_bar,
            text="Cerrar Sesión",
            bg=COLOR_CAFE_MADERA,
            fg=COLOR_MOSTAZA_CLARO,
            font=FONT_BODY_BOLD,
            relief="flat",
            command=self.cerrar_sesion
        )
        btn_logout.pack(side="right")

    def cerrar_sesion(self):
        self.cliente_activo = None
        self.es_admin = False
        self.carrito.clear()
        messagebox.showinfo("Duleittes", "Has cerrado sesión correctamente. ¡Esperamos verte pronto!")
        self.mostrar_vista_cliente_publico()

    def eliminar_cuenta_cliente(self):
        if not self.cliente_activo:
            return
        conf = messagebox.askyesno("Confirmar", "¿Estás seguro de que deseas eliminar tu cuenta de cliente?")
        if conf:
            tel = self.cliente_activo["telefono"]
            if os.path.exists(ARCHIVO_CLIENTES):
                lineas = []
                with open(ARCHIVO_CLIENTES, "r", encoding="utf-8") as f:
                    for l in f:
                        d = l.strip().split("|")
                        if d and d[0] != tel:
                            lineas.append(l)
                with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
                    f.writelines(lineas)
            messagebox.showinfo("Cuenta Eliminada", "Tu cuenta ha sido eliminada con éxito.")
            self.cerrar_sesion()

    # ==========================================
    # MODALES DE AUTENTICACIÓN
    # ==========================================
    def modal_iniciar_sesion(self):
        win = tk.Toplevel(self)
        win.title("Iniciar Sesión · Duleittes")
        win.geometry("400x320")
        win.configure(bg=COLOR_BEIGE_PASTEL)
        win.resizable(False, False)

        tk.Label(win, text="INICIAR SESIÓN", font=FONT_TITULO_SECTION, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO).pack(pady=15)

        tk.Label(win, text="Teléfono (Usuario):", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(anchor="w", padx=40)
        ent_tel = tk.Entry(win, font=FONT_BODY)
        ent_tel.pack(fill="x", padx=40, pady=5)

        tk.Label(win, text="Contraseña:", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(anchor="w", padx=40)
        ent_clave = tk.Entry(win, font=FONT_BODY, show="*")
        ent_clave.pack(fill="x", padx=40, pady=5)

        def ejecutar_login():
            tel = ent_tel.get().strip()
            clave = ent_clave.get().strip()
            clave_cifrada = hash_password(clave)

            if not os.path.exists(ARCHIVO_CLIENTES):
                messagebox.showerror("Error", "No hay usuarios registrados aún.")
                return

            with open(ARCHIVO_CLIENTES, "r", encoding="utf-8") as f:
                for linea in f:
                    datos = linea.strip().split("|")
                    if len(datos) >= 4 and datos[0] == tel:
                        # Soporta tanto contraseñas cifradas (SHA-256) como registros legados previos
                        if datos[3] == clave_cifrada or datos[3] == clave:
                            self.cliente_activo = {"telefono": datos[0], "nombre": datos[1], "correo": datos[2]}
                            messagebox.showinfo("Bienvenido(a)", f"¡Hola de nuevo, {datos[1]}!\nDisfruta de nuestros postres.")
                            win.destroy()
                            self.mostrar_vista_cliente_logueado()
                            return
            messagebox.showerror("Error", "Teléfono o contraseña incorrectos.")

        tk.Button(win, text="Ingresar", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", pady=6, command=ejecutar_login).pack(fill="x", padx=40, pady=20)

    def modal_registrarse(self):
        win = tk.Toplevel(self)
        win.title("Crear Cuenta · Duleittes Bakery")
        win.geometry("420x400")
        win.configure(bg=COLOR_BEIGE_PASTEL)
        win.resizable(False, False)

        tk.Label(win, text="REGISTRO DE CLIENTE", font=FONT_TITULO_SECTION, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO).pack(pady=10)

        fields = [("Teléfono (Usuario):", "tel"), ("Nombre Completo:", "nom"), ("Correo Electrónico:", "cor"), ("Contraseña:", "pas")]
        entries = {}

        for label, key in fields:
            tk.Label(win, text=label, font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(anchor="w", padx=40)
            ent = tk.Entry(win, font=FONT_BODY, show="*" if key == "pas" else "")
            ent.pack(fill="x", padx=40, pady=3)
            entries[key] = ent

        def ejecutar_registro():
            tel = entries["tel"].get().strip()
            nom = entries["nom"].get().strip()
            cor = entries["cor"].get().strip()
            pas = entries["pas"].get().strip()

            if not tel.isdigit():
                messagebox.showerror("Error", "El teléfono debe contener solo números.")
                return

            if not nom or not cor or not pas:
                messagebox.showerror("Error", "Todos los campos son obligatorios.")
                return

            if "@" not in cor or "." not in cor:
                messagebox.showerror("Error", "Ingrese un correo electrónico válido.")
                return

            if os.path.exists(ARCHIVO_CLIENTES):
                with open(ARCHIVO_CLIENTES, "r", encoding="utf-8") as f:
                    for linea in f:
                        datos = linea.strip().split("|")
                        if datos and datos[0] == tel:
                            messagebox.showerror("Error", "Ya existe una cuenta con este número de teléfono.")
                            return

            pas_hash = hash_password(pas)
            with open(ARCHIVO_CLIENTES, "a", encoding="utf-8") as f:
                f.write(f"{tel}|{nom}|{cor}|{pas_hash}\n")

            messagebox.showinfo("Éxito", f"¡Cuenta creada con éxito! Bienvenido(a) {nom}.")
            win.destroy()

        tk.Button(win, text="Crear Cuenta", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", pady=6, command=ejecutar_registro).pack(fill="x", padx=40, pady=18)

    def modal_login_admin(self):
        win = tk.Toplevel(self)
        win.title("Acceso Empresa · Duleittes Bakery")
        win.geometry("360x240")
        win.configure(bg=COLOR_BEIGE_PASTEL)
        win.resizable(False, False)

        tk.Label(win, text="ACCESO ADMINISTRATIVO", font=FONT_TITULO_SECTION, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO).pack(pady=12)
        tk.Label(win, text="Clave de Administración:", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(anchor="w", padx=35)
        ent_clave = tk.Entry(win, font=FONT_BODY, show="*")
        ent_clave.pack(fill="x", padx=35, pady=5)

        def verificar_admin():
            if hash_password(ent_clave.get().strip()) == CLAVE_ADMIN_HASH:
                win.destroy()
                self.es_admin = True
                self.mostrar_vista_administrador()
            else:
                messagebox.showerror("Acceso Denegado", "Clave administrativa incorrecta.")

        tk.Button(win, text="Ingresar al Módulo", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", pady=6, command=verificar_admin).pack(fill="x", padx=35, pady=18)

    # ==========================================
    # HERO BANNER & NAVEGACIÓN COMPLETA DE LA TIENDA
    # ==========================================
    def crear_hero_banner(self, parent):
        banner = tk.Frame(parent, bg=COLOR_BEIGE_PASTEL, height=180, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
        banner.pack(fill="x", pady=(0, 15))
        banner.pack_propagate(False)

        left_content = tk.Frame(banner, bg=COLOR_BEIGE_PASTEL)
        left_content.pack(side="left", fill="both", expand=True, padx=35, pady=20)

        tk.Label(left_content, text="✦ POSTRES FRESCOS & REPOSTERÍA ARTESANAL ✦", font=FONT_SMALL_BOLD, fg=COLOR_AZUL_ARANDANO, bg=COLOR_BEIGE_PASTEL).pack(anchor="w")
        tk.Label(left_content, text="Pequeñas cosas. Buenos momentos.", font=FONT_BRAND, fg=COLOR_CAFE_MADERA, bg=COLOR_BEIGE_PASTEL, justify="left").pack(anchor="w", pady=(4, 4))
        tk.Label(left_content, text="Ingredientes honestos, recetas tradicionales y precios al alcance de todos.\n¡El toque dulce perfecto para acompañar tu día!", font=FONT_BODY, fg=COLOR_TEXTO_MUTED, bg=COLOR_BEIGE_PASTEL, justify="left").pack(anchor="w")

    # ==========================================
    # BARRA LATERAL (SIDEBAR) & CONTENIDO TIENDA
    # ==========================================
    def render_layout_tienda(self, parent):
        container = tk.Frame(parent, bg=COLOR_CREMA_FONDO)
        container.pack(fill="both", expand=True)

        # BARRA SUPERIOR DEL CARRITO
        bar_top_cart = tk.Frame(container, bg=COLOR_CREMA_FONDO)
        bar_top_cart.pack(fill="x", pady=(0, 5))

        if self.cliente_activo:
            cant_total = sum(item["cantidad"] for item in self.carrito.values())
            self.btn_carrito_widget = tk.Button(
                bar_top_cart,
                text=f" 🛒 Mi Carrito Duleittes ({cant_total}) ",
                bg=COLOR_AZUL_ARANDANO,
                fg=COLOR_MOSTAZA_CLARO,
                font=FONT_BODY_BOLD,
                relief="flat",
                bd=0,
                padx=15,
                pady=5,
                command=self.abrir_ventana_carrito
            )
            self.btn_carrito_widget.pack(side="right")

        canvas = tk.Canvas(container, bg=COLOR_CREMA_FONDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg=COLOR_CREMA_FONDO)

        def al_configurar_canvas(event):
            canvas.itemconfig(win_id, width=event.width)

        win_id = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", al_configurar_canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set)

        # Desplazamiento con rueda del mouse o touchpad sin usar la barra
        def al_deslizar_mouse(event):
            if canvas.winfo_exists():
                if getattr(event, "num", None) == 4:
                    canvas.yview_scroll(-1, "units")
                elif getattr(event, "num", None) == 5:
                    canvas.yview_scroll(1, "units")
                elif getattr(event, "delta", 0) != 0:
                    paso = -1 if event.delta > 0 else 1
                    canvas.yview_scroll(paso, "units")

        self.bind_all("<MouseWheel>", al_deslizar_mouse)
        self.bind_all("<Button-4>", al_deslizar_mouse)
        self.bind_all("<Button-5>", al_deslizar_mouse)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.crear_hero_banner(self.scrollable_frame)
        self.crear_seccion_mas_vendido_y_resenas(self.scrollable_frame)

        # CONTENEDOR PRINCIPAL: SIDEBAR A LA IZQUIERDA + GRID DE CATÁLOGO A LA DERECHA
        self.main_shop_split = tk.Frame(self.scrollable_frame, bg=COLOR_CREMA_FONDO)
        self.main_shop_split.pack(fill="both", expand=True, pady=5)

        self.refrescar_grid_productos()

    def crear_sidebar_categorias(self, parent):
        self.sidebar_frame = tk.Frame(parent, bg=COLOR_BEIGE_PASTEL, width=240, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
        self.sidebar_frame.pack(side="left", fill="y", padx=(0, 15), pady=0, anchor="n")
        self.sidebar_frame.pack_propagate(False)
        sidebar = self.sidebar_frame

        # Encabezado del Sidebar
        head_sb = tk.Frame(sidebar, bg=COLOR_CAFE_MADERA, height=40)
        head_sb.pack(fill="x", side="top")
        head_sb.pack_propagate(False)

        tk.Label(head_sb, text="🍰 MENÚ & CATEGORÍAS", font=FONT_BODY_BOLD, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_CAFE_MADERA).pack(side="left", padx=12, expand=True)

        content_sb = tk.Frame(sidebar, bg=COLOR_BEIGE_PASTEL, padx=10, pady=10)
        content_sb.pack(fill="both", expand=True)

        # Jerarquía de Categorías y Subcategorías
        menu_items = [
            ("📌 Todos los productos", "Todos", []),
            ("🧁 Postres", "Postres", ["Marquesas & Vasos", "Cheesecakes", "Tradicionales"]),
            ("🍪 Galletas", "Galletas", ["NyC", "Danesas", "Rellenas & Avena"]),
            ("🎂 Tortas", "Tortas", ["Mundo de chocolates", "Decoradas", "Para compartir"])
        ]

        for label_cat, cat_key, subcats in menu_items:
            es_activa = (self.categoria_filtro == cat_key and self.subcategoria_filtro == "Todas")
            bg_c = COLOR_AZUL_ARANDANO if es_activa else COLOR_BEIGE_PASTEL
            fg_c = COLOR_MOSTAZA_CLARO if es_activa else COLOR_TEXTO_MAIN

            btn_cat = tk.Button(
                content_sb,
                text=label_cat,
                bg=bg_c,
                fg=fg_c,
                font=FONT_BODY_BOLD,
                anchor="w",
                relief="flat",
                bd=0,
                padx=10,
                pady=6,
                command=lambda c=cat_key: self.filtrar_por_categoria(c)
            )
            btn_cat.pack(fill="x", pady=2)

            if subcats:
                sub_frame = tk.Frame(content_sb, bg=COLOR_BEIGE_PASTEL, padx=15)
                sub_frame.pack(fill="x", pady=(0, 4))

                for sub in subcats:
                    es_sub_activa = (self.categoria_filtro == cat_key and self.subcategoria_filtro == sub)
                    bg_s = COLOR_CAFE_MADERA if es_sub_activa else COLOR_BLANCO_PURO
                    fg_s = COLOR_MOSTAZA_CLARO if es_sub_activa else COLOR_TEXTO_MUTED

                    btn_s = tk.Button(
                        sub_frame,
                        text=f"• {sub}",
                        bg=bg_s,
                        fg=fg_s,
                        font=FONT_SMALL_BOLD,
                        anchor="w",
                        relief="flat",
                        bd=1,
                        highlightbackground=COLOR_MOSTAZA_CLARO,
                        padx=8,
                        pady=3,
                        command=lambda c=cat_key, s=sub: self.filtrar_por_cat_and_sub(c, s)
                    )
                    btn_s.pack(fill="x", pady=1)

        # Badge decorativo en la base del sidebar
        box_info_sb = tk.Frame(sidebar, bg=COLOR_AZUL_ARANDANO, padx=10, pady=10)
        box_info_sb.pack(fill="x", side="bottom", padx=8, pady=10)

        tk.Label(box_info_sb, text="✦ 100% ARTESANAL ✦", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack()
        tk.Label(box_info_sb, text="Sin conservantes ni aditivos artificiales.", font=FONT_SMALL, fg=COLOR_BLANCO_PURO, bg=COLOR_AZUL_ARANDANO, justify="center", wraplength=190).pack(pady=(2, 0))

    def filtrar_por_categoria(self, cat):
        self.categoria_filtro = cat
        self.subcategoria_filtro = "Todas"
        self.refrescar_grid_productos()

    def filtrar_por_subcategoria(self, sub):
        self.subcategoria_filtro = sub
        self.refrescar_grid_productos()

    def filtrar_por_cat_and_sub(self, cat, sub):
        self.categoria_filtro = cat
        self.subcategoria_filtro = sub
        self.refrescar_grid_productos()

    def refrescar_grid_productos(self):
        if not hasattr(self, 'main_shop_split') or not self.main_shop_split.winfo_exists():
            return

        for child in list(self.main_shop_split.winfo_children()):
            child.destroy()

        # Renderizar Sidebar a la izquierda
        self.crear_sidebar_categorias(self.main_shop_split)

        # Contenedor derecho del catálogo
        self.right_panel = tk.Frame(self.main_shop_split, bg=COLOR_CREMA_FONDO)
        self.right_panel.pack(side="left", fill="both", expand=True)

        # Barra de búsqueda e indicadores de filtro
        top_filter_bar = tk.Frame(self.right_panel, bg=COLOR_CREMA_FONDO)
        top_filter_bar.pack(fill="x", pady=(0, 10))

        lbl_title = tk.Label(top_filter_bar, text="NUESTROS POSTRES Y REPOSTERÍA", font=FONT_TITULO_SECTION, bg=COLOR_CREMA_FONDO, fg=COLOR_AZUL_ARANDANO)
        lbl_title.pack(side="left")

        # Buscador interactivo
        box_search = tk.Frame(top_filter_bar, bg=COLOR_BLANCO_PURO, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
        box_search.pack(side="right", padx=5)

        tk.Label(box_search, text=" 🔍 ", bg=COLOR_BLANCO_PURO, font=FONT_BODY).pack(side="left")
        self.ent_search = tk.Entry(box_search, font=FONT_BODY, width=25, bd=0, bg=COLOR_BLANCO_PURO)
        self.ent_search.pack(side="left", padx=5, pady=4)
        if self.texto_busqueda:
            self.ent_search.insert(0, self.texto_busqueda)

        def al_buscar(event=None):
            self.texto_busqueda = self.ent_search.get().strip()
            self.refrescar_grid_productos()

        self.ent_search.bind("<KeyRelease>", al_buscar)

        # Indicador de Filtro Activo
        if self.categoria_filtro != "Todos" or self.subcategoria_filtro != "Todas" or self.texto_busqueda:
            bar_active = tk.Frame(self.right_panel, bg=COLOR_BEIGE_PASTEL, padx=10, pady=4)
            bar_active.pack(fill="x", pady=(0, 10))

            txt_filtro = f"Filtro activo: Categoría ➔ '{self.categoria_filtro}'"
            if self.subcategoria_filtro != "Todas":
                txt_filtro += f" | Especialidad ➔ '{self.subcategoria_filtro}'"
            if self.texto_busqueda:
                txt_filtro += f" | Búsqueda ➔ '{self.texto_busqueda}'"

            tk.Label(bar_active, text=txt_filtro, font=FONT_SMALL_BOLD, fg=COLOR_CAFE_MADERA, bg=COLOR_BEIGE_PASTEL).pack(side="left")

            def limpiar_filtros():
                self.categoria_filtro = "Todos"
                self.subcategoria_filtro = "Todas"
                self.texto_busqueda = ""
                self.refrescar_grid_productos()

            btn_clean = tk.Button(bar_active, text=" ✖ Limpiar Filtros ", bg=COLOR_CAFE_MADERA, fg=COLOR_MOSTAZA_CLARO, font=FONT_SMALL_BOLD, bd=0, command=limpiar_filtros)
            btn_clean.pack(side="right")

        # Grid de Productos
        grid_products_frame = tk.Frame(self.right_panel, bg=COLOR_CREMA_FONDO)
        grid_products_frame.pack(fill="x", expand=True)

        if not os.path.exists(ARCHIVO_CATALOGO):
            return

        df = pd.read_excel(ARCHIVO_CATALOGO)
        disponibles = df[df["stock"] > 0] if not df.empty else pd.DataFrame()

        if self.categoria_filtro != "Todos" and "categoria" in disponibles.columns:
            disponibles = disponibles[disponibles["categoria"].astype(str).str.lower() == self.categoria_filtro.lower()]

        if self.subcategoria_filtro != "Todas" and "subcategoria" in disponibles.columns:
            disponibles = disponibles[disponibles["subcategoria"].astype(str).str.lower() == self.subcategoria_filtro.lower()]

        if self.texto_busqueda.strip():
            q = self.texto_busqueda.strip().lower()
            cond_nombre = disponibles["producto"].astype(str).str.lower().str.contains(q)
            cond_ing = disponibles["ingredientes"].astype(str).str.lower().str.contains(q) if "ingredientes" in disponibles.columns else False
            disponibles = disponibles[cond_nombre | cond_ing]

        if disponibles.empty:
            tk.Label(grid_products_frame, text="No encontramos postres que coincidan con la selección actual.", font=FONT_BODY, bg=COLOR_CREMA_FONDO, fg=COLOR_TEXTO_MUTED).pack(pady=30)
            self.crear_footer_informativo(self.scrollable_frame)
            return

        # Dimensiones ampliadas de cada tarjeta de producto
        ANCHO_CARD = 335
        ALTO_CARD = 385
        ALTO_IMG = 195

        def reorganizar_grid(event=None):
            for w in grid_products_frame.winfo_children():
                w.grid_forget()
            width_available = grid_products_frame.winfo_width()
            if width_available <= 1:
                width_available = 750
            num_cols = max(1, width_available // (ANCHO_CARD + 14))

            for idx_card, widget_card in enumerate(grid_products_frame.winfo_children()):
                r = idx_card // num_cols
                c = idx_card % num_cols
                widget_card.grid(row=r, column=c, padx=7, pady=10, sticky="n")

        for idx, (_, row) in enumerate(disponibles.iterrows()):
            cod = str(row['codigo']).strip().upper()
            prod_nombre = str(row['producto'])
            stock_num = safe_int(row['stock'], 0)
            precio_val = safe_float(row['precio'], 0.0)
            img_nombre = str(row['imagen']) if 'imagen' in row and pd.notna(row['imagen']) else ""
            ingredientes_txt = str(row['ingredientes']) if 'ingredientes' in row and pd.notna(row['ingredientes']) else "Ingredientes artesanales seleccionados."
            rating_val = safe_float(row['rating'], 4.9) if 'rating' in row else 4.9
            badge_txt = str(row['badge']) if 'badge' in row and pd.notna(row['badge']) else "✦ Artesanal"

            row_dict = row.to_dict()

            card = tk.Frame(grid_products_frame, bg=COLOR_BLANCO_PURO, width=ANCHO_CARD, height=ALTO_CARD, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
            card.pack_propagate(False)
            card.bind("<Button-1>", lambda e, r_data=row_dict: self.mostrar_detalle_producto(r_data))

            # Fotografía Ampliada
            ruta_img = os.path.join(CARPETA_IMAGENES, img_nombre) if img_nombre else ""
            if ruta_img and os.path.exists(ruta_img):
                try:
                    pil_img = Image.open(ruta_img)
                    pil_img = ImageOps.fit(pil_img, (ANCHO_CARD - 2, ALTO_IMG), Image.Resampling.LANCZOS)
                    tk_img = ImageTk.PhotoImage(pil_img)
                    self.card_images[cod] = tk_img
                    lbl_img = tk.Label(card, image=tk_img, bg=COLOR_BLANCO_PURO)
                except Exception:
                    lbl_img = tk.Label(card, text="[ Fotografía Duleittes ]", bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO)
            else:
                lbl_img = tk.Label(card, text="[ Repostería Artesanal ]", bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO)

            lbl_img.place(x=0, y=0, width=ANCHO_CARD - 2, height=ALTO_IMG)
            lbl_img.bind("<Button-1>", lambda e, r_data=row_dict: self.mostrar_detalle_producto(r_data))

            # Badge de Stock Crítico o Disponible
            if stock_num <= 3:
                lbl_badge = tk.Label(card, text=f" ⚠️ ¡Últimas {stock_num} uds! ", bg=COLOR_ALERTA_BG, fg=COLOR_ALERTA_FG, font=FONT_SMALL_BOLD)
            else:
                lbl_badge = tk.Label(card, text=f" • Disponible · {stock_num} ", bg="#E8F5E9", fg=COLOR_VERDE_EXITO, font=FONT_SMALL)
            lbl_badge.place(x=10, y=10)

            # Badge Destacado Boutique (Derecha)
            lbl_feat = tk.Label(card, text=f" {badge_txt} ", bg=COLOR_CAFE_MADERA, fg=COLOR_MOSTAZA_CLARO, font=FONT_SMALL_BOLD)
            lbl_feat.place(x=ANCHO_CARD - 135, y=10)

            # Estrellas y Título
            lbl_stars = tk.Label(card, text=f"⭐ {rating_val:.1f} / 5.0", font=FONT_SMALL_BOLD, bg=COLOR_BLANCO_PURO, fg="#D4AC0D")
            lbl_stars.place(x=14, y=203)

            lbl_prod = tk.Label(card, text=prod_nombre, font=FONT_CARD_TITULO, bg=COLOR_BLANCO_PURO, fg=COLOR_AZUL_ARANDANO, anchor="nw", justify="left", wraplength=ANCHO_CARD - 28)
            lbl_prod.place(x=14, y=224, width=ANCHO_CARD - 28, height=44)
            lbl_prod.bind("<Button-1>", lambda e, r_data=row_dict: self.mostrar_detalle_producto(r_data))

            lbl_ing = tk.Label(card, text=ingredientes_txt, font=FONT_SMALL, bg=COLOR_BLANCO_PURO, fg=COLOR_TEXTO_MUTED, anchor="nw", justify="left", wraplength=ANCHO_CARD - 28)
            lbl_ing.place(x=14, y=270, width=ANCHO_CARD - 28, height=50)
            lbl_ing.bind("<Button-1>", lambda e, r_data=row_dict: self.mostrar_detalle_producto(r_data))

            lbl_pr = tk.Label(card, text=formatear_cop(precio_val), font=FONT_CARD_PRECIO, bg=COLOR_BLANCO_PURO, fg=COLOR_CAFE_MADERA, anchor="w")
            lbl_pr.place(x=14, y=332, width=220)
            lbl_pr.bind("<Button-1>", lambda e, r_data=row_dict: self.mostrar_detalle_producto(r_data))

            btn_add = tk.Button(
                card,
                text=" + ",
                font=("Segoe UI", 14, "bold"),
                bg=COLOR_AZUL_ARANDANO,
                fg=COLOR_MOSTAZA_CLARO,
                relief="flat",
                bd=0,
                activebackground=COLOR_CAFE_MADERA,
                activeforeground=COLOR_MOSTAZA_CLARO,
                command=lambda r_data=row_dict: self.agregar_al_carrito(r_data)
            )
            btn_add.place(x=ANCHO_CARD - 56, y=325, width=42, height=42)

        grid_products_frame.bind("<Configure>", reorganizar_grid)

        # RENDERIZAR FOOTER INFORMATIVO AL FINAL DE LA PÁGINA
        self.crear_footer_informativo(self.scrollable_frame)

    # ==========================================
    # FOOTER INFORMATIVO Y LEGAL COMPLETO
    # ==========================================
    def crear_footer_informativo(self, parent):
        for child in parent.winfo_children():
            if getattr(child, '_is_footer', False):
                child.destroy()

        footer_main = tk.Frame(parent, bg=COLOR_AZUL_ARANDANO, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
        footer_main.pack(fill="x", pady=(35, 0))
        footer_main._is_footer = True

        # Contenedor Interno de 4 Columnas
        col_container = tk.Frame(footer_main, bg=COLOR_AZUL_ARANDANO, padx=25, pady=25)
        col_container.pack(fill="x", expand=True)

        # COLUMNA 1: Sobre Duleittes Bakery
        col1 = tk.Frame(col_container, bg=COLOR_AZUL_ARANDANO)
        col1.pack(side="left", fill="both", expand=True, padx=10)

        tk.Label(col1, text="DULEITTES · PASTELERÍA DE AUTOR", font=FONT_SUBTITULO_SERIF, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=(0, 6))
        tk.Label(col1, text="• Pequeñas cosas. Buenos momentos.", font=FONT_SLOGAN, fg=COLOR_BEIGE_PASTEL, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=(0, 8))
        tk.Label(
            col1,
            text="Repostería fina artesanal inspirada en las pastelerías boutique del mundo. Precios al alcance de todos con ingredientes 100% naturales.",
            font=FONT_SMALL,
            fg=COLOR_BLANCO_PURO,
            bg=COLOR_AZUL_ARANDANO,
            justify="left",
            wraplength=220
        ).pack(anchor="w")
        tk.Label(col1, text="📍 Av. Cañasgordas # 118-20\nCali, Colombia", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO, justify="left").pack(anchor="w", pady=(10, 0))

        # COLUMNA 2: Atención al Cliente & Contacto
        col2 = tk.Frame(col_container, bg=COLOR_AZUL_ARANDANO)
        col2.pack(side="left", fill="both", expand=True, padx=10)

        tk.Label(col2, text="ATENCIÓN & CONTACTO", font=FONT_SUBTITULO_SERIF, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=(0, 10))
        tk.Label(col2, text="📞 Domicilios: 314 794 8638", font=FONT_BODY_BOLD, fg=COLOR_BLANCO_PURO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=2)
        tk.Label(col2, text="💬 WhatsApp: 317 416 9888", font=FONT_BODY_BOLD, fg=COLOR_BLANCO_PURO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=2)
        tk.Label(col2, text="📧 Correo: pedidos@duleittes.com", font=FONT_SMALL, fg=COLOR_BEIGE_PASTEL, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=2)
        tk.Label(col2, text="✉️ PQRS: servicioalcliente@duleittes.com", font=FONT_SMALL, fg=COLOR_BEIGE_PASTEL, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=2)
        tk.Label(col2, text="🕒 Horarios: Lun-Sáb: 8am-8pm | Dom: 9am-6pm", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=(8, 0))

        # COLUMNA 3: Información Legal & Transparencia
        col3 = tk.Frame(col_container, bg=COLOR_AZUL_ARANDANO)
        col3.pack(side="left", fill="both", expand=True, padx=10)

        tk.Label(col3, text="INFORMACIÓN & LEGAL", font=FONT_SUBTITULO_SERIF, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=(0, 10))

        def mostrar_modal_texto(titulo, contenido):
            w = tk.Toplevel(self)
            w.title(titulo)
            w.geometry("500x400")
            w.configure(bg=COLOR_BEIGE_PASTEL)
            tk.Label(w, text=titulo.upper(), font=FONT_SUBTITULO_SERIF, fg=COLOR_AZUL_ARANDANO, bg=COLOR_BEIGE_PASTEL).pack(pady=12)
            txt_box = tk.Text(w, font=FONT_BODY, wrap="word", padx=15, pady=10, bg=COLOR_BLANCO_PURO)
            txt_box.insert("1.0", contenido)
            txt_box.config(state="disabled")
            txt_box.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        items_legales = [
            ("📜 Términos y Condiciones", "Duleittes Bakery opera conforme a la legislación de comercio colombiano. Todos los precios incluyen impuestos aplicables."),
            ("🔒 Políticas de Privacidad", "Tus datos personales y de contacto se utilizan exclusivamente para despachar tu pedido (Habeas Data Ley 1581)."),
            ("🧾 Facturación Electrónica", "Emitimos factura electrónica por todas tus compras enviada al correo registrado."),
            ("🚚 Envíos y Devoluciones", "Entregas en menos de 90 mins en Cali. Garantía de reemplazo inmediato si el producto llega averiado."),
            ("❓ Preguntas Frecuentes", "¿Tienen opción sin azúcar? Sí, consulta la sección de galletas artesanales de avena. ¿Atienden eventos? Sí, a través de WhatsApp.")
        ]

        for tit_l, doc_l in items_legales:
            btn_l = tk.Button(
                col3,
                text=tit_l,
                font=FONT_SMALL,
                fg=COLOR_BEIGE_PASTEL,
                bg=COLOR_AZUL_ARANDANO,
                activebackground=COLOR_CAFE_MADERA,
                activeforeground=COLOR_MOSTAZA_CLARO,
                anchor="w",
                relief="flat",
                bd=0,
                command=lambda t=tit_l, d=doc_l: mostrar_modal_texto(t, d)
            )
            btn_l.pack(anchor="w", pady=1)

        # COLUMNA 4: Síguenos & Métodos de Pago
        col4 = tk.Frame(col_container, bg=COLOR_AZUL_ARANDANO)
        col4.pack(side="left", fill="both", expand=True, padx=10)

        tk.Label(col4, text="SÍGUENOS & PAGOS", font=FONT_SUBTITULO_SERIF, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=(0, 10))
        tk.Label(col4, text="📸 Instagram: @duleittes", font=FONT_SMALL_BOLD, fg=COLOR_BLANCO_PURO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=2)
        tk.Label(col4, text="👤 Facebook: /DuleittesReposteria", font=FONT_SMALL_BOLD, fg=COLOR_BLANCO_PURO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=2)        
        tk.Label(col4, text="💬 WhatsApp: Chatea con nosotros", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=2)

        tk.Label(col4, text="💳 Medios de Pago Aceptados:", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(anchor="w", pady=(8, 2))
        tk.Label(col4, text="• Nequi  • Daviplata  • Efectivo\n• Tarjetas Crédito & Débito Visa/MC", font=FONT_SMALL, fg=COLOR_BEIGE_PASTEL, bg=COLOR_AZUL_ARANDANO, justify="left").pack(anchor="w")

        # Barra de Copyright Inferior
        bar_copy = tk.Frame(footer_main, bg=COLOR_CAFE_MADERA, height=35)
        bar_copy.pack(fill="x", side="bottom")

        tk.Label(
            bar_copy,
            text="© 2026 DULEITTES · PASTELERÍA DE AUTOR. Todos los derechos reservados.",
            font=FONT_SMALL,
            fg=COLOR_MOSTAZA_CLARO,
            bg=COLOR_CAFE_MADERA
        ).pack(expand=True, pady=6)

    # ==========================================
    # DETALLE DE PRODUCTO & CARRITO
    # ==========================================
    def mostrar_detalle_producto(self, row_data):
        win = tk.Toplevel(self)
        win.title(f"Detalles · {row_data['producto']}")
        win.geometry("520x660")
        win.configure(bg=COLOR_BEIGE_PASTEL)
        win.resizable(False, False)

        img_nombre = str(row_data['imagen']) if 'imagen' in row_data and pd.notna(row_data['imagen']) else ""
        ruta_img = os.path.join(CARPETA_IMAGENES, img_nombre) if img_nombre else ""

        if ruta_img and os.path.exists(ruta_img):
            try:
                pil_img = Image.open(ruta_img)
                pil_img = ImageOps.fit(pil_img, (520, 190), Image.Resampling.LANCZOS)
                tk_det = ImageTk.PhotoImage(pil_img)
                lbl_i = tk.Label(win, image=tk_det, bg=COLOR_BEIGE_PASTEL)
                lbl_i.image = tk_det
                lbl_i.pack(fill="x")
            except Exception:
                tk.Label(win, text="[ Fotografía Duleittes ]", bg="#EAD9A0", height=8).pack(fill="x")
        else:
            tk.Label(win, text="[ Pastelería de Autor ]", bg="#EAD9A0", height=8).pack(fill="x")

        box_info = tk.Frame(win, bg=COLOR_BEIGE_PASTEL, padx=20, pady=10)
        box_info.pack(fill="both", expand=True)

        prod_nombre = str(row_data['producto'])
        cat_nombre = str(row_data['categoria']) if 'categoria' in row_data and pd.notna(row_data['categoria']) else ""
        precio_base = safe_float(row_data['precio'], 0.0)

        tk.Label(box_info, text=prod_nombre, font=FONT_TITULO_SECTION, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO, justify="left").pack(anchor="w")

        ing_text = str(row_data['ingredientes']) if 'ingredientes' in row_data and pd.notna(row_data['ingredientes']) else "Ingredientes artesanales seleccionados."
        lbl_ing = tk.Label(box_info, text=f"Ingredientes & Receta de Casa:\n{ing_text}", font=FONT_BODY, fg=COLOR_TEXTO_MUTED, bg=COLOR_BEIGE_PASTEL, justify="left", wraplength=460)
        lbl_ing.pack(anchor="w", pady=(2, 6))

        # VARIABLES DE OPCIONES
        var_tamano = tk.StringVar(value="Base")
        var_cumple = tk.BooleanVar(value=False)

        es_torta = ("torta" in cat_nombre.lower() or "torta" in prod_nombre.lower())
        es_galleta = ("galleta" in cat_nombre.lower() or "galleta" in prod_nombre.lower())

        # SECCIÓN DE OPCIONES Y TAMAÑO
        frame_opciones = tk.LabelFrame(box_info, text=" Personaliza tu Pedido ", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_CAFE_MADERA, padx=10, pady=6)
        frame_opciones.pack(fill="x", pady=6)

        if es_torta:
            var_tamano.set("15cm")
            tk.Label(frame_opciones, text="Selecciona el tamaño de la Torta:", font=FONT_SMALL_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO).pack(anchor="w")
            
            rb_frame = tk.Frame(frame_opciones, bg=COLOR_BEIGE_PASTEL)
            rb_frame.pack(fill="x", pady=3)

            tk.Radiobutton(rb_frame, text="15 cm (Estándar)", variable=var_tamano, value="15cm", bg=COLOR_BEIGE_PASTEL, font=FONT_SMALL, command=lambda: recalcular()).pack(side="left", padx=4)
            tk.Radiobutton(rb_frame, text="22 cm (+$12.000)", variable=var_tamano, value="22cm", bg=COLOR_BEIGE_PASTEL, font=FONT_SMALL, command=lambda: recalcular()).pack(side="left", padx=4)
            tk.Radiobutton(rb_frame, text="28 cm (+$24.000)", variable=var_tamano, value="28cm", bg=COLOR_BEIGE_PASTEL, font=FONT_SMALL, command=lambda: recalcular()).pack(side="left", padx=4)

            chk_cumple = tk.Checkbutton(
                frame_opciones,
                text="🎂 ¿Es de Cumpleaños? Anexar Topper 'Feliz Día' + Vela Especial (+$6.000 COP)",
                variable=var_cumple,
                bg=COLOR_BEIGE_PASTEL,
                font=FONT_SMALL_BOLD,
                fg=COLOR_CAFE_MADERA,
                command=lambda: recalcular()
            )
            chk_cumple.pack(anchor="w", pady=(4, 2))

        elif es_galleta:
            var_tamano.set("6ud")
            tk.Label(frame_opciones, text="Selecciona la presentación del paquete:", font=FONT_SMALL_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO).pack(anchor="w")
            
            rb_frame = tk.Frame(frame_opciones, bg=COLOR_BEIGE_PASTEL)
            rb_frame.pack(fill="x", pady=3)

            tk.Radiobutton(rb_frame, text="Paquete x 6 Galletas (Estándar)", variable=var_tamano, value="6ud", bg=COLOR_BEIGE_PASTEL, font=FONT_SMALL, command=lambda: recalcular()).pack(side="left", padx=8)
            tk.Radiobutton(rb_frame, text="Paquete x 8 Galletas (+$5.000)", variable=var_tamano, value="8ud", bg=COLOR_BEIGE_PASTEL, font=FONT_SMALL, command=lambda: recalcular()).pack(side="left", padx=8)

        else:
            var_tamano.set("Estandar")
            tk.Label(frame_opciones, text="• Presentación artesanal individual de la casa.", font=FONT_SMALL, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MUTED).pack(anchor="w")

        # ETIQUETA DE PRECIO TOTAL DINÁMICO
        lbl_precio_final = tk.Label(box_info, text="", font=FONT_CARD_PRECIO, bg=COLOR_BEIGE_PASTEL, fg=COLOR_CAFE_MADERA)
        lbl_precio_final.pack(anchor="w", pady=(6, 8))

        def recalcular():
            p_total = precio_base
            detalles_list = []

            if es_torta:
                t = var_tamano.get()
                if t == "22cm":
                    p_total += 12000
                    detalles_list.append("Tamaño 22cm")
                elif t == "28cm":
                    p_total += 24000
                    detalles_list.append("Tamaño 28cm")
                else:
                    detalles_list.append("Tamaño 15cm")

                if var_cumple.get():
                    p_total += 6000
                    detalles_list.append("🎂 Topper & Vela Cumpleaños")

            elif es_galleta:
                g = var_tamano.get()
                if g == "8ud":
                    p_total += 5000
                    detalles_list.append("Paquete x 8 ud")
                else:
                    detalles_list.append("Paquete x 6 ud")

            txt_det = " (" + ", ".join(detalles_list) + ")" if detalles_list else ""
            lbl_precio_final.config(text=f"Precio Total: {formatear_cop(p_total)}{txt_det}")
            return p_total, detalles_list

        recalcular()

        def btn_agregar_accion():
            p_tot, det_list = recalcular()
            win.destroy()
            self.agregar_al_carrito_con_opciones(row_data, p_tot, det_list)

        tk.Button(box_info, text=" 🛒 Agregar al Carrito de Compras ", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", pady=8, command=btn_agregar_accion).pack(fill="x")

    def agregar_al_carrito(self, row_data):
        self.agregar_al_carrito_con_opciones(row_data, safe_float(row_data['precio'], 0.0), [])

    def cantidad_en_carrito_por_codigo_base(self, cod_base):
        """Calcula el total de unidades reservadas en el carrito para un mismo código de inventario."""
        return sum(
            item["cantidad"]
            for item in self.carrito.values()
            if item.get("codigo_base") == cod_base
        )

    def agregar_al_carrito_con_opciones(self, row_data, precio_calculado, detalles_list):
        if not self.cliente_activo:
            messagebox.showinfo("Inicia Sesión", "Para agregar productos al carrito y comprar, por favor inicia sesión o crea una cuenta de cliente.")
            self.modal_iniciar_sesion()
            return

        cod_base = str(row_data['codigo']).strip().upper()
        prod_base = str(row_data['producto'])
        stock_max = safe_int(row_data['stock'], 0)

        var_txt = " (" + ", ".join(detalles_list) + ")" if detalles_list else ""
        key_cart = f"{cod_base}_{var_txt}"
        nombre_completo = f"{prod_base}{var_txt}"

        # Validar stock acumulado considerando todas las variantes del mismo producto
        cant_actual_producto = self.cantidad_en_carrito_por_codigo_base(cod_base)
        if cant_actual_producto + 1 > stock_max:
            messagebox.showwarning("Stock Máximo", f"No puedes agregar más unidades de '{prod_base}'. Stock total disponible: {stock_max}.")
            return

        if key_cart in self.carrito:
            self.carrito[key_cart]["cantidad"] += 1
        else:
            self.carrito[key_cart] = {
                "codigo_base": cod_base,
                "nombre": nombre_completo,
                "precio": precio_calculado,
                "cantidad": 1,
                "stock_max": stock_max
            }

        self.actualizar_boton_carrito()
        messagebox.showinfo("Producto Agregado", f"¡Se agregó '{nombre_completo}' a tu carrito Duleittes!")

    def actualizar_boton_carrito(self):
        cant_total = sum(item["cantidad"] for item in self.carrito.values())
        if hasattr(self, 'btn_carrito_widget') and self.btn_carrito_widget.winfo_exists():
            self.btn_carrito_widget.config(text=f" 🛒 Mi Carrito Duleittes ({cant_total}) ")

    # ==========================================
    # VENTANA DEL CARRITO, CUPONES Y CHECKOUT
    # ==========================================
    def abrir_ventana_carrito(self):
        if not self.carrito:
            messagebox.showinfo("Carrito Vacío", "Tu carrito de compras está vacío actualmente.")
            return

        win = tk.Toplevel(self)
        win.title("Mi Carrito de Compras · Duleittes Bakery")
        win.geometry("640x600")
        win.configure(bg=COLOR_BLANCO_PURO)

        tk.Label(win, text="MI CARRITO DE COMPRAS", font=FONT_TITULO_SECTION, bg=COLOR_BLANCO_PURO, fg=COLOR_AZUL_ARANDANO).pack(pady=15)

        frame_items = tk.Frame(win, bg=COLOR_BLANCO_PURO)
        frame_items.pack(fill="both", expand=True, padx=20)

        cols = ("producto", "precio", "cant", "subtotal")
        tree = ttk.Treeview(frame_items, columns=cols, show="headings", height=7)
        tree.heading("producto", text="Producto Artesanal")
        tree.heading("precio", text="Precio Unit.")
        tree.heading("cant", text="Cant.")
        tree.heading("subtotal", text="Subtotal")
        tree.column("producto", width=250, anchor="w")
        tree.column("precio", width=100, anchor="e")
        tree.column("cant", width=60, anchor="center")
        tree.column("subtotal", width=120, anchor="e")
        tree.pack(fill="both", expand=True)

        # SECCIÓN DE CUPÓN DE DESCUENTO
        frame_cupon = tk.Frame(win, bg=COLOR_BEIGE_PASTEL, padx=10, pady=6)
        frame_cupon.pack(fill="x", padx=20, pady=8)

        tk.Label(frame_cupon, text="🎟️ ¿Tienes un Cupón?:", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(side="left", padx=5)
        ent_cupon = tk.Entry(frame_cupon, font=FONT_BODY, width=15)
        ent_cupon.pack(side="left", padx=5)

        lbl_cupon_msg = tk.Label(frame_cupon, text="", font=FONT_SMALL_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_VERDE_EXITO)
        lbl_cupon_msg.pack(side="left", padx=5)

        descuento_aplicado = [0.0]

        lbl_grand_total = tk.Label(win, text="TOTAL: $0 COP", font=FONT_SUBTITULO_SERIF, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO, pady=8)
        lbl_grand_total.pack(fill="x", padx=20, pady=5)

        def recargar_tabla_carrito():
            tree.delete(*tree.get_children())
            total_general = 0.0
            for cod, item in self.carrito.items():
                subtotal = item["precio"] * item["cantidad"]
                total_general += subtotal
                tree.insert("", "end", iid=cod, values=(item["nombre"], formatear_cop(item["precio"]), item["cantidad"], formatear_cop(subtotal)))

            if descuento_aplicado[0] > 0:
                monto_desc = total_general * descuento_aplicado[0]
                total_final = total_general - monto_desc
                lbl_grand_total.config(text=f"SUBTOTAL: {formatear_cop(total_general)} | DESC: -{formatear_cop(monto_desc)} | TOTAL FINAL: {formatear_cop(total_final)}")
            else:
                lbl_grand_total.config(text=f"TOTAL A PAGAR: {formatear_cop(total_general)}")

        def aplicar_cupon():
            code = ent_cupon.get().strip().upper()
            if code == "DULEITTES10":
                descuento_aplicado[0] = 0.10
                lbl_cupon_msg.config(text="¡10% Descuento Aplicado!", fg=COLOR_VERDE_EXITO)
            elif code == "POSTRE15":
                descuento_aplicado[0] = 0.15
                lbl_cupon_msg.config(text="¡15% Descuento Aplicado!", fg=COLOR_VERDE_EXITO)
            else:
                descuento_aplicado[0] = 0.0
                lbl_cupon_msg.config(text="Cupón Inválido", fg=COLOR_ALERTA_FG)
            recargar_tabla_carrito()

        tk.Button(frame_cupon, text="Aplicar", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_SMALL_BOLD, bd=0, command=aplicar_cupon).pack(side="left", padx=5)

        recargar_tabla_carrito()

        frame_acciones = tk.Frame(win, bg=COLOR_BLANCO_PURO)
        frame_acciones.pack(fill="x", padx=20, pady=5)

        def modificar_cantidad(delta):
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Selección", "Por favor selecciona un producto de la lista.")
                return
            cod = selected[0]
            item = self.carrito[cod]
            cod_base = item.get("codigo_base", cod.split("_")[0])
            nueva_cant = item["cantidad"] + delta

            if nueva_cant <= 0:
                del self.carrito[cod]
            elif delta > 0 and self.cantidad_en_carrito_por_codigo_base(cod_base) + delta > item["stock_max"]:
                messagebox.showwarning("Stock Insuficiente", f"Solo hay {item['stock_max']} unidades disponibles en total de este producto.")
                return
            else:
                item["cantidad"] = nueva_cant

            self.actualizar_boton_carrito()
            if not self.carrito:
                win.destroy()
            else:
                recargar_tabla_carrito()

        def eliminar_item():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Selección", "Por favor selecciona un producto a eliminar.")
                return
            cod = selected[0]
            del self.carrito[cod]
            self.actualizar_boton_carrito()
            if not self.carrito:
                win.destroy()
            else:
                recargar_tabla_carrito()

        tk.Button(frame_acciones, text=" -1 ", bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN, font=FONT_BODY_BOLD, command=lambda: modificar_cantidad(-1)).pack(side="left", padx=2)
        tk.Button(frame_acciones, text=" +1 ", bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN, font=FONT_BODY_BOLD, command=lambda: modificar_cantidad(1)).pack(side="left", padx=2)
        tk.Button(frame_acciones, text=" Eliminar ", bg="#C0392B", fg=COLOR_BLANCO_PURO, font=FONT_BODY_BOLD, command=eliminar_item).pack(side="left", padx=10)

        def ir_a_checkout():
            if not self.carrito:
                return
            pct = descuento_aplicado[0]
            win.destroy()
            self.modal_checkout_envio_pago(pct_descuento=pct)

        tk.Button(win, text=" Proceder al Pago y Envío ➔ ", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", pady=8, command=ir_a_checkout).pack(fill="x", padx=20, pady=15)

    def modal_checkout_envio_pago(self, pct_descuento=0.0):
        win = tk.Toplevel(self)
        win.title("Finalizar Compra · Datos de Envío")
        win.geometry("480x550")
        win.configure(bg=COLOR_BEIGE_PASTEL)
        win.resizable(False, False)

        tk.Label(win, text="DATOS DE ENVÍO Y PAGO", font=FONT_TITULO_SECTION, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO).pack(pady=15)

        total_base = sum(item["precio"] * item["cantidad"] for item in self.carrito.values())
        total_compra = total_base * (1.0 - pct_descuento)

        tk.Label(win, text=f"Total Final a Confirmar: {formatear_cop(total_compra)}", font=FONT_SUBTITULO_SERIF, bg=COLOR_BEIGE_PASTEL, fg=COLOR_CAFE_MADERA).pack(pady=(0, 10))

        frame_form = tk.Frame(win, bg=COLOR_BEIGE_PASTEL, padx=25)
        frame_form.pack(fill="both", expand=True)

        tk.Label(frame_form, text="Dirección de Envío en Cali:", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(anchor="w", pady=(5, 0))
        ent_dir1 = tk.Entry(frame_form, font=FONT_BODY)
        ent_dir1.pack(fill="x", pady=3)

        tk.Label(frame_form, text="Confirmación de Dirección:", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(anchor="w", pady=(5, 0))
        ent_dir2 = tk.Entry(frame_form, font=FONT_BODY)
        ent_dir2.pack(fill="x", pady=3)

        tk.Label(frame_form, text="Método de Pago:", font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).pack(anchor="w", pady=(10, 0))
        combo_pago = ttk.Combobox(frame_form, values=["Efectivo contra entrega", "Nequi / Daviplata", "Tarjeta de Crédito / Débito"], state="readonly", font=FONT_BODY)
        combo_pago.current(0)
        combo_pago.pack(fill="x", pady=3)

        def procesar_pedido_final():
            dir1 = ent_dir1.get().strip()
            dir2 = ent_dir2.get().strip()
            metodo = combo_pago.get()

            if not dir1 or not dir2:
                messagebox.showerror("Campos Requeridos", "Por favor ingresa y confirma la dirección de envío.")
                return

            if dir1.lower() != dir2.lower():
                messagebox.showerror("Error de Confirmación", "Las direcciones ingresadas no coinciden. Verifícalas e inténtalo de nuevo.")
                return

            try:
                df = pd.read_excel(ARCHIVO_CATALOGO)
            except PermissionError:
                messagebox.showerror("Archivo Abierto", "El catálogo Excel está abierto en otro programa. Ciérralo para confirmar tu pedido.")
                return

            df["codigo"] = df["codigo"].astype(str).str.strip().str.upper()

            ahora = datetime.now()
            fecha_hoy = ahora.strftime("%Y-%m-%d")
            hora_act = ahora.strftime("%H:%M:%S")
            archivo_p = os.path.join(DIR_BASE, f"pedidos_{fecha_hoy}.txt")

            # Generar Archivo de Recibo Digital
            nombre_recibo = os.path.join(DIR_BASE, f"recibo_DULEITTES_{ahora.strftime('%Y%m%d_%H%M%S')}.txt")
            lineas_recibo = [
                "===========================================\n",
                "         DULEITTES · PASTELERÍA DE AUTOR - RECIBO DIGITAL \n",
                "   Repostería Artesanal & Postres de Autor  \n",
                "===========================================\n",
                f"Fecha: {fecha_hoy} {hora_act}\n",
                f"Cliente: {self.cliente_activo['nombre']} ({self.cliente_activo['telefono']})\n",
                f"Dirección de Envío: {dir1}\n",
                f"Método de Pago: {metodo}\n",
                "-------------------------------------------\n",
                "CANT | PRODUCTO                      | TOTAL\n",
                "-------------------------------------------\n"
            ]

            with open(archivo_p, "a", encoding="utf-8") as f:
                for key_cart, item in self.carrito.items():
                    cant = item["cantidad"]
                    costo_item = item["precio"] * cant * (1.0 - pct_descuento)
                    cod_real = item.get("codigo_base", key_cart.split("_")[0]).strip().upper()

                    if cod_real in df["codigo"].values:
                        st_actual = safe_int(df.loc[df["codigo"] == cod_real, "stock"].values[0], 0)
                        df.loc[df["codigo"] == cod_real, "stock"] = max(0, st_actual - cant)

                    f.write(f"{self.cliente_activo['telefono']}|{self.cliente_activo['nombre']}|{cod_real}|{item['nombre']}|{cant}|{costo_item:.2f}|{fecha_hoy} {hora_act}|{dir1}|{metodo}\n")
                    lineas_recibo.append(f" {cant:2d}  | {item['nombre'][:28]:28s} | {formatear_cop(costo_item)}\n")

            lineas_recibo.append("-------------------------------------------\n")
            lineas_recibo.append(f"TOTAL PAGADO: {formatear_cop(total_compra)}\n")
            lineas_recibo.append("===========================================\n")
            lineas_recibo.append(" ¡Gracias por elegir Duleittes Bakery! 🧁\n")
            lineas_recibo.append("===========================================\n")

            try:
                df.to_excel(ARCHIVO_CATALOGO, index=False)
            except PermissionError:
                messagebox.showerror("Archivo Abierto", "No se pudo actualizar el inventario porque catalogo_duleittes.xlsx está abierto en Excel.")
                return

            with open(nombre_recibo, "w", encoding="utf-8") as f_rec:
                f_rec.writelines(lineas_recibo)

            self.carrito.clear()
            self.actualizar_boton_carrito()

            win.destroy()
            msg_exito = (
                f"¡Muchas gracias por tu compra en Duleittes Bakery!\n\n"
                f"Tu pedido con destino a:\n'{dir1}'\n"
                f"ha sido confirmado con éxito (Método de Pago: {metodo}).\n\n"
                f"Se ha generado tu comprobante digital:\n'{nombre_recibo}'\n\n"
                f"¡Que disfrutes tu antojo artesanal!"
            )
            messagebox.showinfo("¡Pedido Confirmado!", msg_exito)
            self.mostrar_vista_cliente_logueado()

        tk.Button(win, text=" Confirmar Pedido y Despachar ", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", pady=10, command=procesar_pedido_final).pack(fill="x", padx=30, pady=20)

    def render_mis_compras(self, parent):
        frame = tk.Frame(parent, bg=COLOR_CREMA_FONDO)
        frame.pack(fill="both", expand=True, pady=10)

        cols = ("fecha", "producto", "cantidad", "total")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
        tree.heading("fecha", text="Fecha y Hora")
        tree.heading("producto", text="Producto Artesanal")
        tree.heading("cantidad", text="Cantidad")
        tree.heading("total", text="Total Pagado (COP)")
        tree.column("fecha", width=180, anchor="center")
        tree.column("producto", width=350, anchor="w")
        tree.column("cantidad", width=100, anchor="center")
        tree.column("total", width=180, anchor="e")
        tree.pack(fill="both", expand=True)

        tel = self.cliente_activo["telefono"]
        archivos = [f for f in os.listdir(DIR_BASE) if f.startswith("pedidos_") and f.endswith(".txt")]

        total_acumulado = 0.0
        for arch in sorted(archivos, reverse=True):
            with open(os.path.join(DIR_BASE, arch), "r", encoding="utf-8") as f:
                for linea in f:
                    datos = linea.strip().split("|")
                    if len(datos) >= 7 and datos[0] == tel:
                        monto = safe_float(datos[5], 0.0)
                        total_acumulado += monto
                        tree.insert("", "end", values=(datos[6], datos[3], datos[4], formatear_cop(monto)))

        lbl_total = tk.Label(frame, text=f"Total Acumulado en Compras: {formatear_cop(total_acumulado)}", font=FONT_SUBTITULO_SERIF, fg=COLOR_AZUL_ARANDANO, bg=COLOR_BEIGE_PASTEL, pady=8)
        lbl_total.pack(fill="x", pady=10)

    # ==========================================
    # MÓDULO EMPRESA / ADMINISTRACIÓN
    # ==========================================
    def mostrar_vista_administrador(self):
        self.limpiar_contenedor()
        self.lbl_usuario_status.config(text="⚙️ Modo: Administrador Duleittes")

        notebook = ttk.Notebook(self.contenedor)
        notebook.pack(fill="both", expand=True)

        tab_inv = tk.Frame(notebook, bg=COLOR_CREMA_FONDO)
        notebook.add(tab_inv, text="  📦 Inventario (Excel)  ")
        self.render_admin_inventario(tab_inv)

        tab_rep = tk.Frame(notebook, bg=COLOR_CREMA_FONDO)
        notebook.add(tab_rep, text="  📊 Reporte de Ventas Diarias  ")
        self.render_admin_reportes(tab_rep)

        btn_salir_admin = tk.Button(
            self.contenedor,
            text="Volver a la Tienda Principal",
            bg=COLOR_AZUL_ARANDANO,
            fg=COLOR_MOSTAZA_CLARO,
            font=FONT_BODY_BOLD,
            relief="flat",
            command=self.mostrar_vista_cliente_publico
        )
        btn_salir_admin.pack(anchor="e", pady=10)

    def render_admin_inventario(self, parent):
        df = pd.read_excel(ARCHIVO_CATALOGO)
        cols = ("codigo", "producto", "categoria", "subcategoria", "stock", "precio", "imagen", "ingredientes")

        tree = ttk.Treeview(parent, columns=cols, show="headings", height=7)
        tree.heading("codigo", text="Código")
        tree.heading("producto", text="Producto Artesanal")
        tree.heading("categoria", text="Categoría")
        tree.heading("subcategoria", text="Subcategoría")
        tree.heading("stock", text="Stock")
        tree.heading("precio", text="Precio (COP)")
        tree.heading("imagen", text="Archivo Foto")
        tree.heading("ingredientes", text="Ingredientes / Notas")

        tree.column("codigo", width=65, anchor="center")
        tree.column("producto", width=190, anchor="w")
        tree.column("categoria", width=90, anchor="center")
        tree.column("subcategoria", width=110, anchor="center")
        tree.column("stock", width=55, anchor="center")
        tree.column("precio", width=100, anchor="e")
        tree.column("imagen", width=95, anchor="center")
        tree.column("ingredientes", width=210, anchor="w")
        tree.pack(fill="both", expand=True, pady=5)

        # Tag para resaltar stock crítico en rojo
        tree.tag_configure("stock_critico", background=COLOR_ALERTA_BG, foreground=COLOR_ALERTA_FG)

        for _, row in df.iterrows():
            nombre_img = str(row['imagen']) if 'imagen' in row and pd.notna(row['imagen']) else ""
            ing = str(row['ingredientes']) if 'ingredientes' in row and pd.notna(row['ingredientes']) else ""
            cat_val = str(row['categoria']) if 'categoria' in row and pd.notna(row['categoria']) else "Postres"
            sub_val = str(row['subcategoria']) if 'subcategoria' in row and pd.notna(row['subcategoria']) else "Tradicionales"
            st_val = safe_int(row['stock'], 0)

            tag_aplicada = ("stock_critico",) if st_val <= 3 else ()

            tree.insert("", "end", values=(
                str(row['codigo']),
                row['producto'],
                cat_val,
                sub_val,
                st_val,
                formatear_cop(row['precio']),
                nombre_img,
                ing
            ), tags=tag_aplicada)

        box_edit = tk.LabelFrame(parent, text=" Formulario de Inventario (Excel) ", font=FONT_SUBTITULO_SERIF, bg=COLOR_BEIGE_PASTEL, fg=COLOR_AZUL_ARANDANO, padx=10, pady=8)
        box_edit.pack(fill="x", pady=5)

        entries = {}
        campos = [
            ("Código:", "cod", 0, 0),
            ("Nombre Producto:", "prod", 0, 2),
            ("Categoría:", "cat", 1, 0),
            ("Subcategoría:", "sub", 1, 2),
            ("Stock:", "stock", 2, 0),
            ("Precio (COP):", "precio", 2, 2),
            ("Foto Producto:", "img", 3, 0),
            ("Ingredientes:", "ing", 3, 2)
        ]

        for label, key, r, c in campos:
            tk.Label(box_edit, text=label, font=FONT_BODY_BOLD, bg=COLOR_BEIGE_PASTEL, fg=COLOR_TEXTO_MAIN).grid(row=r, column=c, padx=4, pady=3, sticky="e")
            ent = tk.Entry(box_edit, font=FONT_BODY)
            ent.grid(row=r, column=c+1, padx=4, pady=3)
            entries[key] = ent

        def al_seleccionar_inventario(event):
            selected = tree.selection()
            if selected:
                item = tree.item(selected[0])
                vals = item['values']
                entries["cod"].delete(0, tk.END); entries["cod"].insert(0, str(vals[0]))
                entries["prod"].delete(0, tk.END); entries["prod"].insert(0, str(vals[1]))
                entries["cat"].delete(0, tk.END); entries["cat"].insert(0, str(vals[2]))
                entries["sub"].delete(0, tk.END); entries["sub"].insert(0, str(vals[3]))
                entries["stock"].delete(0, tk.END); entries["stock"].insert(0, str(vals[4]))
                precio_num = str(vals[5]).replace("$", "").replace(" COP", "").replace(".", "").strip()
                entries["precio"].delete(0, tk.END); entries["precio"].insert(0, precio_num)
                entries["img"].delete(0, tk.END); entries["img"].insert(0, str(vals[6]))
                entries["ing"].delete(0, tk.END); entries["ing"].insert(0, str(vals[7]))

        tree.bind("<<TreeviewSelect>>", al_seleccionar_inventario)

        def seleccionar_imagen():
            if not self.es_admin:
                messagebox.showerror("Acceso Denegado", "Solo el administrador puede subir imágenes.")
                return
            ruta_origen = filedialog.askopenfilename(
                title="Seleccionar fotografía del postre (Admin)",
                filetypes=[("Archivos de Imagen", "*.png *.jpg *.jpeg *.webp")]
            )
            if ruta_origen:
                nombre_arch = os.path.basename(ruta_origen)
                ruta_dest = os.path.join(CARPETA_IMAGENES, nombre_arch)
                try:
                    shutil.copy(ruta_origen, ruta_dest)
                    entries["img"].delete(0, tk.END)
                    entries["img"].insert(0, nombre_arch)
                    messagebox.showinfo("Imagen Guardada", f"Fotografía '{nombre_arch}' copiada con éxito a la carpeta '{CARPETA_IMAGENES}'.")
                except Exception as ex:
                    messagebox.showerror("Error", f"No se pudo copiar la imagen: {ex}")

        btn_img = tk.Button(box_edit, text="Buscar Foto...", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_SMALL_BOLD, bd=0, command=seleccionar_imagen)
        btn_img.grid(row=3, column=4, padx=5, pady=3)

        def guardar_producto():
            if not self.es_admin:
                messagebox.showerror("Acceso Denegado", "Operación no permitida.")
                return

            c = entries["cod"].get().strip().upper()
            p = entries["prod"].get().strip()
            cat_v = entries["cat"].get().strip() or "Postres"
            sub_v = entries["sub"].get().strip() or "Tradicionales"
            im = entries["img"].get().strip()
            ing_val = entries["ing"].get().strip()

            raw_stock = entries["stock"].get().strip()
            raw_precio = entries["precio"].get().strip().replace("$", "").replace("COP", "").replace("cop", "").replace(".", "").replace(",", "").strip()

            try:
                s = int(raw_stock)
                pr = float(raw_precio)
            except ValueError:
                messagebox.showerror("Error de Formato", "Stock y precio deben ser valores numéricos válidos (ej: Stock=10, Precio=24000).")
                return

            if not c or not p:
                messagebox.showerror("Error", "Código y nombre de producto son obligatorios.")
                return

            try:
                df_curr = pd.read_excel(ARCHIVO_CATALOGO)
            except Exception as ex:
                messagebox.showerror("Error al Leer Excel", f"No se pudo abrir '{ARCHIVO_CATALOGO}': {ex}")
                return

            df_curr["codigo"] = df_curr["codigo"].astype(str).str.strip().str.upper()

            if not df_curr.empty and c in df_curr["codigo"].values:
                df_curr.loc[df_curr["codigo"] == c, ["producto", "categoria", "subcategoria", "stock", "precio", "imagen", "ingredientes"]] = [p, cat_v, sub_v, s, pr, im, ing_val]
                msg = f"¡Producto '{c}' ({p}) actualizado con éxito en Excel!"
            else:
                nuevo = pd.DataFrame({"codigo": [c], "producto": [p], "categoria": [cat_v], "subcategoria": [sub_v], "stock": [s], "precio": [pr], "imagen": [im], "ingredientes": [ing_val], "rating": [4.9], "badge": ["Artesanal"]})
                df_curr = pd.concat([df_curr, nuevo], ignore_index=True)
                msg = f"¡Producto '{p}' registrado con éxito en el catálogo!"

            try:
                df_curr.to_excel(ARCHIVO_CATALOGO, index=False)
                messagebox.showinfo("Éxito al Guardar", msg)
                self.mostrar_vista_administrador()
            except PermissionError:
                messagebox.showerror("Archivo Abierto en Excel", "No se pudo guardar porque catalogo_duleittes.xlsx esta abierto en Microsoft Excel. Por favor cierralo e intenta de nuevo.")
            except Exception as ex:
                messagebox.showerror("Error al Guardar", f"No se pudo guardar en Excel: {ex}")

        def eliminar_producto():
            if not self.es_admin:
                messagebox.showerror("Acceso Denegado", "Operación no permitida.")
                return

            c = entries["cod"].get().strip().upper()
            if not c:
                messagebox.showerror("Error", "Seleccione o ingrese el código del producto a eliminar.")
                return

            df_curr = pd.read_excel(ARCHIVO_CATALOGO)
            df_curr["codigo"] = df_curr["codigo"].astype(str).str.strip().str.upper()

            if c not in df_curr["codigo"].values:
                messagebox.showerror("Error", f"El código '{c}' no existe en el inventario.")
                return

            confirmacion = messagebox.askyesno("Confirmar Eliminación", f"¿Está seguro de eliminar '{c}' del inventario?")
            if confirmacion:
                df_curr = df_curr[df_curr["codigo"] != c]
                try:
                    df_curr.to_excel(ARCHIVO_CATALOGO, index=False)
                    messagebox.showinfo("Éxito", f"Producto '{c}' eliminado correctamente del archivo Excel.")
                    self.mostrar_vista_administrador()
                except PermissionError:
                    messagebox.showerror("Archivo Abierto en Excel", "No se pudo guardar porque catalogo_duleittes.xlsx esta abierto en Microsoft Excel. Por favor cierralo e intenta de nuevo.")
                except Exception as ex:
                    messagebox.showerror("Error al Guardar", f"No se pudo actualizar el archivo Excel: {ex}")

        btn_guardar = tk.Button(box_edit, text=" Guardar en Inventario ", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", command=guardar_producto)
        btn_guardar.grid(row=4, column=1, padx=10, pady=8)

        btn_eliminar = tk.Button(box_edit, text=" Eliminar Producto ", bg="#C0392B", fg=COLOR_BLANCO_PURO, font=FONT_BODY_BOLD, relief="flat", command=eliminar_producto)
        btn_eliminar.grid(row=4, column=3, padx=10, pady=8)

    def render_admin_reportes(self, parent):
        top_frame = tk.Frame(parent, bg=COLOR_CREMA_FONDO)
        top_frame.pack(fill="x", pady=10)

        tk.Label(top_frame, text="Consultar Fecha (YYYY-MM-DD):", font=FONT_BODY_BOLD, bg=COLOR_CREMA_FONDO, fg=COLOR_AZUL_ARANDANO).pack(side="left", padx=5)
        ent_fecha = tk.Entry(top_frame, font=FONT_BODY, width=15)
        ent_fecha.insert(0, datetime.now().strftime("%Y-%m-%d"))
        ent_fecha.pack(side="left", padx=5)

        cols = ("hora", "cliente", "tel", "producto", "cant", "monto")
        tree = ttk.Treeview(parent, columns=cols, show="headings", height=8)
        tree.heading("hora", text="Fecha/Hora")
        tree.heading("cliente", text="Cliente")
        tree.heading("tel", text="Teléfono")
        tree.heading("producto", text="Producto Artesanal")
        tree.heading("cant", text="Cant.")
        tree.heading("monto", text="Total (COP)")

        tree.column("hora", width=140, anchor="center")
        tree.column("cliente", width=180, anchor="w")
        tree.column("tel", width=100, anchor="center")
        tree.column("producto", width=220, anchor="w")
        tree.column("cant", width=60, anchor="center")
        tree.column("monto", width=120, anchor="e")
        tree.pack(fill="both", expand=True)

        lbl_stats = tk.Label(parent, text="TOTAL VENTAS DEL DÍA: $0 COP", font=FONT_SUBTITULO_SERIF, fg=COLOR_AZUL_ARANDANO, bg=COLOR_BEIGE_PASTEL, pady=8)
        lbl_stats.pack(fill="x", pady=10)

        def cargar_reporte():
            tree.delete(*tree.get_children())
            fecha = ent_fecha.get().strip()
            archivo = os.path.join(DIR_BASE, f"pedidos_{fecha}.txt")

            if not os.path.exists(archivo):
                lbl_stats.config(text=f"No hay ventas registradas para la fecha: {fecha}")
                return

            total_dia = 0.0
            num_pedidos = 0
            conteo_prod = {}

            with open(archivo, "r", encoding="utf-8") as f:
                for linea in f:
                    d = linea.strip().split("|")
                    if len(d) >= 7:
                        monto = safe_float(d[5], 0.0)
                        cant = safe_int(d[4], 1)
                        prod_n = d[3]

                        total_dia += monto
                        num_pedidos += 1
                        conteo_prod[prod_n] = conteo_prod.get(prod_n, 0) + cant

                        tree.insert("", "end", values=(d[6], d[1], d[0], prod_n, cant, formatear_cop(monto)))

            prod_top = max(conteo_prod, key=conteo_prod.get) if conteo_prod else "N/A"
            lbl_stats.config(text=f"REPORTE ({fecha}) ➔ TOTAL: {formatear_cop(total_dia)}  |  PEDIDOS: {num_pedidos}  |  TOP VENDIDO: '{prod_top}'")

        tk.Button(top_frame, text=" Consultar Ventas ", bg=COLOR_AZUL_ARANDANO, fg=COLOR_MOSTAZA_CLARO, font=FONT_BODY_BOLD, relief="flat", command=cargar_reporte).pack(side="left", padx=10)
        cargar_reporte()

    def toggle_sidebar(self):
        if hasattr(self, 'sidebar_frame') and self.sidebar_frame and self.sidebar_frame.winfo_exists():
            if self.sidebar_visible:
                self.sidebar_frame.pack_forget()
                self.sidebar_visible = False
            else:
                if self.right_panel and self.right_panel.winfo_exists():
                    self.sidebar_frame.pack(side="left", fill="y", padx=(0, 15), anchor="n", before=self.right_panel)
                else:
                    self.sidebar_frame.pack(side="left", fill="y", padx=(0, 15), anchor="n")
                self.sidebar_visible = True

    def crear_seccion_mas_vendido_y_resenas(self, parent):
        section_frame = tk.Frame(parent, bg=COLOR_AZUL_ARANDANO, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
        section_frame.pack(fill="x", pady=(0, 15), padx=2)

        # Header de la sección
        header_box = tk.Frame(section_frame, bg=COLOR_CAFE_MADERA, height=38)
        header_box.pack(fill="x", side="top")
        header_box.pack_propagate(False)

        tk.Label(
            header_box,
            text="🔥 LO MÁS VENDIDO DEL MES · SELECCIÓN DE AUTOR & RESEÑAS DE CLIENTES 🔥",
            font=FONT_BODY_BOLD,
            fg=COLOR_MOSTAZA_CLARO,
            bg=COLOR_CAFE_MADERA
        ).pack(expand=True)

        body_box = tk.Frame(section_frame, bg=COLOR_AZUL_ARANDANO, padx=15, pady=12)
        body_box.pack(fill="x")

        # Datos del Producto Más Vendido (#1 Torta Marquesa D001)
        row_mas_vendido = {
            "codigo": "D001",
            "producto": "Torta de Chocolate Semiamargo",
            "categoria": "Tortas",
            "subcategoria": "Mundo de chocolates",
            "stock": 8,
            "precio": 24000,
            "imagen": "D001.png",
            "ingredientes": "Cacao 70% de origen, galleta crujiente, crema Chantilly y avellanas tostadas.",
            "rating": 4.9,
            "badge": "✦ Edición Especial ✦"
        }
        if os.path.exists(ARCHIVO_CATALOGO):
            try:
                df_cat = pd.read_excel(ARCHIVO_CATALOGO)
                if not df_cat.empty:
                    match_mv = df_cat[df_cat["codigo"].astype(str).str.strip().str.upper() == "D001"]
                    if not match_mv.empty:
                        row_mas_vendido = match_mv.iloc[0].to_dict()
            except Exception:
                pass

        def abrir_compra_mas_vendido(event=None):
            self.mostrar_detalle_producto(row_mas_vendido)

        # LADO IZQUIERDO: Producto Estrella Conectado a Especificaciones de Compra
        left_star = tk.Frame(body_box, bg="#142433", padx=12, pady=10, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1, cursor="hand2")
        left_star.pack(side="left", fill="both", expand=True, padx=(0, 10))
        left_star.bind("<Button-1>", abrir_compra_mas_vendido)

        # Fotografía del producto más vendido a la derecha dentro de left_star
        img_mv_nombre = str(row_mas_vendido.get("imagen", "")) if pd.notna(row_mas_vendido.get("imagen", "")) else ""
        ruta_img_mv = os.path.join(CARPETA_IMAGENES, img_mv_nombre) if img_mv_nombre else ""
        if ruta_img_mv and os.path.exists(ruta_img_mv):
            try:
                pil_mv = Image.open(ruta_img_mv)
                pil_mv = ImageOps.fit(pil_mv, (165, 165), Image.Resampling.LANCZOS)
                tk_mv = ImageTk.PhotoImage(pil_mv)
                self.card_images["MAS_VENDIDO_D001"] = tk_mv
                lbl_img_mv = tk.Label(left_star, image=tk_mv, bg="#142433", highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1, cursor="hand2")
                lbl_img_mv.pack(side="right", padx=(12, 2), pady=2)
                lbl_img_mv.bind("<Button-1>", abrir_compra_mas_vendido)
            except Exception:
                pass

        info_mv = tk.Frame(left_star, bg="#142433", cursor="hand2")
        info_mv.pack(side="left", fill="both", expand=True)
        info_mv.bind("<Button-1>", abrir_compra_mas_vendido)

        l1 = tk.Label(info_mv, text="🏆 #1 MÁS PEDIDO DE LA TEMPORADA", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg="#142433", cursor="hand2")
        l1.pack(anchor="w")
        l1.bind("<Button-1>", abrir_compra_mas_vendido)

        l2 = tk.Label(info_mv, text="Torta de Chocolate Semiamargo", font=FONT_TITULO_SERIF, fg=COLOR_BLANCO_PURO, bg="#142433", cursor="hand2")
        l2.pack(anchor="w", pady=(2, 2))
        l2.bind("<Button-1>", abrir_compra_mas_vendido)

        l3 = tk.Label(
            info_mv,
            text="Bizcocho artesanal esponjoso con capas de cacao 70% de origen, galleta crujiente y crema Chantilly. Disponible en presentaciones de 15cm, 22cm y 28cm con opción de Topper de Cumpleaños.",
            font=FONT_SMALL,
            fg=COLOR_BEIGE_PASTEL,
            bg="#142433",
            justify="left",
            wraplength=340,
            cursor="hand2"
        )
        l3.pack(anchor="w", pady=(2, 6))
        l3.bind("<Button-1>", abrir_compra_mas_vendido)

        l4 = tk.Label(info_mv, text="⭐ 4.9 / 5.0 (Más de 140 pedidos entregados este mes)", font=FONT_SMALL_BOLD, fg="#F4D03F", bg="#142433", cursor="hand2")
        l4.pack(anchor="w", pady=(0, 8))
        l4.bind("<Button-1>", abrir_compra_mas_vendido)

        btn_comprar_mv = tk.Button(
            info_mv,
            text=" 🛒 Personalizar & Comprar #1 Más Vendido ",
            bg=COLOR_MOSTAZA_CLARO,
            fg=COLOR_AZUL_ARANDANO,
            font=FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            padx=10,
            pady=5,
            cursor="hand2",
            activebackground=COLOR_CAFE_MADERA,
            activeforeground=COLOR_MOSTAZA_CLARO,
            command=abrir_compra_mas_vendido
        )
        btn_comprar_mv.pack(anchor="w", pady=(2, 0))

        # LADO DERECHO: Reseñas y Comentarios de Clientes
        right_reviews = tk.Frame(body_box, bg="#142433", padx=12, pady=10, highlightbackground=COLOR_MOSTAZA_CLARO, highlightthickness=1)
        right_reviews.pack(side="right", fill="both", expand=True)

        tk.Label(right_reviews, text="💬 RESEÑAS & OPINIONES DE COMPRADORES", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg="#142433").pack(anchor="w", pady=(0, 5))

        reviews_data = [
            ("⭐ 5.0", "Isabella Moreno", "¡La pedimos de 22cm con topper de cumpleaños para mi mamá y todos quedaron fascinados! El dulce es equilibrado."),
            ("⭐ 5.0", "Madeleine Palacios", "Las galletas NyC Red Velvet en paquete de 8 vienen súper frescas. La mejor pastelería de autor en Cali."),
            ("⭐ 4.9", "Carlos Rodríguez", "Atención excelente por WhatsApp y la entrega llegó rápida en menos de 60 minutos.")
        ]

        for rating_txt, cliente_nom, comentario_txt in reviews_data:
            r_box = tk.Frame(right_reviews, bg=COLOR_AZUL_ARANDANO, padx=8, pady=4)
            r_box.pack(fill="x", pady=3)

            top_line = tk.Frame(r_box, bg=COLOR_AZUL_ARANDANO)
            top_line.pack(fill="x")
            tk.Label(top_line, text=f"{rating_txt} • {cliente_nom}", font=FONT_SMALL_BOLD, fg=COLOR_MOSTAZA_CLARO, bg=COLOR_AZUL_ARANDANO).pack(side="left")

            tk.Label(r_box, text=f'"{comentario_txt}"', font=FONT_SMALL, fg=COLOR_BLANCO_PURO, bg=COLOR_AZUL_ARANDANO, justify="left", wraplength=420).pack(anchor="w")


# ==========================================
# PUNTO DE ENTRADA DE LA APLICACIÓN
# ==========================================
if __name__ == "__main__":
    app = DuleittesApp()
    app.mainloop()
import os
import sys
import time
import queue
import threading
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import config_manager
import bot_engine

# Colores y estilo de la interfaz moderna
COLOR_BG = "#f4f6f9"
COLOR_HEADER = "#1e293b"
COLOR_HEADER_TEXT = "#ffffff"
COLOR_CARD_BG = "#ffffff"
COLOR_PRIMARY = "#2563eb"
COLOR_PRIMARY_HOVER = "#1d4ed8"
COLOR_SUCCESS = "#16a34a"
COLOR_WARNING = "#d97706"
COLOR_DANGER = "#dc2626"
COLOR_TEXT = "#0f172a"
COLOR_TEXT_MUTED = "#64748b"
COLOR_BORDER = "#e2e8f0"
COLOR_CONSOLE_BG = "#0f172a"
COLOR_CONSOLE_TEXT = "#f8fafc"


class IEEPOBotGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Bot de Inventarios IEEPO v2.0 Pro - Descargas Automáticas")
        self.root.geometry("1000x720")
        self.root.minsize(860, 580)
        self.root.configure(bg=COLOR_BG)

        # Cargar configuración persistente
        self.config = config_manager.load_config()

        # Variables de control
        self.var_url = tk.StringVar(value=self.config.get("url", config_manager.DEFAULT_URL))
        self.var_user = tk.StringVar(value=self.config.get("user", ""))
        self.var_pass = tk.StringVar(value=self.config.get("password", ""))
        self.var_download_dir = tk.StringVar(value=self.config.get("download_dir", ""))
        self.var_excel_path = tk.StringVar(value=self.config.get("excel_path", ""))
        self.var_headless = tk.BooleanVar(value=self.config.get("headless", True))
        self.var_skip_existing = tk.BooleanVar(value=self.config.get("skip_existing", True))
        self.var_auto_open = tk.BooleanVar(value=self.config.get("auto_open_folder", True))
        self.var_timeout = tk.IntVar(value=self.config.get("timeout_seconds", 30))
        self.var_retries = tk.IntVar(value=self.config.get("retries", 2))
        self.var_show_pass = tk.BooleanVar(value=False)

        # Estado de ejecución
        self.is_running = False
        self.stop_event = None
        self.worker_thread = None
        self.msg_queue = queue.Queue()

        # Métricas
        self.count_total = tk.IntVar(value=0)
        self.count_exitosos = tk.IntVar(value=0)
        self.count_omitidos = tk.IntVar(value=0)
        self.count_fallidos = tk.IntVar(value=0)
        self.centros_cargados = []

        # Configurar estilos ttk
        self._configurar_estilos()

        # Construir Interfaz
        self._construir_header()
        self._construir_notebook()
        self._construir_statusbar()

        # Iniciar despachador de eventos de la cola
        self.root.after(100, self._procesar_cola_mensajes)

        # Manejar cierre seguro de la aplicación
        self.root.protocol("WM_DELETE_WINDOW", self._al_cerrar_ventana)

        # Cargar lista inicial de centros
        self._recargar_centros_en_tabla(silent=True)
        self.log("INFO", "Aplicación iniciada. Configuración cargada correctamente.")

    # -------------------------------------------------------------------------
    # ESTILOS Y TEMAS
    # -------------------------------------------------------------------------
    def _configurar_estilos(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 9))
        self.style.configure("TNotebook", background=COLOR_BG, borderwidth=0)
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[16, 8], background="#cbd5e1", foreground="#334155")
        self.style.map("TNotebook.Tab",
                       background=[("selected", "#ffffff"), ("active", "#e2e8f0")],
                       foreground=[("selected", COLOR_PRIMARY), ("active", "#1e293b")])

        self.style.configure("Card.TFrame", background=COLOR_CARD_BG, relief="flat")
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background=COLOR_PRIMARY, foreground="#ffffff", padding=[14, 7])
        self.style.map("Primary.TButton", background=[("active", COLOR_PRIMARY_HOVER), ("disabled", "#94a3b8")])

        self.style.configure("Success.TButton", font=("Segoe UI", 10, "bold"), background=COLOR_SUCCESS, foreground="#ffffff", padding=[14, 7])
        self.style.map("Success.TButton", background=[("active", "#15803d"), ("disabled", "#94a3b8")])

        self.style.configure("Danger.TButton", font=("Segoe UI", 10, "bold"), background=COLOR_DANGER, foreground="#ffffff", padding=[14, 7])
        self.style.map("Danger.TButton", background=[("active", "#b91c1c"), ("disabled", "#94a3b8")])

        self.style.configure("Treeview", font=("Segoe UI", 9), rowheight=26, background="#ffffff", fieldbackground="#ffffff")
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#e2e8f0", foreground=COLOR_TEXT)

    # -------------------------------------------------------------------------
    # CONSTRUCCIÓN DE COMPONENTES
    # -------------------------------------------------------------------------
    def _construir_header(self):
        header_frame = tk.Frame(self.root, bg=COLOR_HEADER, height=64)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        # Contenedor interno para título y subtítulo
        title_box = tk.Frame(header_frame, bg=COLOR_HEADER)
        title_box.pack(side="left", padx=18, pady=8)

        lbl_title = tk.Label(
            title_box,
            text="BOT DE INVENTARIOS IEEPO",
            font=("Segoe UI", 14, "bold"),
            bg=COLOR_HEADER,
            fg=COLOR_HEADER_TEXT
        )
        lbl_title.pack(anchor="w")

        lbl_sub = tk.Label(
            title_box,
            text="Descarga masiva automatizada de reportes de área • Versión 2.0 Pro",
            font=("Segoe UI", 9),
            bg=COLOR_HEADER,
            fg="#94a3b8"
        )
        lbl_sub.pack(anchor="w")

        # Botones de cabecera rápida
        btn_box = tk.Frame(header_frame, bg=COLOR_HEADER)
        btn_box.pack(side="right", padx=18, pady=12)

        btn_folder = tk.Button(
            btn_box,
            text="📁 Abrir Carpeta",
            font=("Segoe UI", 9, "bold"),
            bg="#334155",
            fg="#ffffff",
            activebackground="#475569",
            activeforeground="#ffffff",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._abrir_carpeta_descargas
        )
        btn_folder.pack(side="left", padx=4)

        btn_excel = tk.Button(
            btn_box,
            text="📊 Abrir Excel",
            font=("Segoe UI", 9, "bold"),
            bg="#334155",
            fg="#ffffff",
            activebackground="#475569",
            activeforeground="#ffffff",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._abrir_archivo_excel
        )
        btn_excel.pack(side="left", padx=4)

    def _construir_notebook(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(8, 4))

        # Pestaña 1: Ejecución y Descargas
        self.tab_ejecucion = tk.Frame(self.notebook, bg=COLOR_BG)
        self.notebook.add(self.tab_ejecucion, text="  🚀 Descargas y Ejecución  ")
        self._construir_tab_ejecucion()

        # Pestaña 2: Configuración
        self.tab_config = tk.Frame(self.notebook, bg=COLOR_BG)
        self.notebook.add(self.tab_config, text="  ⚙️ Configuración y Conexión  ")
        self._construir_tab_config()

        # Pestaña 3: Centros y Reportes
        self.tab_centros = tk.Frame(self.notebook, bg=COLOR_BG)
        self.notebook.add(self.tab_centros, text="  📋 Centros y Resultados  ")
        self._construir_tab_centros()

    def _construir_tab_ejecucion(self):
        # Panel superior: Opciones y controles de inicio
        top_card = tk.Frame(self.tab_ejecucion, bg=COLOR_CARD_BG, bd=1, relief="solid", highlightbackground=COLOR_BORDER)
        top_card.pack(fill="x", padx=10, pady=8)

        # Fila 1: Opciones de ejecución rápida
        opts_frame = tk.Frame(top_card, bg=COLOR_CARD_BG)
        opts_frame.pack(fill="x", padx=12, pady=(10, 6))

        chk_headless = tk.Checkbutton(
            opts_frame,
            text="🌐 Modo Silencioso (Segundo plano)",
            variable=self.var_headless,
            bg=COLOR_CARD_BG,
            activebackground=COLOR_CARD_BG,
            font=("Segoe UI", 9, "bold"),
            command=self._guardar_config_silenciosa
        )
        chk_headless.pack(side="left", padx=(0, 16))

        chk_skip = tk.Checkbutton(
            opts_frame,
            text="⏩ Omitir PDFs ya descargados",
            variable=self.var_skip_existing,
            bg=COLOR_CARD_BG,
            activebackground=COLOR_CARD_BG,
            font=("Segoe UI", 9, "bold"),
            command=self._guardar_config_silenciosa
        )
        chk_skip.pack(side="left", padx=(0, 16))

        chk_open = tk.Checkbutton(
            opts_frame,
            text="📂 Abrir carpeta al finalizar",
            variable=self.var_auto_open,
            bg=COLOR_CARD_BG,
            activebackground=COLOR_CARD_BG,
            font=("Segoe UI", 9),
            command=self._guardar_config_silenciosa
        )
        chk_open.pack(side="left")

        # Fila 2: Botones de Acción
        btn_frame = tk.Frame(top_card, bg=COLOR_CARD_BG)
        btn_frame.pack(fill="x", padx=12, pady=(6, 12))

        self.btn_iniciar = tk.Button(
            btn_frame,
            text="▶ INICIAR DESCARGA",
            font=("Segoe UI", 11, "bold"),
            bg=COLOR_SUCCESS,
            fg="#ffffff",
            activebackground="#15803d",
            activeforeground="#ffffff",
            relief="flat",
            padx=18,
            pady=8,
            cursor="hand2",
            command=self._iniciar_descarga
        )
        self.btn_iniciar.pack(side="left", padx=(0, 8))

        self.btn_detener = tk.Button(
            btn_frame,
            text="⏹ DETENER",
            font=("Segoe UI", 11, "bold"),
            bg=COLOR_DANGER,
            fg="#ffffff",
            activebackground="#b91c1c",
            activeforeground="#ffffff",
            relief="flat",
            padx=18,
            pady=8,
            cursor="hand2",
            state="disabled",
            command=self._detener_descarga
        )
        self.btn_detener.pack(side="left", padx=(0, 12))

        btn_reload = tk.Button(
            btn_frame,
            text="🔄 Recargar Excel",
            font=("Segoe UI", 9),
            bg="#f1f5f9",
            fg=COLOR_TEXT,
            activebackground="#e2e8f0",
            relief="groove",
            padx=10,
            pady=6,
            cursor="hand2",
            command=self._recargar_centros_en_tabla
        )
        btn_reload.pack(side="left")

        # Tarjetas de Estadísticas en Vivo
        stats_frame = tk.Frame(self.tab_ejecucion, bg=COLOR_BG)
        stats_frame.pack(fill="x", padx=10, pady=(0, 6))

        self._crear_tarjeta_metrica(stats_frame, "Total Centros", self.count_total, "#3b82f6")
        self._crear_tarjeta_metrica(stats_frame, "Exitosos ✓", self.count_exitosos, "#22c55e")
        self._crear_tarjeta_metrica(stats_frame, "Omitidos ⏩", self.count_omitidos, "#a855f7")
        self._crear_tarjeta_metrica(stats_frame, "Fallidos ⚠️", self.count_fallidos, "#ef4444")

        # Barra de progreso
        prog_card = tk.Frame(self.tab_ejecucion, bg=COLOR_CARD_BG, bd=1, relief="solid", highlightbackground=COLOR_BORDER)
        prog_card.pack(fill="x", padx=10, pady=4)

        prog_inner = tk.Frame(prog_card, bg=COLOR_CARD_BG)
        prog_inner.pack(fill="x", padx=12, pady=8)

        self.lbl_progreso_estado = tk.Label(
            prog_inner,
            text="Listo para iniciar.",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD_BG,
            fg=COLOR_TEXT
        )
        self.lbl_progreso_estado.pack(side="left")

        self.lbl_progreso_pct = tk.Label(
            prog_inner,
            text="0%",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD_BG,
            fg=COLOR_PRIMARY
        )
        self.lbl_progreso_pct.pack(side="right")

        self.progress_bar = ttk.Progressbar(prog_card, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", padx=12, pady=(0, 8))

        # Consola de Logs Integrada
        log_card = tk.Frame(self.tab_ejecucion, bg=COLOR_CARD_BG, bd=1, relief="solid", highlightbackground=COLOR_BORDER)
        log_card.pack(fill="both", expand=True, padx=10, pady=(4, 8))

        log_header = tk.Frame(log_card, bg="#1e293b")
        log_header.pack(fill="x")

        lbl_log_title = tk.Label(
            log_header,
            text="📟 Registro de Eventos en Tiempo Real (Log)",
            font=("Segoe UI", 9, "bold"),
            bg="#1e293b",
            fg="#f8fafc"
        )
        lbl_log_title.pack(side="left", padx=10, pady=6)

        btn_copy_log = tk.Button(
            log_header,
            text="📋 Copiar",
            font=("Segoe UI", 8),
            bg="#334155",
            fg="#ffffff",
            relief="flat",
            padx=6,
            pady=2,
            command=self._copiar_log
        )
        btn_copy_log.pack(side="right", padx=6, pady=4)

        btn_clear_log = tk.Button(
            log_header,
            text="🧹 Limpiar",
            font=("Segoe UI", 8),
            bg="#334155",
            fg="#ffffff",
            relief="flat",
            padx=6,
            pady=2,
            command=self._limpiar_log
        )
        btn_clear_log.pack(side="right", padx=2, pady=4)

        # ScrolledText para la consola
        self.log_text = scrolledtext.ScrolledText(
            log_card,
            bg=COLOR_CONSOLE_BG,
            fg=COLOR_CONSOLE_TEXT,
            font=("Consolas", 9),
            wrap="word",
            bd=0,
            insertbackground="#ffffff"
        )
        self.log_text.pack(fill="both", expand=True, padx=4, pady=4)

        # Configurar tags de colores en la consola
        self.log_text.tag_config("INFO", foreground="#93c5fd")
        self.log_text.tag_config("SUCCESS", foreground="#4ade80")
        self.log_text.tag_config("WARNING", foreground="#fde047")
        self.log_text.tag_config("ERROR", foreground="#f87171")
        self.log_text.tag_config("TIMESTAMP", foreground="#94a3b8")

    def _crear_tarjeta_metrica(self, parent, titulo, variable, color_acento):
        card = tk.Frame(parent, bg=COLOR_CARD_BG, bd=1, relief="solid", highlightbackground=COLOR_BORDER)
        card.pack(side="left", fill="both", expand=True, padx=4)

        # Barra decorativa de color
        accent = tk.Frame(card, bg=color_acento, height=3)
        accent.pack(fill="x", side="top")

        lbl_val = tk.Label(card, textvariable=variable, font=("Segoe UI", 16, "bold"), bg=COLOR_CARD_BG, fg=color_acento)
        lbl_val.pack(pady=(4, 0))

        lbl_tit = tk.Label(card, text=titulo, font=("Segoe UI", 8, "bold"), bg=COLOR_CARD_BG, fg=COLOR_TEXT_MUTED)
        lbl_tit.pack(pady=(0, 6))

    def _construir_tab_config(self):
        # Scrollable container si la pantalla es pequeña
        main_frame = tk.Frame(self.tab_config, bg=COLOR_BG)
        main_frame.pack(fill="both", expand=True, padx=16, pady=12)

        # Sección 1: Credenciales y URL
        grp_cred = tk.LabelFrame(
            main_frame,
            text="  🌐 Acceso al Portal IEEPO  ",
            font=("Segoe UI", 10, "bold"),
            bg=COLOR_CARD_BG,
            fg=COLOR_PRIMARY,
            bd=1,
            relief="solid",
            highlightbackground=COLOR_BORDER,
            padx=14,
            pady=12
        )
        grp_cred.pack(fill="x", pady=(0, 12))

        # URL
        tk.Label(grp_cred, text="URL del Portal:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD_BG).grid(row=0, column=0, sticky="w", pady=4)
        entry_url = ttk.Entry(grp_cred, textvariable=self.var_url, width=60)
        entry_url.grid(row=0, column=1, columnspan=2, sticky="we", padx=(8, 0), pady=4)

        # Usuario
        tk.Label(grp_cred, text="Usuario:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD_BG).grid(row=1, column=0, sticky="w", pady=4)
        entry_user = ttk.Entry(grp_cred, textvariable=self.var_user, width=30)
        entry_user.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=4)

        # Contraseña
        tk.Label(grp_cred, text="Contraseña:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD_BG).grid(row=2, column=0, sticky="w", pady=4)
        pass_frame = tk.Frame(grp_cred, bg=COLOR_CARD_BG)
        pass_frame.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=4)

        self.entry_pass = ttk.Entry(pass_frame, textvariable=self.var_pass, width=30, show="*")
        self.entry_pass.pack(side="left")

        self.btn_toggle_pass = tk.Button(
            pass_frame,
            text="👁️",
            font=("Segoe UI", 8),
            bg="#f1f5f9",
            relief="groove",
            command=self._toggle_ver_password
        )
        self.btn_toggle_pass.pack(side="left", padx=4)

        # Botón Probar Conexión
        self.btn_test_conn = tk.Button(
            grp_cred,
            text="🔍 Probar Conexión y Credenciales",
            font=("Segoe UI", 9, "bold"),
            bg="#0284c7",
            fg="#ffffff",
            activebackground="#0369a1",
            activeforeground="#ffffff",
            relief="flat",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._probar_conexion
        )
        self.btn_test_conn.grid(row=3, column=1, sticky="w", padx=(8, 0), pady=(8, 4))

        self.lbl_test_result = tk.Label(grp_cred, text="", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD_BG)
        self.lbl_test_result.grid(row=4, column=0, columnspan=3, sticky="w", pady=4)

        grp_cred.columnconfigure(1, weight=1)

        # Sección 2: Rutas y Archivos
        grp_paths = tk.LabelFrame(
            main_frame,
            text="  📁 Carpetas y Archivos de Trabajo  ",
            font=("Segoe UI", 10, "bold"),
            bg=COLOR_CARD_BG,
            fg=COLOR_PRIMARY,
            bd=1,
            relief="solid",
            highlightbackground=COLOR_BORDER,
            padx=14,
            pady=12
        )
        grp_paths.pack(fill="x", pady=(0, 12))

        # Carpeta Descargas
        tk.Label(grp_paths, text="Carpeta de Descargas:", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD_BG).grid(row=0, column=0, sticky="w", pady=4)
        entry_dir = ttk.Entry(grp_paths, textvariable=self.var_download_dir, width=50)
        entry_dir.grid(row=0, column=1, sticky="we", padx=(8, 6), pady=4)

        btn_browse_dir = tk.Button(
            grp_paths,
            text="Examinar...",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            relief="groove",
            command=self._seleccionar_carpeta_descargas
        )
        btn_browse_dir.grid(row=0, column=2, sticky="w", pady=4)

        # Archivo Excel
        tk.Label(grp_paths, text="Archivo Excel (CTs):", font=("Segoe UI", 9, "bold"), bg=COLOR_CARD_BG).grid(row=1, column=0, sticky="w", pady=4)
        entry_excel = ttk.Entry(grp_paths, textvariable=self.var_excel_path, width=50)
        entry_excel.grid(row=1, column=1, sticky="we", padx=(8, 6), pady=4)

        btn_browse_excel = tk.Button(
            grp_paths,
            text="Examinar...",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            relief="groove",
            command=self._seleccionar_archivo_excel
        )
        btn_browse_excel.grid(row=1, column=2, sticky="w", pady=4)

        # Botones de Acción para Excel
        excel_actions_frame = tk.Frame(grp_paths, bg=COLOR_CARD_BG)
        excel_actions_frame.grid(row=2, column=1, columnspan=2, sticky="w", padx=(8, 0), pady=(10, 4))

        btn_gen_excel = tk.Button(
            excel_actions_frame,
            text="📄 Generar Plantilla Excel ('centros.xlsx')",
            font=("Segoe UI", 9, "bold"),
            bg="#059669",
            fg="#ffffff",
            activebackground="#047857",
            activeforeground="#ffffff",
            relief="flat",
            padx=10,
            pady=5,
            cursor="hand2",
            command=self._generar_plantilla_excel
        )
        btn_gen_excel.pack(side="left", padx=(0, 8))

        btn_open_excel = tk.Button(
            excel_actions_frame,
            text="📊 Abrir Excel",
            font=("Segoe UI", 9),
            bg="#f1f5f9",
            relief="groove",
            padx=10,
            pady=5,
            cursor="hand2",
            command=self._abrir_archivo_excel
        )
        btn_open_excel.pack(side="left")

        grp_paths.columnconfigure(1, weight=1)

        # Sección 3: Guardar Configuración
        bottom_box = tk.Frame(main_frame, bg=COLOR_BG)
        bottom_box.pack(fill="x", pady=6)

        btn_save = tk.Button(
            bottom_box,
            text="💾 GUARDAR TODA LA CONFIGURACIÓN",
            font=("Segoe UI", 10, "bold"),
            bg=COLOR_PRIMARY,
            fg="#ffffff",
            activebackground=COLOR_PRIMARY_HOVER,
            activeforeground="#ffffff",
            relief="flat",
            padx=16,
            pady=8,
            cursor="hand2",
            command=self._guardar_config_manual
        )
        btn_save.pack(side="left")

    def _construir_tab_centros(self):
        container = tk.Frame(self.tab_centros, bg=COLOR_BG)
        container.pack(fill="both", expand=True, padx=12, pady=10)

        # Barra de herramientas superior
        tb = tk.Frame(container, bg=COLOR_CARD_BG, bd=1, relief="solid", highlightbackground=COLOR_BORDER)
        tb.pack(fill="x", pady=(0, 8))

        lbl_info = tk.Label(
            tb,
            text="📋 Lista de Claves de Centro de Trabajo leídas del archivo Excel:",
            font=("Segoe UI", 9, "bold"),
            bg=COLOR_CARD_BG
        )
        lbl_info.pack(side="left", padx=12, pady=8)

        btn_open_pdf = tk.Button(
            tb,
            text="📄 Abrir PDF Seleccionado",
            font=("Segoe UI", 9),
            bg="#f1f5f9",
            relief="groove",
            padx=8,
            pady=4,
            command=self._abrir_pdf_seleccionado
        )
        btn_open_pdf.pack(side="right", padx=10, pady=6)

        btn_open_resumen = tk.Button(
            tb,
            text="📊 Abrir Reporte Resumen",
            font=("Segoe UI", 9),
            bg="#f1f5f9",
            relief="groove",
            padx=8,
            pady=4,
            command=self._abrir_resumen_excel
        )
        btn_open_resumen.pack(side="right", padx=4, pady=6)

        # Tabla Treeview
        tree_card = tk.Frame(container, bg=COLOR_CARD_BG, bd=1, relief="solid", highlightbackground=COLOR_BORDER)
        tree_card.pack(fill="both", expand=True)

        columns = ("num", "centro", "estado", "archivo", "detalle")
        self.tree = ttk.Treeview(tree_card, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("num", text="#")
        self.tree.heading("centro", text="Centro de Trabajo (CCT)")
        self.tree.heading("estado", text="Estado")
        self.tree.heading("archivo", text="Archivo Descargado")
        self.tree.heading("detalle", text="Detalle / Observación")

        self.tree.column("num", width=40, anchor="center")
        self.tree.column("centro", width=180, anchor="w")
        self.tree.column("estado", width=120, anchor="center")
        self.tree.column("archivo", width=220, anchor="w")
        self.tree.column("detalle", width=280, anchor="w")

        scroll_y = ttk.Scrollbar(tree_card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)

        self.tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        scroll_y.pack(side="right", fill="y", pady=4)

        # Tags de colores en la tabla
        self.tree.tag_configure("EXITO", background="#dcfce7", foreground="#166534")
        self.tree.tag_configure("OMITIDO", background="#f3e8ff", foreground="#6b21a8")
        self.tree.tag_configure("ERROR", background="#fee2e2", foreground="#991b1b")
        self.tree.tag_configure("PENDIENTE", background="#ffffff", foreground="#475569")

        self.tree.bind("<Double-1>", lambda e: self._abrir_pdf_seleccionado())

    def _construir_statusbar(self):
        status_bar = tk.Frame(self.root, bg="#e2e8f0", height=24)
        status_bar.pack(fill="x", side="bottom")

        self.lbl_status_left = tk.Label(
            status_bar,
            text="Estado: En espera",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            fg=COLOR_TEXT_MUTED
        )
        self.lbl_status_left.pack(side="left", padx=10)

        self.lbl_status_right = tk.Label(
            status_bar,
            text=f"Directorio: {config_manager.get_base_dir().name}",
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            fg=COLOR_TEXT_MUTED
        )
        self.lbl_status_right.pack(side="right", padx=10)

    # -------------------------------------------------------------------------
    # FUNCIONALIDADES Y EVENTOS
    # -------------------------------------------------------------------------
    def log(self, level: str, message: str):
        """Envía un mensaje de log de forma segura a la cola de eventos."""
        self.msg_queue.put(("LOG", (level, message)))

    def _procesar_cola_mensajes(self):
        """Procesa mensajes acumulados en la cola y actualiza la GUI sin bloquearla."""
        try:
            while not self.msg_queue.empty():
                msg_type, data = self.msg_queue.get_nowait()

                if msg_type == "LOG":
                    level, text = data
                    now = datetime.now().strftime("%H:%M:%S")
                    self.log_text.insert(tk.END, f"[{now}] ", "TIMESTAMP")
                    self.log_text.insert(tk.END, f"[{level}] {text}\n", level)
                    self.log_text.see(tk.END)

                elif msg_type == "PROGRESS":
                    curr, total, pct, text = data
                    self.progress_bar["value"] = pct
                    self.lbl_progreso_pct.config(text=f"{pct}%")
                    self.lbl_progreso_estado.config(text=text)
                    self.lbl_status_left.config(text=f"Estado: {text}")

                elif msg_type == "ITEM_STATUS":
                    idx, ct_name, status, file_path, detail = data
                    # Actualizar árbol
                    item_id = str(idx)
                    if self.tree.exists(item_id):
                        tag = "PENDIENTE"
                        if "ÉXITO" in status.upper():
                            tag = "EXITO"
                        elif "OMITIDO" in status.upper():
                            tag = "OMITIDO"
                        elif "ERROR" in status.upper():
                            tag = "ERROR"

                        fname = Path(file_path).name if file_path else ""
                        self.tree.item(item_id, values=(idx + 1, ct_name, status, fname, detail), tags=(tag,))
                        self.tree.see(item_id)

                    # Actualizar contadores
                    self._actualizar_contadores_desde_tabla()

                elif msg_type == "TEST_RESULT":
                    success, msg = data
                    self.btn_test_conn.config(state="normal", text="🔍 Probar Conexión y Credenciales")
                    if success:
                        self.lbl_test_result.config(text=f"✓ {msg}", fg=COLOR_SUCCESS)
                        messagebox.showinfo("Conexión Exitosa", f"¡Conexión y credenciales verificadas con éxito!\n\n{msg}")
                    else:
                        self.lbl_test_result.config(text=f"⚠️ {msg}", fg=COLOR_DANGER)
                        messagebox.showerror("Fallo de Conexión", f"No se pudo completar la conexión:\n\n{msg}")

                elif msg_type == "FINISHED":
                    res = data
                    self.is_running = False
                    self.btn_iniciar.config(state="normal")
                    self.btn_detener.config(state="disabled")
                    self.lbl_status_left.config(text="Estado: Finalizado")

                    if not res.get("cancelado", False):
                        if self.var_auto_open.get():
                            self._abrir_carpeta_descargas()

                        messagebox.showinfo(
                            "Descargas Finalizadas",
                            f"Proceso de descarga completado.\n\n"
                            f"• Total: {res.get('total', 0)}\n"
                            f"• Exitosos: {res.get('exitosos', 0)}\n"
                            f"• Omitidos: {res.get('omitidos', 0)}\n"
                            f"• Fallidos: {res.get('fallidos', 0)}\n\n"
                            f"Reporte guardado en:\n{res.get('resumen_excel', '')}"
                        )
                    else:
                        messagebox.showwarning("Proceso Detenido", "El proceso fue cancelado por el usuario.")

        except Exception as e:
            print(f"Error en procesar cola: {e}")

        # Re-programar revisión de la cola cada 100ms
        self.root.after(100, self._procesar_cola_mensajes)

    def _actualizar_contadores_desde_tabla(self):
        total = len(self.tree.get_children())
        exitos = 0
        omitidos = 0
        fallidos = 0

        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id, "values")
            if len(vals) >= 3:
                st = str(vals[2]).upper()
                if "ÉXITO" in st:
                    exitos += 1
                elif "OMITIDO" in st:
                    omitidos += 1
                elif "ERROR" in st or "TIMEOUT" in st:
                    fallidos += 1

        self.count_total.set(total)
        self.count_exitosos.set(exitos)
        self.count_omitidos.set(omitidos)
        self.count_fallidos.set(fallidos)

    def _recargar_centros_en_tabla(self, silent: bool = False):
        """Lee el Excel y llena la tabla de centros."""
        excel_path = Path(self.var_excel_path.get())
        for item in self.tree.get_children():
            self.tree.delete(item)

        dest_dir = Path(self.var_download_dir.get())
        centros, msg = bot_engine.leer_centros_excel(excel_path)
        self.centros_cargados = centros

        for idx, ct in enumerate(centros):
            ya_esta = bot_engine.ya_fue_descargado(ct, dest_dir)
            st = "OMITIDO" if ya_esta else "PENDIENTE"
            fname = f"{ct.upper()}.pdf" if ya_esta else ""
            det = "Archivo ya existe en carpeta" if ya_esta else "Listo para descargar"
            tag = "OMITIDO" if ya_esta else "PENDIENTE"
            self.tree.insert("", "end", iid=str(idx), values=(idx + 1, ct, st, fname, det), tags=(tag,))

        self._actualizar_contadores_desde_tabla()
        if not silent:
            self.log("INFO", f"Excel recargado: {len(centros)} centros de trabajo encontrados.")
            messagebox.showinfo("Centros Cargados", f"Se han cargado {len(centros)} Centros de Trabajo desde:\n{excel_path.resolve()}")

    def _toggle_ver_password(self):
        if self.var_show_pass.get():
            self.entry_pass.config(show="*")
            self.var_show_pass.set(False)
            self.btn_toggle_pass.config(text="👁️")
        else:
            self.entry_pass.config(show="")
            self.var_show_pass.set(True)
            self.btn_toggle_pass.config(text="🔒")

    def _seleccionar_carpeta_descargas(self):
        carpeta = filedialog.askdirectory(
            title="Seleccionar Carpeta de Descargas",
            initialdir=self.var_download_dir.get() or str(config_manager.get_base_dir())
        )
        if carpeta:
            self.var_download_dir.set(carpeta)
            self._guardar_config_silenciosa()
            self._recargar_centros_en_tabla(silent=True)
            self.log("INFO", f"Carpeta de descargas actualizada: {carpeta}")

    def _seleccionar_archivo_excel(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar Archivo Excel con Centros de Trabajo",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
            initialdir=str(Path(self.var_excel_path.get()).parent) if self.var_excel_path.get() else str(config_manager.get_base_dir())
        )
        if archivo:
            self.var_excel_path.set(archivo)
            self._guardar_config_silenciosa()
            self._recargar_centros_en_tabla(silent=False)

    def _generar_plantilla_excel(self):
        from tkinter import filedialog
        
        # Sugerir la ruta que ya está en el campo, o una por defecto
        ruta_actual = self.var_excel_path.get()
        dir_inicial = os.path.dirname(ruta_actual) if ruta_actual else os.path.expanduser("~")
        
        nuevo_archivo = filedialog.asksaveasfilename(
            title="Guardar plantilla Excel como...",
            initialdir=dir_inicial,
            initialfile="centros_ejemplo.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        
        if not nuevo_archivo:
            return
            
        excel_path = Path(nuevo_archivo)
        
        ok, msg = bot_engine.generar_plantilla_excel(excel_path, sobreescribir=True)
        if ok:
            # Actualizar la ruta en la interfaz
            self.var_excel_path.set(str(excel_path.resolve()))
            self._guardar_config_silenciosa()
            
            self.log("SUCCESS", msg)
            self._recargar_centros_en_tabla(silent=True)
            messagebox.showinfo("Plantilla Creada", f"Plantilla guardada en:\n{excel_path}")
        else:
            self.log("ERROR", msg)
            messagebox.showerror("Error", msg)

    def _abrir_archivo_excel(self):
        ruta = Path(self.var_excel_path.get()).resolve()
        if not ruta.exists():
            messagebox.showwarning("Archivo no encontrado", f"El archivo Excel no existe:\n{ruta}")
            return
        try:
            os.startfile(ruta)
        except Exception as e:
            messagebox.showerror("Error al abrir archivo", f"No se pudo abrir el archivo:\n{e}")

    def _abrir_carpeta_descargas(self):
        carpeta = Path(self.var_download_dir.get()).resolve()
        carpeta.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(carpeta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{e}")

    def _abrir_resumen_excel(self):
        carpeta = Path(self.var_download_dir.get()).resolve()
        resumen = carpeta / "resumen_ejecucion.xlsx"
        if resumen.exists():
            try:
                os.startfile(resumen)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el reporte de resumen:\n{e}")
        else:
            messagebox.showinfo("Sin reporte", "Aún no se ha generado un reporte 'resumen_ejecucion.xlsx' en esta carpeta.")

    def _abrir_pdf_seleccionado(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Selección requerida", "Por favor selecciona un Centro de Trabajo en la tabla.")
            return

        item_vals = self.tree.item(sel[0], "values")
        if len(item_vals) >= 2:
            ct = item_vals[1]
            carpeta = Path(self.var_download_dir.get()).resolve()
            pdf_path = carpeta / f"{bot_engine.limpiar_nombre_archivo(ct.upper())}.pdf"
            if pdf_path.exists():
                try:
                    os.startfile(pdf_path)
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo abrir el PDF:\n{e}")
            else:
                messagebox.showwarning("PDF no encontrado", f"El archivo PDF para '{ct}' aún no ha sido descargado en:\n{pdf_path}")

    def _obtener_dict_config(self) -> dict:
        return {
            "url": self.var_url.get().strip(),
            "user": self.var_user.get().strip(),
            "password": self.var_pass.get().strip(),
            "download_dir": self.var_download_dir.get().strip(),
            "excel_path": self.var_excel_path.get().strip(),
            "headless": self.var_headless.get(),
            "skip_existing": self.var_skip_existing.get(),
            "auto_open_folder": self.var_auto_open.get(),
            "timeout_seconds": self.var_timeout.get(),
            "retries": self.var_retries.get()
        }

    def _guardar_config_silenciosa(self):
        cfg = self._obtener_dict_config()
        config_manager.save_config(cfg)

    def _guardar_config_manual(self):
        cfg = self._obtener_dict_config()
        ok = config_manager.save_config(cfg)
        if ok:
            self.log("SUCCESS", "Configuración guardada exitosamente en config.json y .env")
            messagebox.showinfo("Configuración Guardada", "Todos los datos y preferencias fueron guardados correctamente.")
        else:
            messagebox.showerror("Error", "No se pudo guardar la configuración.")

    def _copiar_log(self):
        texto = self.log_text.get("1.0", tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(texto)
        messagebox.showinfo("Copiado", "Registro de logs copiado al portapapeles.")

    def _limpiar_log(self):
        self.log_text.delete("1.0", tk.END)

    # -------------------------------------------------------------------------
    # EJECUCIÓN MULTIHILO (BACKGROUND WORKERS)
    # -------------------------------------------------------------------------
    def _probar_conexion(self):
        if self.is_running:
            messagebox.showwarning("Proceso activo", "No puedes probar la conexión mientras hay una descarga en curso.")
            return

        cfg = self._obtener_dict_config()
        url = cfg.get("url")
        user = cfg.get("user")
        pwd = cfg.get("password")

        if not user or not pwd:
            messagebox.showwarning("Credenciales requeridas", "Por favor ingresa tu Usuario y Contraseña antes de probar la conexión.")
            return

        self.btn_test_conn.config(state="disabled", text="⏳ Conectando y validando...")
        self.lbl_test_result.config(text="Verificando credenciales en el servidor IEEPO...", fg=COLOR_PRIMARY)
        self.log("INFO", "Iniciando prueba de conexión y credenciales...")

        def _worker():
            def _log_cb(lvl, msg):
                self.msg_queue.put(("LOG", (lvl, msg)))

            success, msg = bot_engine.test_connection_and_credentials(url, user, pwd, log_cb=_log_cb)
            self.msg_queue.put(("TEST_RESULT", (success, msg)))

        t = threading.Thread(target=_worker, daemon=True)
        t.start()

    def _iniciar_descarga(self):
        if self.is_running:
            return

        cfg = self._obtener_dict_config()
        if not cfg.get("user") or not cfg.get("password"):
            messagebox.showwarning(
                "Credenciales Requeridas",
                "Por favor ve a la pestaña '⚙️ Configuración' e ingresa tu Usuario y Contraseña del IEEPO."
            )
            self.notebook.select(self.tab_config)
            return

        excel_path = Path(cfg.get("excel_path", ""))
        if not excel_path.exists():
            messagebox.showerror(
                "Excel no encontrado",
                f"No se encontró el archivo Excel de centros en:\n{excel_path.resolve()}\n\n"
                "Puedes crearlo usando el botón 'Generar Plantilla Excel' en la pestaña Configuración."
            )
            return

        # Guardar configuración actual
        self._guardar_config_silenciosa()

        # Cambiar estado a ejecutando
        self.is_running = True
        self.stop_event = threading.Event()
        self.btn_iniciar.config(state="disabled")
        self.btn_detener.config(state="normal")
        self.lbl_status_left.config(text="Estado: Descargando reportes...")

        # Resetear progreso
        self.progress_bar["value"] = 0
        self.lbl_progreso_pct.config(text="0%")

        # Lanzar hilo de trabajo
        def _worker():
            def _log_cb(lvl, msg):
                self.msg_queue.put(("LOG", (lvl, msg)))

            def _prog_cb(curr, total, pct, txt):
                self.msg_queue.put(("PROGRESS", (curr, total, pct, txt)))

            def _item_status_cb(idx, ct_name, status, file_path, detail):
                self.msg_queue.put(("ITEM_STATUS", (idx, ct_name, status, file_path, detail)))

            res = bot_engine.ejecutar_proceso_descargas(
                config=cfg,
                log_cb=_log_cb,
                progress_cb=_prog_cb,
                item_status_cb=_item_status_cb,
                stop_event=self.stop_event
            )
            self.msg_queue.put(("FINISHED", res))

        self.worker_thread = threading.Thread(target=_worker, daemon=True)
        self.worker_thread.start()

    def _detener_descarga(self):
        if self.is_running and self.stop_event:
            self.btn_detener.config(state="disabled")
            self.lbl_status_left.config(text="Estado: Solicitando detención...")
            self.log("WARNING", "Solicitando cancelación del proceso... (esperando cierre seguro)")
            self.stop_event.set()

    def _al_cerrar_ventana(self):
        if self.is_running:
            if messagebox.askyesno("Proceso en curso", "¿Deseas detener la descarga y salir de la aplicación?"):
                if self.stop_event:
                    self.stop_event.set()
                self.root.destroy()
        else:
            self._guardar_config_silenciosa()
            self.root.destroy()


def main():
    root = tk.Tk()
    app = IEEPOBotGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

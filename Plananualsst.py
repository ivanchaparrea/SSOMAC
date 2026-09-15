import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk
import pymupdf  # PyMuPDF para renderizar páginas PDF


class PlanAnualSSTApp:

    def __init__(self, root):
        self.root = root
        self.root.title(
            "Sistema de Gestión - Plan Anual de Seguridad y Salud en el Trabajo"
        )
        self.root.geometry("1050x750")
        self.root.config(bg="#1E293B")

        # Nombres de archivos predeterminados en la misma carpeta
        self.pdf_plan_filename = "plan_anual_sst.pdf"
        self.pdf_plano_filename = "plano_evacuacion.pdf"
        self.dwg_plano_filename = "plano_evacuacion.dwg"
        
        # Archivos para Matriz IPERC
        self.excel_iperc_filename = "matriz_iperc.xlsx"
        self.pdf_iperc_filename = "matriz_iperc.pdf"

        # Variables de control
        self.logo_img = None
        self.fotos_integrantes = []
        
        # Documentos PDF renderizables
        self.pdf_plan_doc = None
        self.current_page_plan = 0

        self.pdf_plano_doc = None
        self.current_page_plano = 0

        self.pdf_iperc_doc = None
        self.current_page_iperc = 0

        # Estilos de pestañas
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background="#1E293B", borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding=[15, 8],
            background="#334155",
            foreground="#E2E8F0",
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#2563EB")],
            foreground=[("selected", "#FFFFFF")],
        )

        # Franja animada de colores en la parte superior
        self.build_animacion_superior()

        # Navegador de pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Creación de pestañas
        self.tab_portada = ttk.Frame(self.notebook)
        self.tab_plan = ttk.Frame(self.notebook)
        self.tab_matriz_iperc = ttk.Frame(self.notebook)
        self.tab_plano_evacuacion = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_portada, text=" 🏠 Portada ")
        self.notebook.add(self.tab_plan, text=" 📄 Plan Anual de SST ")
        self.notebook.add(self.tab_matriz_iperc, text=" ⚠️ Matriz IPERC ")
        self.notebook.add(
            self.tab_plano_evacuacion, text=" 🗺️ Plano de Evacuación "
        )

        # Construcción de la interfaz
        self.build_portada()
        self.build_plan_anual_pdf_viewer()
        self.build_iperc_viewer()
        self.build_plano_evacuacion_viewer()

    # --- FRANJA DE COLORES ANIMADA EN MOVIMIENTO ---
    def build_animacion_superior(self):
        self.canvas_anim = tk.Canvas(
            self.root, height=10, bg="#1E293B", highlightthickness=0
        )
        self.canvas_anim.pack(fill="x", side="top")

        self.colores_franja = [
            "#EF4444",
            "#F97316",
            "#F59E0B",
            "#10B981",
            "#3B82F6",
            "#6366F1",
            "#8B5CF6",
            "#EC4899",
        ]
        self.offset_franja = 0
        self.animar_franja()

    def animar_franja(self):
        self.canvas_anim.delete("all")
        ancho = self.root.winfo_width()
        if ancho <= 1:
            ancho = 1050

        ancho_bloque = 60
        num_bloques = (ancho // ancho_bloque) + 4

        for i in range(-2, num_bloques):
            x1 = (i * ancho_bloque) + self.offset_franja
            x2 = x1 + ancho_bloque
            color = self.colores_franja[i % len(self.colores_franja)]
            self.canvas_anim.create_rectangle(
                x1, 0, x2, 10, fill=color, outline=""
            )

        self.offset_franja += 2
        if self.offset_franja >= ancho_bloque:
            self.offset_franja = 0

        self.root.after(40, self.animar_franja)

    # --- CÓDIGO DE LA PORTADA CON FONDO DE COLOR ---
    def cargar_e_imprimir_imagen(self, img_path, max_ancho, max_alto):
        if not os.path.exists(img_path):
            return None
        try:
            img = Image.open(img_path)
            img.thumbnail((max_ancho, max_alto), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None

    def build_portada(self):
        bg_portada = "#0F172A"  # Fondo oscuro elegante para la portada

        canvas = tk.Canvas(self.tab_portada, bg=bg_portada, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            self.tab_portada, orient="vertical", command=canvas.yview
        )
        scroll_frame = tk.Frame(canvas, bg=bg_portada)

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(window_id, width=event.width)

        scroll_frame.bind("<Configure>", _on_frame_configure)
        window_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        main_card = tk.Frame(
            scroll_frame,
            bg="#1E293B",
            bd=1,
            relief="solid",
            highlightbackground="#334155",
        )
        main_card.pack(fill="x", expand=True, padx=40, pady=20)

        # Cabecera
        header_frame = tk.Frame(main_card, bg="#1E293B")
        header_frame.pack(anchor="center", padx=30, pady=20)

        logo_label = tk.Label(header_frame, bg="#1E293B")
        self.logo_img = self.cargar_e_imprimir_imagen(
            "logo_universidad.png", max_ancho=120, max_alto=120
        )

        if self.logo_img:
            logo_label.config(image=self.logo_img)
        else:
            logo_label.config(
                text="[ LOGO ]\nSAN MARCOS",
                font=("Segoe UI", 9, "italic"),
                bg="#334155",
                fg="#94A3B8",
                width=16,
                height=6,
            )
        logo_label.pack(side="left", padx=(0, 20))

        uni_info_frame = tk.Frame(header_frame, bg="#1E293B")
        uni_info_frame.pack(side="left", fill="y")

        tk.Label(
            uni_info_frame,
            text="UNIVERSIDAD NACIONAL MAYOR DE SAN MARCOS",
            font=("Segoe UI", 14, "bold"),
            fg="#F8FAFC",
            bg="#1E293B",
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            uni_info_frame,
            text="FACULTAD DE INGENIERÍA INDUSTRIAL",
            font=("Segoe UI", 11, "bold"),
            fg="#CBD5E1",
            bg="#1E293B",
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            uni_info_frame,
            text=(
                "ESCUELA PROFESIONAL DE INGENIERÍA DE SEGURIDAD Y SALUD EN EL"
                " TRABAJO"
            ),
            font=("Segoe UI", 9),
            fg="#94A3B8",
            bg="#1E293B",
            anchor="w",
        ).pack(anchor="w")

        ttk.Separator(main_card, orient="horizontal").pack(
            fill="x", padx=40, pady=10
        )

        # Título del Trabajo
        proyecto_title = tk.Label(
            main_card,
            text="PLAN ANUAL DE SEGURIDAD Y SALUD EN EL TRABAJO",
            font=("Segoe UI", 16, "bold"),
            fg="#60A5FA",
            bg="#1E293B",
        )
        proyecto_title.pack(pady=(10, 2))

        sub_title = tk.Label(
            main_card,
            text="Sistema de Gestión de Prevención de Riesgos Laborales",
            font=("Segoe UI", 10, "italic"),
            fg="#94A3B8",
            bg="#1E293B",
        )
        sub_title.pack(pady=(0, 15))

        # Asignatura y Docente
        info_academic_frame = tk.Frame(
            main_card, bg="#0F172A", bd=1, relief="solid"
        )
        info_academic_frame.pack(fill="x", padx=50, pady=10)

        tk.Label(
            info_academic_frame,
            text=(
                "💻 ASIGNATURA: PREVENCIÓN Y CONTROL DE RIESGOS LABORALES Y"
                " AMBIENTALES"
            ),
            font=("Segoe UI", 10, "bold"),
            bg="#0F172A",
            fg="#F8FAFC",
        ).pack(anchor="w", padx=20, pady=(10, 3))
        tk.Label(
            info_academic_frame,
            text="👨‍🏫 PROFESOR: Ing. ROJAS LINARES, EDITO LUIS",
            font=("Segoe UI", 10, "bold"),
            bg="#0F172A",
            fg="#F8FAFC",
        ).pack(anchor="w", padx=20, pady=(0, 10))

        # Integrantes
        integrantes_frame = tk.LabelFrame(
            main_card,
            text=" Equipo de Trabajo / Integrantes ",
            font=("Segoe UI", 10, "bold"),
            bg="#1E293B",
            fg="#F8FAFC",
            padx=15,
            pady=15,
        )
        integrantes_frame.pack(fill="x", padx=50, pady=15)

        datos_integrantes = [
            {"foto": "foto1.png", "nombre": "Chaparrea Andrade Ivan", "codigo": "Código: 23170283"},
            {"foto": "foto2.png", "nombre": "Integrante 2", "codigo": "Código: 20260002"},
            {"foto": "foto3.png", "nombre": "Integrante 3", "codigo": "Código: 20260003"},
            {"foto": "foto4.png", "nombre": "Integrante 4", "codigo": "Código: 20260004"},
            {"foto": "foto5.png", "nombre": "Integrante 5", "codigo": "Código: 20260005"},
        ]

        grid_frame = tk.Frame(integrantes_frame, bg="#1E293B")
        grid_frame.pack(expand=True)

        for idx, persona in enumerate(datos_integrantes):
            col_frame = tk.Frame(grid_frame, bg="#1E293B", padx=12)
            col_frame.grid(row=0, column=idx, padx=8, pady=5)

            img_label = tk.Label(col_frame, bg="#1E293B")
            foto_obj = self.cargar_e_imprimir_imagen(
                persona["foto"], max_ancho=75, max_alto=75
            )

            if foto_obj:
                self.fotos_integrantes.append(foto_obj)
                img_label.config(image=foto_obj)
            else:
                img_label.config(
                    text="[ FOTO ]",
                    font=("Segoe UI", 8),
                    bg="#334155",
                    fg="#94A3B8",
                    width=10,
                    height=5,
                )
            img_label.pack()

            tk.Label(
                col_frame,
                text=persona["nombre"],
                font=("Segoe UI", 9, "bold"),
                bg="#1E293B",
                fg="#F8FAFC",
            ).pack(pady=(4, 0))
            tk.Label(
                col_frame,
                text=persona["codigo"],
                font=("Segoe UI", 8),
                bg="#1E293B",
                fg="#94A3B8",
            ).pack()

        btn_comenzar = tk.Button(
            main_card,
            text="Ir a Documentación del Plan SST ➔",
            font=("Segoe UI", 10, "bold"),
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            padx=20,
            pady=10,
            bd=0,
            cursor="hand2",
            command=lambda: self.notebook.select(self.tab_plan),
        )
        btn_comenzar.pack(pady=25)

    # --- CÓDIGO DEL VISOR AUTOMÁTICO DE PDF (PLAN ANUAL) ---
    def build_plan_anual_pdf_viewer(self):
        toolbar = tk.Frame(self.tab_plan, bg="#1E293B", padx=10, pady=6)
        toolbar.pack(fill="x", side="top")

        btn_externo = tk.Button(
            toolbar,
            text="🖥️ Abrir en Visor del Sistema",
            font=("Segoe UI", 9, "bold"),
            bg="#10B981",
            fg="#FFFFFF",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
            command=lambda: self.abrir_archivo_externo(self.pdf_plan_filename),
        )
        btn_externo.pack(side="left", padx=5)

        self.btn_prev_plan = tk.Button(
            toolbar,
            text="◀ Anterior",
            font=("Segoe UI", 9),
            bg="#475569",
            fg="#FFFFFF",
            bd=0,
            padx=10,
            pady=5,
            state="disabled",
            command=self.pagina_anterior_plan,
        )
        self.btn_prev_plan.pack(side="left", padx=(20, 2))

        self.lbl_page_num_plan = tk.Label(
            toolbar,
            text="Página: 0 / 0",
            font=("Segoe UI", 9, "bold"),
            bg="#1E293B",
            fg="#FFFFFF",
            padx=10,
        )
        self.lbl_page_num_plan.pack(side="left")

        self.btn_next_plan = tk.Button(
            toolbar,
            text="Siguiente ▶",
            font=("Segoe UI", 9),
            bg="#475569",
            fg="#FFFFFF",
            bd=0,
            padx=10,
            pady=5,
            state="disabled",
            command=self.pagina_siguiente_plan,
        )
        self.btn_next_plan.pack(side="left", padx=2)

        pdf_container = tk.Frame(self.tab_plan, bg="#334155")
        pdf_container.pack(fill="both", expand=True)

        v_scroll = ttk.Scrollbar(pdf_container, orient="vertical")
        h_scroll = ttk.Scrollbar(pdf_container, orient="horizontal")

        self.pdf_canvas_plan = tk.Canvas(
            pdf_container,
            bg="#525659",
            xscrollcommand=h_scroll.set,
            yscrollcommand=v_scroll.set,
        )

        v_scroll.config(command=self.pdf_canvas_plan.yview)
        h_scroll.config(command=self.pdf_canvas_plan.xview)

        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.pdf_canvas_plan.pack(side="left", fill="both", expand=True)

        self.cargar_pdf_plan_automatico()

    def cargar_pdf_plan_automatico(self):
        if os.path.exists(self.pdf_plan_filename):
            try:
                self.pdf_plan_doc = pymupdf.open(self.pdf_plan_filename)
                self.current_page_plan = 0
                self.mostrar_pagina_pdf_plan()
            except Exception as e:
                messagebox.showerror(
                    "Error al leer PDF", f"No se pudo cargar el archivo PDF:\n{e}"
                )
        else:
            self.pdf_canvas_plan.delete("all")
            self.pdf_canvas_plan.create_text(
                300,
                150,
                text=(
                    f"⚠️ No se encontró el archivo '{self.pdf_plan_filename}'\n\n"
                    "Guarda tu archivo PDF en esta misma carpeta con el nombre:\n"
                    f"👉 {self.pdf_plan_filename}"
                ),
                fill="#FFFFFF",
                font=("Segoe UI", 11, "bold"),
                justify="center",
            )

    def mostrar_pagina_pdf_plan(self):
        if not self.pdf_plan_doc:
            return

        page = self.pdf_plan_doc.load_page(self.current_page_plan)
        zoom = 1.3
        mat = pymupdf.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.pdf_image_tk_plan = ImageTk.PhotoImage(img)

        self.pdf_canvas_plan.delete("all")
        self.pdf_canvas_plan.create_image(
            0, 0, image=self.pdf_image_tk_plan, anchor="nw"
        )
        self.pdf_canvas_plan.config(scrollregion=(0, 0, pix.width, pix.height))

        total_paginas = len(self.pdf_plan_doc)
        self.lbl_page_num_plan.config(
            text=f"Página: {self.current_page_plan + 1} / {total_paginas}"
        )

        self.btn_prev_plan.config(
            state="normal" if self.current_page_plan > 0 else "disabled"
        )
        self.btn_next_plan.config(
            state="normal"
            if self.current_page_plan < total_paginas - 1
            else "disabled"
        )

    def pagina_anterior_plan(self):
        if self.pdf_plan_doc and self.current_page_plan > 0:
            self.current_page_plan -= 1
            self.mostrar_pagina_pdf_plan()

    def pagina_siguiente_plan(self):
        if self.pdf_plan_doc and self.current_page_plan < len(self.pdf_plan_doc) - 1:
            self.current_page_plan += 1
            self.mostrar_pagina_pdf_plan()

    # --- SECCIÓN Y VISOR DE MATRIZ IPERC (EXCEL Y PDF) ---
    def build_iperc_viewer(self):
        toolbar = tk.Frame(self.tab_matriz_iperc, bg="#1E293B", padx=10, pady=6)
        toolbar.pack(fill="x", side="top")

        # Botón para Excel
        btn_excel = tk.Button(
            toolbar,
            text="📊 Abrir Matriz en Excel (.XLSX)",
            font=("Segoe UI", 9, "bold"),
            bg="#16A34A",
            fg="#FFFFFF",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
            command=lambda: self.abrir_archivo_externo(self.excel_iperc_filename),
        )
        btn_excel.pack(side="left", padx=5)

        # Botón para PDF Externo (si se guardó como PDF)
        btn_externo = tk.Button(
            toolbar,
            text="📄 Abrir IPERC en PDF Externo",
            font=("Segoe UI", 9, "bold"),
            bg="#10B981",
            fg="#FFFFFF",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
            command=lambda: self.abrir_archivo_externo(self.pdf_iperc_filename),
        )
        btn_externo.pack(side="left", padx=5)

        self.btn_prev_iperc = tk.Button(
            toolbar,
            text="◀ Anterior",
            font=("Segoe UI", 9),
            bg="#475569",
            fg="#FFFFFF",
            bd=0,
            padx=10,
            pady=5,
            state="disabled",
            command=self.pagina_anterior_iperc,
        )
        self.btn_prev_iperc.pack(side="left", padx=(20, 2))

        self.lbl_page_num_iperc = tk.Label(
            toolbar,
            text="Página: 0 / 0",
            font=("Segoe UI", 9, "bold"),
            bg="#1E293B",
            fg="#FFFFFF",
            padx=10,
        )
        self.lbl_page_num_iperc.pack(side="left")

        self.btn_next_iperc = tk.Button(
            toolbar,
            text="Siguiente ▶",
            font=("Segoe UI", 9),
            bg="#475569",
            fg="#FFFFFF",
            bd=0,
            padx=10,
            pady=5,
            state="disabled",
            command=self.pagina_siguiente_iperc,
        )
        self.btn_next_iperc.pack(side="left", padx=2)

        pdf_container = tk.Frame(self.tab_matriz_iperc, bg="#334155")
        pdf_container.pack(fill="both", expand=True)

        v_scroll = ttk.Scrollbar(pdf_container, orient="vertical")
        h_scroll = ttk.Scrollbar(pdf_container, orient="horizontal")

        self.pdf_canvas_iperc = tk.Canvas(
            pdf_container,
            bg="#525659",
            xscrollcommand=h_scroll.set,
            yscrollcommand=v_scroll.set,
        )

        v_scroll.config(command=self.pdf_canvas_iperc.yview)
        h_scroll.config(command=self.pdf_canvas_iperc.xview)

        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.pdf_canvas_iperc.pack(side="left", fill="both", expand=True)

        self.cargar_pdf_iperc_automatico()

    def cargar_pdf_iperc_automatico(self):
        if os.path.exists(self.pdf_iperc_filename):
            try:
                self.pdf_iperc_doc = pymupdf.open(self.pdf_iperc_filename)
                self.current_page_iperc = 0
                self.mostrar_pagina_pdf_iperc()
            except Exception as e:
                messagebox.showerror(
                    "Error al leer PDF", f"No se pudo cargar el archivo PDF:\n{e}"
                )
        else:
            self.pdf_canvas_iperc.delete("all")
            self.pdf_canvas_iperc.create_text(
                350,
                150,
                text=(
                    f"⚠️ Archivo principal: '{self.excel_iperc_filename}'\n\n"
                    "Haz clic en 'Abrir Matriz en Excel' para ver/editar la matriz completa.\n\n"
                    "Opcional: Si deseas ver una vista previa gráfica aquí dentro,\n"
                    f"guarda una copia en PDF con el nombre: 👉 {self.pdf_iperc_filename}"
                ),
                fill="#FFFFFF",
                font=("Segoe UI", 11, "bold"),
                justify="center",
            )

    def mostrar_pagina_pdf_iperc(self):
        if not self.pdf_iperc_doc:
            return

        page = self.pdf_iperc_doc.load_page(self.current_page_iperc)
        zoom = 1.3
        mat = pymupdf.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.pdf_image_tk_iperc = ImageTk.PhotoImage(img)

        self.pdf_canvas_iperc.delete("all")
        self.pdf_canvas_iperc.create_image(
            0, 0, image=self.pdf_image_tk_iperc, anchor="nw"
        )
        self.pdf_canvas_iperc.config(scrollregion=(0, 0, pix.width, pix.height))

        total_paginas = len(self.pdf_iperc_doc)
        self.lbl_page_num_iperc.config(
            text=f"Página: {self.current_page_iperc + 1} / {total_paginas}"
        )

        self.btn_prev_iperc.config(
            state="normal" if self.current_page_iperc > 0 else "disabled"
        )
        self.btn_next_iperc.config(
            state="normal"
            if self.current_page_iperc < total_paginas - 1
            else "disabled"
        )

    def pagina_anterior_iperc(self):
        if self.pdf_iperc_doc and self.current_page_iperc > 0:
            self.current_page_iperc -= 1
            self.mostrar_pagina_pdf_iperc()

    def pagina_siguiente_iperc(self):
        if self.pdf_iperc_doc and self.current_page_iperc < len(self.pdf_iperc_doc) - 1:
            self.current_page_iperc += 1
            self.mostrar_pagina_pdf_iperc()

    # --- SECCIÓN Y VISOR DE PLANO DE EVACUACIÓN (AUTOCAD Y PDF) ---
    def build_plano_evacuacion_viewer(self):
        toolbar = tk.Frame(self.tab_plano_evacuacion, bg="#1E293B", padx=10, pady=6)
        toolbar.pack(fill="x", side="top")

        # Botón para AutoCAD
        btn_autocad = tk.Button(
            toolbar,
            text="📐 Abrir Plano en AutoCAD (.DWG)",
            font=("Segoe UI", 9, "bold"),
            bg="#D97706",
            fg="#FFFFFF",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
            command=lambda: self.abrir_archivo_externo(self.dwg_plano_filename),
        )
        btn_autocad.pack(side="left", padx=5)

        # Botón para PDF Externo
        btn_externo = tk.Button(
            toolbar,
            text="📄 Abrir Plano PDF Externo",
            font=("Segoe UI", 9, "bold"),
            bg="#10B981",
            fg="#FFFFFF",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
            command=lambda: self.abrir_archivo_externo(self.pdf_plano_filename),
        )
        btn_externo.pack(side="left", padx=5)

        self.btn_prev_plano = tk.Button(
            toolbar,
            text="◀ Anterior",
            font=("Segoe UI", 9),
            bg="#475569",
            fg="#FFFFFF",
            bd=0,
            padx=10,
            pady=5,
            state="disabled",
            command=self.pagina_anterior_plano,
        )
        self.btn_prev_plano.pack(side="left", padx=(20, 2))

        self.lbl_page_num_plano = tk.Label(
            toolbar,
            text="Página: 0 / 0",
            font=("Segoe UI", 9, "bold"),
            bg="#1E293B",
            fg="#FFFFFF",
            padx=10,
        )
        self.lbl_page_num_plano.pack(side="left")

        self.btn_next_plano = tk.Button(
            toolbar,
            text="Siguiente ▶",
            font=("Segoe UI", 9),
            bg="#475569",
            fg="#FFFFFF",
            bd=0,
            padx=10,
            pady=5,
            state="disabled",
            command=self.pagina_siguiente_plano,
        )
        self.btn_next_plano.pack(side="left", padx=2)

        pdf_container = tk.Frame(self.tab_plano_evacuacion, bg="#334155")
        pdf_container.pack(fill="both", expand=True)

        v_scroll = ttk.Scrollbar(pdf_container, orient="vertical")
        h_scroll = ttk.Scrollbar(pdf_container, orient="horizontal")

        self.pdf_canvas_plano = tk.Canvas(
            pdf_container,
            bg="#525659",
            xscrollcommand=h_scroll.set,
            yscrollcommand=v_scroll.set,
        )

        v_scroll.config(command=self.pdf_canvas_plano.yview)
        h_scroll.config(command=self.pdf_canvas_plano.xview)

        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.pdf_canvas_plano.pack(side="left", fill="both", expand=True)

        self.cargar_pdf_plano_automatico()

    def cargar_pdf_plano_automatico(self):
        if os.path.exists(self.pdf_plano_filename):
            try:
                self.pdf_plano_doc = pymupdf.open(self.pdf_plano_filename)
                self.current_page_plano = 0
                self.mostrar_pagina_pdf_plano()
            except Exception as e:
                messagebox.showerror(
                    "Error al leer PDF", f"No se pudo cargar el archivo PDF:\n{e}"
                )
        else:
            self.pdf_canvas_plano.delete("all")
            self.pdf_canvas_plano.create_text(
                350,
                150,
                text=(
                    f"⚠️ No se encontró el archivo '{self.pdf_plano_filename}'\n\n"
                    "Para ver la vista previa en la aplicación, guarda el PDF como:\n"
                    f"👉 {self.pdf_plano_filename}\n\n"
                    "Para abrirlo en AutoCAD, guarda el plano como:\n"
                    f"👉 {self.dwg_plano_filename}"
                ),
                fill="#FFFFFF",
                font=("Segoe UI", 11, "bold"),
                justify="center",
            )

    def mostrar_pagina_pdf_plano(self):
        if not self.pdf_plano_doc:
            return

        page = self.pdf_plano_doc.load_page(self.current_page_plano)
        zoom = 1.3
        mat = pymupdf.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)

        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.pdf_image_tk_plano = ImageTk.PhotoImage(img)

        self.pdf_canvas_plano.delete("all")
        self.pdf_canvas_plano.create_image(
            0, 0, image=self.pdf_image_tk_plano, anchor="nw"
        )
        self.pdf_canvas_plano.config(scrollregion=(0, 0, pix.width, pix.height))

        total_paginas = len(self.pdf_plano_doc)
        self.lbl_page_num_plano.config(
            text=f"Página: {self.current_page_plano + 1} / {total_paginas}"
        )

        self.btn_prev_plano.config(
            state="normal" if self.current_page_plano > 0 else "disabled"
        )
        self.btn_next_plano.config(
            state="normal"
            if self.current_page_plano < total_paginas - 1
            else "disabled"
        )

    def pagina_anterior_plano(self):
        if self.pdf_plano_doc and self.current_page_plano > 0:
            self.current_page_plano -= 1
            self.mostrar_pagina_pdf_plano()

    def pagina_siguiente_plano(self):
        if self.pdf_plano_doc and self.current_page_plano < len(self.pdf_plano_doc) - 1:
            self.current_page_plano += 1
            self.mostrar_pagina_pdf_plano()

    # --- FUNCIÓN GENERAL PARA ABRIR ARCHIVOS EXTERNOS (.PDF, .XLSX, .DWG, ETC.) ---
    def abrir_archivo_externo(self, filepath):
        if os.path.exists(filepath):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(filepath)
                elif sys.platform.startswith("darwin"):
                    subprocess.run(["open", filepath])
                else:
                    subprocess.run(["xdg-open", filepath])
            except Exception as e:
                messagebox.showerror(
                    "Error al abrir archivo",
                    f"No se pudo abrir el archivo '{filepath}':\n{e}",
                )
        else:
            messagebox.showwarning(
                "Archivo no encontrado",
                f"No se encontró el archivo '{filepath}' en la carpeta del programa.",
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = PlanAnualSSTApp(root)
    root.mainloop()

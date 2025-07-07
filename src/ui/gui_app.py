# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, scrolledtext


class CustomDialog(tk.Toplevel):
    """Caixa de diálogo para adicionar um novo funcionário com nome e cargo."""

    def __init__(self, parent, cargos):
        super().__init__(parent)
        self.transient(parent)
        self.title("Novo Funcionário")
        self.parent = parent
        self.result = None

        # Estilo
        self.configure(bg="#F0F2F5")

        body = ttk.Frame(self, padding="10")
        self.initial_focus = self.body(body, cargos)
        body.pack(padx=10, pady=10)

        self.buttonbox()
        self.grab_set()

        if not self.initial_focus:
            self.initial_focus = self

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.geometry(f"+{parent.winfo_rootx()+50}+{parent.winfo_rooty()+50}")
        self.initial_focus.focus_set()
        self.wait_window(self)

    def body(self, master, cargos):
        ttk.Label(master, text="Nome:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.nome_entry = ttk.Entry(master, width=30)
        self.nome_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(master, text="Cargo:").grid(row=1, column=0, sticky=tk.W)
        self.cargo_combo = ttk.Combobox(
            master, values=cargos, state="readonly", width=28
        )
        self.cargo_combo.grid(row=1, column=1, padx=5, pady=5)
        if cargos:
            self.cargo_combo.current(0)

        return self.nome_entry

    def buttonbox(self):
        box = ttk.Frame(self)
        w = ttk.Button(
            box, text="OK", width=10, command=self.ok, style="Accent.TButton"
        )
        w.pack(side=tk.LEFT, padx=5, pady=5)
        w = ttk.Button(box, text="Cancelar", width=10, command=self.cancel)
        w.pack(side=tk.LEFT, padx=5, pady=5)
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
        box.pack(pady=5)

    def ok(self, event=None):
        nome = self.nome_entry.get().strip()
        cargo = self.cargo_combo.get()
        if not nome or not cargo:
            messagebox.showwarning(
                "Entrada Inválida", "Nome e cargo são obrigatórios.", parent=self
            )
            return
        self.result = (nome, cargo)
        self.withdraw()
        self.update_idletasks()
        self.cancel()

    def cancel(self, event=None):
        self.parent.focus_set()
        self.destroy()


class GuiApp(tk.Tk):
    """
    Interface gráfica com gestão de Lotes e alocação de horas por pessoa.
    """

    def __init__(self, app_controller):
        super().__init__()
        self.app_controller = app_controller
        self.title("Planejador de Recursos por Lotes")
        self.geometry("1600x900")
        self.configure(bg="#F0F2F5")

        self.setup_styles()

        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        control_panel = ttk.Frame(main_frame, width=280, style="Modern.TFrame")
        control_panel.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        control_panel.grid_propagate(False)

        content_area = ttk.Frame(main_frame, style="Modern.TFrame")
        content_area.grid(row=0, column=1, sticky="nsew")
        content_area.grid_rowconfigure(0, weight=1)
        content_area.grid_columnconfigure(0, weight=1)

        # --- Painel de Controlo ---
        ttk.Button(
            control_panel,
            text="CALCULAR ALOCAÇÃO",
            command=self.processar_calculo,
            style="Accent.TButton",
        ).pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=20, ipady=10)

        setup_frame = ttk.LabelFrame(control_panel, text="Configuração", padding=10)
        setup_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        func_frame = ttk.LabelFrame(control_panel, text="Equipa", padding=10)
        func_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # --- Conteúdo dos Frames do Painel de Controlo ---
        listbox_container = ttk.Frame(func_frame)
        listbox_container.pack(fill=tk.BOTH, expand=True)

        func_scrollbar = ttk.Scrollbar(listbox_container, orient="vertical")
        self.func_listbox = tk.Listbox(
            listbox_container,
            yscrollcommand=func_scrollbar.set,
            bg="white",
            borderwidth=0,
            highlightthickness=0,
        )
        func_scrollbar.config(command=self.func_listbox.yview)

        func_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.func_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(func_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(
            btn_frame, text="Adicionar", command=self.adicionar_funcionario
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        ttk.Button(btn_frame, text="Remover", command=self.remover_funcionario).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0)
        )

        ttk.Label(setup_frame, text="Horas Trabalháveis/Mês:").pack(
            anchor=tk.W, pady=(0, 2)
        )
        self.horas_mes_entry = ttk.Entry(setup_frame)
        self.horas_mes_entry.insert(0, "160")
        self.horas_mes_entry.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(setup_frame, text="Número de Lotes:").pack(anchor=tk.W, pady=(0, 2))
        self.num_lotes_entry = ttk.Entry(setup_frame)
        self.num_lotes_entry.insert(0, "3")
        self.num_lotes_entry.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            setup_frame, text="Gerar Abas de Lotes", command=self.gerar_abas_lotes
        ).pack(fill=tk.X)

        # --- Área de Conteúdo ---
        self.notebook = ttk.Notebook(content_area)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        dashboard_tab = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(dashboard_tab, text="Dashboard Consolidado")
        self.dash_tree = ttk.Treeview(dashboard_tab, style="Modern.Treeview")
        self.dash_tree.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)

        self.lote_widgets = {}

    def setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        BG_COLOR = "#F0F2F5"
        PRIMARY_COLOR = "#0078D7"
        TEXT_COLOR = "#212121"
        LIGHT_TEXT_COLOR = "#616161"
        BORDER_COLOR = "#CCCCCC"

        style.configure(
            ".", background=BG_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 10)
        )
        style.configure("TFrame", background=BG_COLOR)
        style.configure(
            "Modern.TFrame",
            background="white",
            relief="solid",
            borderwidth=1,
            bordercolor=BORDER_COLOR,
        )

        style.configure("TLabel", background=BG_COLOR)
        style.configure("TLabelframe", background=BG_COLOR, bordercolor=BORDER_COLOR)
        style.configure(
            "TLabelframe.Label", background=BG_COLOR, font=("Segoe UI", 11, "bold")
        )

        style.configure(
            "TButton",
            padding=6,
            relief="flat",
            background="#E1E1E1",
            font=("Segoe UI", 9),
        )
        style.map("TButton", background=[("active", "#CFCFCF")])

        style.configure(
            "Accent.TButton",
            foreground="white",
            background=PRIMARY_COLOR,
            font=("Segoe UI", 10, "bold"),
        )
        style.map("Accent.TButton", background=[("active", "#005A9E")])

        style.configure("TNotebook", background=BG_COLOR, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            padding=[10, 5],
            font=("Segoe UI", 10),
            background=["#DCE4EC"],
            foreground=[LIGHT_TEXT_COLOR],
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "white")],
            foreground=[("selected", PRIMARY_COLOR)],
        )

        style.configure(
            "Modern.Treeview", background="white", fieldbackground="white", rowheight=25
        )
        style.configure(
            "Modern.Treeview.Heading", font=("Segoe UI", 10, "bold"), padding=5
        )
        style.map("Modern.Treeview.Heading", background=[("active", "#E5E5E5")])

    def adicionar_funcionario(self):
        dialog = CustomDialog(self, self.app_controller.cargos_disponiveis)
        if dialog.result:
            nome, cargo = dialog.result
            self.app_controller.adicionar_funcionario(nome, cargo)

    def remover_funcionario(self):
        selecionados = self.func_listbox.curselection()
        if selecionados:
            display_string = self.func_listbox.get(selecionados[0])
            self.app_controller.remover_funcionario(display_string)

    def atualizar_lista_funcionarios(self, funcionarios_display):
        self.func_listbox.delete(0, tk.END)
        for item in funcionarios_display:
            self.func_listbox.insert(tk.END, item)
        for lote in self.lote_widgets.values():
            for disc in lote.values():
                for aloc_widget in disc["alocacoes_widgets"]:
                    if aloc_widget["combo"].winfo_exists():
                        aloc_widget["combo"].config(values=funcionarios_display)

    def gerar_abas_lotes(self):
        for i in range(len(self.notebook.tabs()) - 1, 0, -1):
            self.notebook.forget(i)
        self.lote_widgets = {}

        try:
            num_lotes = int(self.num_lotes_entry.get())
            if num_lotes <= 0:
                messagebox.showwarning(
                    "Aviso", "O número de lotes deve ser maior que zero."
                )
                return
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira um número válido de lotes.")
            return

        for i in range(num_lotes):
            lote_nome = str(i + 1)
            tab = ttk.Frame(self.notebook, padding=10)
            self.notebook.add(tab, text=f"Lote {lote_nome}")
            self.criar_conteudo_aba_lote(tab, lote_nome)

    def criar_conteudo_aba_lote(self, tab, lote_nome):
        canvas = tk.Canvas(tab, highlightthickness=0, bg="#FFFFFF")
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Modern.TFrame")

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_mousewheel(event, widget):
            widget.yview_scroll(-1 * (event.delta // 120), "units")

        canvas.bind(
            "<Enter>",
            lambda e: canvas.bind_all(
                "<MouseWheel>", lambda event: _on_mousewheel(event, canvas)
            ),
        )
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        lote_widgets = {}
        for disc in self.app_controller.disciplinas:
            disc_frame = ttk.LabelFrame(scrollable_frame, text=disc, padding=10)
            disc_frame.pack(fill=tk.X, expand=True, padx=10, pady=10)

            date_frame = ttk.Frame(disc_frame)
            date_frame.pack(fill=tk.X, pady=5)
            ttk.Label(date_frame, text="Início (DD/MM/AAAA):").pack(side=tk.LEFT)
            start_date_entry = ttk.Entry(date_frame, width=15)
            start_date_entry.pack(side=tk.LEFT, padx=5)
            ttk.Label(date_frame, text="Fim (DD/MM/AAAA):").pack(side=tk.LEFT, padx=10)
            end_date_entry = ttk.Entry(date_frame, width=15)
            end_date_entry.pack(side=tk.LEFT, padx=5)

            aloc_frame = ttk.Frame(disc_frame)
            aloc_frame.pack(fill=tk.X, expand=True, pady=5)

            lote_widgets[disc] = {
                "start_date": start_date_entry,
                "end_date": end_date_entry,
                "alocacoes_frame": aloc_frame,
                "alocacoes_widgets": [],
            }
            ttk.Button(
                disc_frame,
                text="+ Alocar Pessoa",
                command=lambda ln=lote_nome, d=disc: self.adicionar_alocacao(ln, d),
            ).pack(anchor=tk.W, pady=5)

        self.lote_widgets[lote_nome] = lote_widgets

    def adicionar_alocacao(self, lote_nome, disciplina):
        frame = self.lote_widgets[lote_nome][disciplina]["alocacoes_frame"]
        row_frame = ttk.Frame(frame)
        row_frame.pack(fill=tk.X, pady=2)

        combo = ttk.Combobox(
            row_frame,
            values=self.app_controller.get_funcionarios_para_display(),
            state="readonly",
            width=30,
        )
        combo.pack(side=tk.LEFT, padx=5)

        ttk.Label(row_frame, text="Horas Totais na Tarefa:").pack(side=tk.LEFT, padx=5)
        entry = ttk.Entry(row_frame, width=10)
        entry.pack(side=tk.LEFT, padx=5)

        monthly_hours_frame = ttk.Frame(row_frame)
        monthly_hours_frame.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        remove_btn = ttk.Button(row_frame, text="X", width=3, command=row_frame.destroy)
        remove_btn.pack(side=tk.RIGHT)

        self.lote_widgets[lote_nome][disciplina]["alocacoes_widgets"].append(
            {
                "combo": combo,
                "entry": entry,
                "frame": row_frame,
                "monthly_hours_frame": monthly_hours_frame,
            }
        )

    def processar_calculo(self):
        try:
            horas_mes = int(self.horas_mes_entry.get() or 160)
            lotes_data = []

            for lote_nome, lote_widgets in self.lote_widgets.items():
                cronograma = {}
                alocacao = {}
                for disc, disc_widgets in lote_widgets.items():
                    cronograma[disc] = {
                        "inicio": disc_widgets["start_date"].get(),
                        "fim": disc_widgets["end_date"].get(),
                    }
                    alocacao[disc] = []
                    for aloc_widget in disc_widgets["alocacoes_widgets"]:
                        if not aloc_widget["frame"].winfo_exists():
                            continue
                        display_string = aloc_widget["combo"].get()
                        valor_horas_totais = float(aloc_widget["entry"].get() or 0.0)
                        if display_string:
                            nome, resto = display_string.rsplit(" (", 1)
                            cargo = resto.rstrip(")")
                            alocacao[disc].append(
                                {
                                    "funcionario": (nome, cargo),
                                    "horas_totais": valor_horas_totais,
                                }
                            )

                lotes_data.append(
                    {"nome": lote_nome, "cronograma": cronograma, "alocacao": alocacao}
                )

            self.app_controller.processar_portfolio(lotes_data, horas_mes)

        except ValueError:
            messagebox.showerror(
                "Erro de Entrada", "Verifique os dados. As horas devem ser um número."
            )
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}")

    def atualizar_dashboard(self, df_relatorio, detalhes_tarefas={}):
        self.notebook.select(0)

        tree = self.dash_tree
        tree.delete(*tree.get_children())

        if df_relatorio.empty:
            return
        if "Erro" in df_relatorio.columns:
            messagebox.showerror("Erro no Processamento", df_relatorio["Erro"].iloc[0])
            return

        tree["columns"] = list(df_relatorio.columns)
        tree["show"] = "headings"
        tree.tag_configure("excedido", background="#ffdddd", foreground="red")
        tree.tag_configure(
            "header", background="#e0e0e0", font=("Helvetica", 10, "bold")
        )

        for col in tree["columns"]:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.W)

        df_sorted = df_relatorio.sort_values(by=["Cargo", "Funcionário"])
        last_cargo = None
        for _, row in df_sorted.iterrows():
            if row["Cargo"] != last_cargo:
                last_cargo = row["Cargo"]
                header_values = [last_cargo.upper()] + [""] * (
                    len(df_relatorio.columns) - 1
                )
                tree.insert("", "end", values=header_values, tags=("header",))

            tags = ("excedido",) if row["Status"] == "Alocação Excedida" else ()
            tree.insert("", "end", values=list(row), tags=tags)

        # Atualiza as abas dos lotes com as horas mensais
        for lote_nome, lote_widgets in self.lote_widgets.items():
            contador_tarefa_disc = {disc: 0 for disc in self.app_controller.disciplinas}
            for disc, disc_widgets in lote_widgets.items():
                for aloc_widget in disc_widgets["alocacoes_widgets"]:
                    if not aloc_widget["frame"].winfo_exists():
                        continue

                    for child in aloc_widget["monthly_hours_frame"].winfo_children():
                        child.destroy()

                    display_string = aloc_widget["combo"].get()
                    if not display_string:
                        continue

                    try:
                        nome, resto = display_string.rsplit(" (", 1)
                        cargo = resto.rstrip(")")
                        pessoa_tuplo = (nome, cargo)

                        tarefa_key = (
                            lote_nome,
                            disc,
                            pessoa_tuplo,
                            contador_tarefa_disc[disc],
                        )
                        contador_tarefa_disc[disc] += 1

                        dados_mensais = detalhes_tarefas.get(tarefa_key)

                        if dados_mensais:
                            texto_horas = " | ".join(
                                [
                                    f"{mes}: {horas}h"
                                    for mes, horas in sorted(dados_mensais.items())
                                ]
                            )
                            ttk.Label(
                                aloc_widget["monthly_hours_frame"],
                                text=texto_horas,
                                foreground="blue",
                            ).pack(side=tk.LEFT)
                    except ValueError:
                        continue

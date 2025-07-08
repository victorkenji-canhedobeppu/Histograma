# -*- coding: utf-8 -*-

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.label import DataLabelList
import datetime
from functools import partial
import json  # NOVO: Import para lidar com o arquivo de estado
import os  # NOVO: Import para verificar a existência do arquivo

ESTADO_APP_FILE = "estado_app.json"  # NOVO: Nome do arquivo para salvar o estado


class CustomDialog(tk.Toplevel):
    # ... (código da classe CustomDialog permanece inalterado)
    def __init__(self, parent, cargos):
        super().__init__(parent)
        self.transient(parent)
        self.title("Novo Funcionário")
        self.parent = parent
        self.result = None
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


class TeamManager(tk.Toplevel):
    # ... (código da classe TeamManager permanece inalterado)
    def __init__(self, parent, app_controller):
        super().__init__(parent)
        self.transient(parent)
        self.app_controller = app_controller
        self.title("Gerenciar Equipe")
        self.geometry("400x500")
        self.resizable(False, False)
        func_frame = ttk.LabelFrame(self, text="Equipe", padding=10)
        func_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        listbox_container = ttk.Frame(func_frame)
        listbox_container.pack(fill=tk.BOTH, expand=True)
        func_scrollbar = ttk.Scrollbar(listbox_container, orient="vertical")
        self.func_listbox = tk.Listbox(
            listbox_container,
            yscrollcommand=func_scrollbar.set,
            bg="white",
            borderwidth=0,
            highlightthickness=0,
            font=("Segoe UI", 10),
        )
        func_scrollbar.config(command=self.func_listbox.yview)
        func_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.func_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        btn_frame = ttk.Frame(func_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        ttk.Button(
            btn_frame, text="Adicionar", command=self.adicionar_funcionario
        ).grid(row=0, column=0, sticky="ew", padx=(0, 2))
        ttk.Button(btn_frame, text="Remover", command=self.remover_funcionario).grid(
            row=0, column=1, sticky="ew", padx=(2, 0)
        )
        self.populate_listbox()
        self.grab_set()

    def populate_listbox(self):
        self.func_listbox.delete(0, tk.END)
        for item in self.app_controller.get_funcionarios_para_display():
            self.func_listbox.insert(tk.END, item)

    def adicionar_funcionario(self):
        dialog = CustomDialog(self, self.app_controller.get_cargos_disponiveis())
        if dialog.result:
            self.app_controller.adicionar_funcionario(*dialog.result)
            self.populate_listbox()

    def remover_funcionario(self):
        selecionados = self.func_listbox.curselection()
        if not selecionados:
            return
        display_string = self.func_listbox.get(selecionados[0])
        if messagebox.askyesno(
            "Confirmar Remoção",
            f"Remover '{display_string}' e todas as suas alocações?",
            parent=self,
        ):
            self.app_controller.ui.remover_alocacoes_de_funcionario(display_string)
            self.app_controller.remover_funcionario(display_string)
            self.populate_listbox()


class GuiApp(tk.Tk):
    def __init__(self, app_controller):
        super().__init__()
        self.app_controller = app_controller
        self.title("Planejador de Recursos por Lotes")
        self.geometry("1600x900")
        self.configure(bg="#F0F2F5")
        self.setup_styles()
        self.team_window = None
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        control_panel = ttk.Frame(main_frame, width=300, style="Modern.TFrame")
        control_panel.grid(row=0, column=0, sticky="ns", padx=(0, 10))
        control_panel.grid_propagate(False)
        content_area = ttk.Frame(main_frame, style="Modern.TFrame")
        content_area.grid(row=0, column=1, sticky="nsew")
        content_area.grid_rowconfigure(0, weight=1)
        content_area.grid_columnconfigure(0, weight=1)
        self.create_control_panel_widgets(control_panel)
        self.notebook = ttk.Notebook(content_area)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        dashboard_tab = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(dashboard_tab, text="Dashboard Consolidado")
        self.dash_tree = self._criar_treeview(dashboard_tab)
        self.lote_widgets = {}

        # NOVO: Interceptar o fechamento da janela para salvar o estado
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _on_closing(self):
        # NOVO: Método que é chamado ao fechar a janela
        self.salvar_estado()
        self.destroy()

    # ... (métodos create_control_panel_widgets, open_team_manager, setup_styles, etc. permanecem iguais)
    def create_control_panel_widgets(self, parent):
        # CORREÇÃO: Estrutura de packing para eliminar o espaço vertical
        ttk.Button(
            parent, text="Gerenciar Equipe", command=self.open_team_manager
        ).pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5), ipady=4)

        setup_frame = ttk.LabelFrame(parent, text="Configuração", padding=10)
        setup_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        # Frame de ação é empacotado no final para ficar embaixo de tudo
        action_frame = ttk.Frame(parent)
        action_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        action_frame.columnconfigure(0, weight=1)

        ttk.Button(
            action_frame,
            text="Calcular Alocação",
            command=self.processar_calculo,
            style="Accent.TButton",
        ).grid(row=0, column=0, sticky="ew", ipady=8, pady=2)
        ttk.Button(
            action_frame,
            text="Exportar para Excel",
            command=self.exportar_para_excel,
            style="Accent.TButton",
        ).grid(row=1, column=0, sticky="ew", ipady=4, pady=2)

        ttk.Label(setup_frame, text="Horas Trabalháveis/Mês:").pack(
            anchor=tk.W, pady=(0, 2)
        )
        self.horas_mes_entry = ttk.Entry(setup_frame)
        self.horas_mes_entry.insert(0, "160")
        self.horas_mes_entry.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(setup_frame, text="Número de Lotes:").pack(anchor=tk.W, pady=(0, 2))
        self.num_lotes_entry = ttk.Entry(setup_frame)
        self.num_lotes_entry.insert(0, "1")
        self.num_lotes_entry.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(
            setup_frame, text="Gerar Abas de Lotes", command=self.gerar_abas_lotes
        ).pack(fill=tk.X)

    def open_team_manager(self):
        if self.team_window and self.team_window.winfo_exists():
            self.team_window.focus()
        else:
            self.team_window = TeamManager(self, self.app_controller)

    def setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        BG_COLOR, PRIMARY_COLOR, TEXT_COLOR, LIGHT_TEXT_COLOR, BORDER_COLOR = (
            "#F0F2F5",
            "#0078D7",
            "#212121",
            "#616161",
            "#CCCCCC",
        )
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
            "TLabelframe.Label", background=BG_COLOR, font=("Segoe UI", 10, "bold")
        )
        style.configure(
            "TButton",
            padding=6,
            relief="flat",
            background="#E1E1E1",
            font=("Segoe UI", 10),
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
        style.configure("TNotebook.Tab", padding=[10, 5], font=("Segoe UI", 10))
        try:
            style.layout(
                "TNotebook.Tab",
                [
                    (
                        "Notebook.tab",
                        {
                            "sticky": "nswe",
                            "children": [
                                (
                                    "Notebook.padding",
                                    {
                                        "side": "top",
                                        "sticky": "nswe",
                                        "children": [
                                            (
                                                "Notebook.label",
                                                {"side": "top", "sticky": ""},
                                            )
                                        ],
                                    },
                                )
                            ],
                        },
                    )
                ],
            )
        except tk.TclError:
            pass
        style.map(
            "TNotebook.Tab",
            background=[("selected", "white"), ("!selected", BG_COLOR)],
            foreground=[("selected", PRIMARY_COLOR), ("!selected", LIGHT_TEXT_COLOR)],
        )
        style.configure(
            "Modern.Treeview", background="white", fieldbackground="white", rowheight=25
        )
        style.configure(
            "Modern.Treeview.Heading", font=("Segoe UI", 10, "bold"), padding=5
        )
        style.map("Modern.Treeview.Heading", background=[("active", "#E5E5E5")])

    def remover_alocacoes_de_funcionario(self, display_string):
        for lote_data in self.lote_widgets.values():
            if "disciplinas" in lote_data:
                for disc_widgets in lote_data["disciplinas"].values():
                    for aloc_widget in disc_widgets["alocacoes_widgets"][:]:
                        if (
                            aloc_widget["frame"].winfo_exists()
                            and aloc_widget.get("type") == "equipe"
                            and aloc_widget["combo"].get() == display_string
                        ):
                            aloc_widget["frame"].destroy()
                            disc_widgets["alocacoes_widgets"].remove(aloc_widget)
        self.processar_calculo()

    def _criar_treeview(self, parent_frame):
        container = ttk.Frame(parent_frame)
        container.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(container, style="Modern.Treeview")
        yscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        return tree

    def atualizar_lista_funcionarios(self, funcionarios_display):
        if self.team_window and self.team_window.winfo_exists():
            self.team_window.populate_listbox()
        for lote_data in self.lote_widgets.values():
            if "disciplinas" in lote_data:
                for disc_widgets in lote_data["disciplinas"].values():
                    for aloc_widget in disc_widgets.get("alocacoes_widgets", []):
                        if (
                            aloc_widget.get("type") == "equipe"
                            and aloc_widget.get("combo")
                            and aloc_widget["combo"].winfo_exists()
                        ):
                            aloc_widget["combo"].config(values=funcionarios_display)

    def gerar_abas_lotes(self):
        for i in range(len(self.notebook.tabs()) - 1, 0, -1):
            self.notebook.forget(i)
        self.lote_widgets.clear()
        try:
            num_lotes = int(self.num_lotes_entry.get())
            if num_lotes <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "Erro", "Por favor, insira um número válido e positivo de lotes."
            )
            return
        for i in range(num_lotes):
            lote_nome = str(i + 1)
            lote_tab_frame = ttk.Frame(self.notebook)
            self.notebook.add(lote_tab_frame, text=f"Lote {lote_nome}")

            # CORREÇÃO: Restaurada a criação das sub-abas para cada lote
            lote_notebook = ttk.Notebook(lote_tab_frame)
            lote_notebook.pack(fill=tk.BOTH, expand=True)
            entrada_tab = ttk.Frame(lote_notebook, padding=10)
            dashboard_lote_tab = ttk.Frame(lote_notebook, padding=5)
            lote_notebook.add(entrada_tab, text="Entrada de Dados")
            lote_notebook.add(dashboard_lote_tab, text="Dashboard do Lote")

            self.lote_widgets[lote_nome] = {
                **self.criar_conteudo_aba_entrada(entrada_tab, lote_nome),
                "dash_tree": self._criar_treeview(dashboard_lote_tab),
            }

    def criar_conteudo_aba_entrada(self, parent_tab, lote_nome):
        canvas = tk.Canvas(parent_tab, highlightthickness=0, bg="#FFFFFF")
        scrollbar = ttk.Scrollbar(parent_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Modern.TFrame")
        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # CORREÇÃO: Scroll localizado para este canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")

        def _bind_scroll(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_scroll(event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_scroll)
        canvas.bind("<Leave>", _unbind_scroll)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # CORREÇÃO: Usando grid para um layout estável
        scrollable_frame.columnconfigure(0, weight=1)

        disciplinas_widgets = {}
        row_counter = 0
        for disc in self.app_controller.get_disciplinas():
            disc_frame = ttk.LabelFrame(scrollable_frame, text=disc, padding=10)
            disc_frame.grid(
                row=row_counter, column=0, sticky="ew", padx=10, pady=(10, 5)
            )
            row_counter += 1

            # Usando grid para layout interno estável
            disc_frame.columnconfigure(0, weight=1)

            date_frame = ttk.Frame(disc_frame)
            date_frame.grid(row=0, column=0, sticky="w", pady=(0, 5))
            ttk.Label(date_frame, text="Início (DD/MM/AAAA):").pack(side=tk.LEFT)
            start_date_entry = ttk.Entry(date_frame, width=15)
            start_date_entry.pack(side=tk.LEFT, padx=5)
            ttk.Label(date_frame, text="Fim (DD/MM/AAAA):").pack(side=tk.LEFT, padx=10)
            end_date_entry = ttk.Entry(date_frame, width=15)
            end_date_entry.pack(side=tk.LEFT, padx=5)

            aloc_frame = ttk.Frame(disc_frame)
            aloc_frame.grid(row=1, column=0, sticky="ew", pady=5)

            disciplinas_widgets[disc] = {
                "start_date": start_date_entry,
                "end_date": end_date_entry,
                "alocacoes_frame": aloc_frame,
                "alocacoes_widgets": [],
            }

            cmd = partial(self.adicionar_alocacao_equipe, lote_nome, disc)
            ttk.Button(disc_frame, text="+ Alocar Equipe", command=cmd).grid(
                row=2, column=0, sticky="w", pady=5
            )

        sub_frame_container = ttk.LabelFrame(
            scrollable_frame, text="Subcontratos", padding=10
        )
        sub_frame_container.grid(
            row=row_counter, column=0, sticky="ew", padx=10, pady=10
        )

        subcontratos_widgets = {"alocacoes_widgets": []}
        ttk.Button(
            sub_frame_container,
            text="+ Adicionar Subcontrato",
            command=lambda: self.adicionar_linha_subcontrato(
                sub_frame_container, subcontratos_widgets
            ),
        ).pack(anchor=tk.W, pady=5)

        return {
            "disciplinas": disciplinas_widgets,
            "subcontratos": subcontratos_widgets,
        }

    # ... (métodos adicionar_alocacao_equipe, adicionar_linha_subcontrato, dividir_tarefa permanecem iguais)
    def adicionar_alocacao_equipe(self, lote_nome, disciplina):
        frame = self.lote_widgets[lote_nome]["disciplinas"][disciplina][
            "alocacoes_frame"
        ]
        row_frame = ttk.Frame(frame)
        row_frame.pack(fill=tk.X, pady=2)
        ttk.Label(row_frame, text="Funcionário:").pack(side=tk.LEFT, padx=(0, 5))
        combo = ttk.Combobox(
            row_frame,
            values=self.app_controller.get_funcionarios_para_display(),
            state="readonly",
            width=30,
        )
        combo.pack(side=tk.LEFT, padx=5)
        ttk.Label(row_frame, text="Horas Totais:").pack(side=tk.LEFT, padx=5)
        entry = ttk.Entry(row_frame, width=10)
        entry.pack(side=tk.LEFT, padx=5)
        action_buttons_frame = ttk.Frame(row_frame)
        action_buttons_frame.pack(side=tk.RIGHT, padx=(10, 0))

        aloc_widget_dict = {
            "type": "equipe",
            "combo": combo,
            "entry": entry,
            "frame": row_frame,
        }

        split_btn = ttk.Button(
            action_buttons_frame,
            text="Dividir",
            width=7,
            command=lambda aw=aloc_widget_dict: self.dividir_tarefa(
                lote_nome, disciplina, aw
            ),
        )
        split_btn.pack(side=tk.LEFT, padx=(0, 2))
        remove_btn = ttk.Button(
            action_buttons_frame, text="X", width=3, command=row_frame.destroy
        )
        remove_btn.pack(side=tk.LEFT)

        self.lote_widgets[lote_nome]["disciplinas"][disciplina][
            "alocacoes_widgets"
        ].append(aloc_widget_dict)

    def adicionar_linha_subcontrato(self, container, widgets_dict):
        row_frame = ttk.Frame(container)
        row_frame.pack(fill=tk.X, padx=10, pady=(5, 0), anchor=tk.N)
        ttk.Label(row_frame, text="Tipo:").pack(side=tk.LEFT)
        combo = ttk.Combobox(
            row_frame,
            values=self.app_controller.get_subcontratos_disponiveis(),
            state="readonly",
            width=20,
        )
        combo.pack(side=tk.LEFT, padx=5)
        ttk.Label(row_frame, text="Início:").pack(side=tk.LEFT, padx=(10, 0))
        start_entry = ttk.Entry(row_frame, width=12)
        start_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(row_frame, text="Fim:").pack(side=tk.LEFT, padx=(10, 0))
        end_entry = ttk.Entry(row_frame, width=12)
        end_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(row_frame, text="Horas Totais:").pack(side=tk.LEFT, padx=(10, 0))
        hours_entry = ttk.Entry(row_frame, width=10)
        hours_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(row_frame, text="X", width=3, command=row_frame.destroy).pack(
            side=tk.RIGHT
        )
        widgets_dict["alocacoes_widgets"].append(
            {
                "type": "subcontrato",
                "combo": combo,
                "start_entry": start_entry,
                "end_entry": end_entry,
                "hours_entry": hours_entry,
                "frame": row_frame,
            }
        )

    def dividir_tarefa(self, lote_nome, disciplina, aloc_origem):
        try:
            horas_originais = float(aloc_origem["entry"].get().replace(",", "."))
            if horas_originais <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Aviso",
                "A tarefa precisa ter horas positivas para ser dividida.",
                parent=self,
            )
            return
        dialog = CustomDialog(self, self.app_controller.get_cargos_disponiveis())
        if not dialog.result:
            return
        novo_nome, novo_cargo = dialog.result
        self.app_controller.adicionar_funcionario(novo_nome, novo_cargo)
        novo_display_string = f"{novo_nome} ({novo_cargo})"
        self.adicionar_alocacao_equipe(lote_nome, disciplina)
        nova_aloc_widget = self.lote_widgets[lote_nome]["disciplinas"][disciplina][
            "alocacoes_widgets"
        ][-1]
        horas_divididas = horas_originais / 2.0
        aloc_origem["entry"].delete(0, tk.END)
        aloc_origem["entry"].insert(0, f"{horas_divididas:.2f}".replace(".", ","))
        nova_aloc_widget["combo"].set(novo_display_string)
        nova_aloc_widget["entry"].delete(0, tk.END)
        nova_aloc_widget["entry"].insert(0, f"{horas_divididas:.2f}".replace(".", ","))

    def _coletar_dados_da_ui(self):
        # NOVO: Método auxiliar para extrair todos os dados da UI
        # A lógica é a mesma de `processar_calculo`, mas retorna um dicionário Python
        try:
            estado_geral = {
                "config": {
                    "horas_mes": int(self.horas_mes_entry.get() or 160),
                    "num_lotes": int(self.num_lotes_entry.get() or 0),
                },
                "funcionarios": self.app_controller.funcionarios,
                "lotes": [],
            }

            for lote_nome, lote_data in self.lote_widgets.items():
                lote_dict = {
                    "nome": lote_nome,
                    "disciplinas": {},
                    "subcontratos": [],
                }

                if "disciplinas" in lote_data:
                    for disc, disc_widgets in lote_data["disciplinas"].items():
                        disciplina_dict = {
                            "cronograma": {
                                "inicio": disc_widgets["start_date"].get(),
                                "fim": disc_widgets["end_date"].get(),
                            },
                            "alocacoes": [],
                        }
                        for aloc_widget in disc_widgets["alocacoes_widgets"]:
                            if (
                                aloc_widget["frame"].winfo_exists()
                                and aloc_widget["combo"].get()
                            ):
                                horas_str = (
                                    aloc_widget["entry"].get().strip().replace(",", ".")
                                )
                                display_string = aloc_widget["combo"].get()
                                nome, resto = display_string.rsplit(" (", 1)
                                cargo = resto.rstrip(")")
                                disciplina_dict["alocacoes"].append(
                                    {
                                        "funcionario": (nome, cargo),
                                        "horas_totais": float(horas_str or 0.0),
                                        "display_string": display_string,
                                    }
                                )
                        lote_dict["disciplinas"][disc] = disciplina_dict

                if "subcontratos" in lote_data:
                    for aloc_widget in lote_data["subcontratos"]["alocacoes_widgets"]:
                        if (
                            aloc_widget["frame"].winfo_exists()
                            and aloc_widget["combo"].get()
                        ):
                            horas_str = (
                                aloc_widget["hours_entry"]
                                .get()
                                .strip()
                                .replace(",", ".")
                            )
                            lote_dict["subcontratos"].append(
                                {
                                    "nome": aloc_widget["combo"].get(),
                                    "inicio": aloc_widget["start_entry"].get(),
                                    "fim": aloc_widget["end_entry"].get(),
                                    "horas_totais": float(horas_str or 0.0),
                                }
                            )
                estado_geral["lotes"].append(lote_dict)

            return estado_geral
        except Exception as e:
            print(f"Erro ao coletar dados da UI: {e}")
            return None

    def salvar_estado(self):
        # NOVO: Salva o estado atual da UI em um arquivo JSON
        estado_para_salvar = self._coletar_dados_da_ui()
        if estado_para_salvar:
            try:
                with open(ESTADO_APP_FILE, "w", encoding="utf-8") as f:
                    json.dump(estado_para_salvar, f, ensure_ascii=False, indent=4)
                print("Estado salvo com sucesso.")
            except Exception as e:
                print(f"Erro ao salvar o estado: {e}")

    def carregar_estado(self):
        # NOVO: Carrega o estado de um arquivo JSON e popula a UI
        if not os.path.exists(ESTADO_APP_FILE):
            print("Nenhum arquivo de estado encontrado. Iniciando com padrão.")
            return

        try:
            with open(ESTADO_APP_FILE, "r", encoding="utf-8") as f:
                estado = json.load(f)

            # 1. Restaurar configurações gerais
            self.horas_mes_entry.delete(0, tk.END)
            self.horas_mes_entry.insert(0, str(estado["config"]["horas_mes"]))
            self.num_lotes_entry.delete(0, tk.END)
            self.num_lotes_entry.insert(0, str(estado["config"]["num_lotes"]))

            # 2. Restaurar funcionários
            for nome, cargo in estado["funcionarios"]:
                self.app_controller.adicionar_funcionario(nome, cargo)

            # 3. Gerar abas de lotes
            self.gerar_abas_lotes()

            # 4. Popular dados dos lotes
            for lote_data in estado["lotes"]:
                lote_nome = lote_data["nome"]
                if lote_nome in self.lote_widgets:
                    # Popula disciplinas
                    for disc_nome, disc_data in lote_data["disciplinas"].items():
                        disc_widgets = self.lote_widgets[lote_nome]["disciplinas"][
                            disc_nome
                        ]
                        disc_widgets["start_date"].insert(
                            0, disc_data["cronograma"]["inicio"]
                        )
                        disc_widgets["end_date"].insert(
                            0, disc_data["cronograma"]["fim"]
                        )

                        for aloc_data in disc_data["alocacoes"]:
                            self.adicionar_alocacao_equipe(lote_nome, disc_nome)
                            nova_aloc_widget = disc_widgets["alocacoes_widgets"][-1]
                            nova_aloc_widget["combo"].set(aloc_data["display_string"])
                            nova_aloc_widget["entry"].insert(
                                0, str(aloc_data["horas_totais"]).replace(".", ",")
                            )

                    # Popula subcontratos
                    sub_container = self.lote_widgets[lote_nome]["subcontratos"][
                        "alocacoes_widgets"
                    ]
                    sub_frame_container = next(
                        fw
                        for fw in self.lote_widgets[lote_nome].values()
                        if isinstance(fw, dict) and "alocacoes_widgets" in fw
                    ).get(
                        "frame"
                    )  # Heuristica para achar o frame

                    for sub_data in lote_data["subcontratos"]:
                        self.adicionar_linha_subcontrato(
                            self.lote_widgets[lote_nome]["subcontratos"][
                                "alocacoes_widgets"
                            ][0][
                                "frame"
                            ].master,  # Pega o container
                            self.lote_widgets[lote_nome]["subcontratos"],
                        )
                        nova_sub_widget = self.lote_widgets[lote_nome]["subcontratos"][
                            "alocacoes_widgets"
                        ][-1]
                        nova_sub_widget["combo"].set(sub_data["nome"])
                        nova_sub_widget["start_entry"].insert(0, sub_data["inicio"])
                        nova_sub_widget["end_entry"].insert(0, sub_data["fim"])
                        nova_sub_widget["hours_entry"].insert(
                            0, str(sub_data["horas_totais"]).replace(".", ",")
                        )

            print("Estado carregado com sucesso.")
            self.processar_calculo()  # Opcional: calcular ao carregar
        except Exception as e:
            messagebox.showerror(
                "Erro ao Carregar", f"Não foi possível carregar o estado salvo: {e}"
            )

    def processar_calculo(self):
        dados_coletados = self._coletar_dados_da_ui()
        if not dados_coletados:
            messagebox.showerror(
                "Erro de Coleta",
                "Não foi possível coletar os dados da interface para o cálculo.",
            )
            return

        try:
            lotes_data_para_processar = []
            for lote_dict in dados_coletados["lotes"]:
                disciplinas_data = {}
                for disc_nome, disc_valores in lote_dict["disciplinas"].items():
                    disciplinas_data[disc_nome] = {
                        "cronograma": disc_valores["cronograma"],
                        "alocacoes": [
                            {
                                "funcionario": aloc["funcionario"],
                                "horas_totais": aloc["horas_totais"],
                            }
                            for aloc in disc_valores["alocacoes"]
                        ],
                    }

                subcontratos_data = lote_dict["subcontratos"]

                lotes_data_para_processar.append(
                    {
                        "nome": lote_dict["nome"],
                        "disciplinas": disciplinas_data,
                        "subcontratos": subcontratos_data,
                    }
                )

            horas_mes = dados_coletados["config"]["horas_mes"]
            self.app_controller.processar_portfolio(
                lotes_data_para_processar, horas_mes
            )
        except Exception as e:
            messagebox.showerror(
                "Erro Inesperado", f"Ocorreu um erro no processamento: {e}"
            )

    # ... (métodos _preencher_treeview, atualizar_dashboards, exportar_para_excel, etc. permanecem iguais)
    def _preencher_treeview(self, tree, df_relatorio):
        tree.delete(*tree.get_children())
        if df_relatorio.empty:
            return
        tree["columns"] = list(df_relatorio.columns)
        tree["show"] = "headings"
        tree.tag_configure("excedido", background="#ffdddd", foreground="red")
        tree.tag_configure(
            "header", background="#e0e0e0", font=("Segoe UI", 10, "bold")
        )
        for col in tree["columns"]:
            tree.heading(col, text=col)
            width = 180
            if col == "Funcionário":
                width = 200
            elif col == "Disciplina":
                width = 150
            tree.column(col, width=width, anchor=tk.W)
        df_sorted = df_relatorio.sort_values(
            by=["Disciplina", "Cargo/Subcontrato", "Funcionário"]
        )
        last_disc = None
        for _, row in df_sorted.iterrows():
            if row["Disciplina"] != last_disc:
                last_disc = row["Disciplina"]
                tree.insert(
                    "",
                    "end",
                    values=[last_disc.upper()] + [""] * (len(df_relatorio.columns) - 1),
                    tags=("header",),
                )
            tags = (
                ("excedido",)
                if "Status" in row and row["Status"] == "Alocação Excedida"
                else ()
            )
            tree.insert("", "end", values=list(row), tags=tags)

    def atualizar_dashboards(
        self, df_consolidado, dashboards_lotes, detalhes_tarefas=None
    ):
        self.notebook.select(0)
        self._preencher_treeview(self.dash_tree, df_consolidado)
        for nome_lote, df_lote in dashboards_lotes.items():
            if nome_lote in self.lote_widgets and self.lote_widgets[nome_lote].get(
                "dash_tree"
            ):
                self._preencher_treeview(
                    self.lote_widgets[nome_lote]["dash_tree"], df_lote
                )

    def exportar_para_excel(self):
        df_consolidado = self.app_controller.get_ultimo_df_consolidado()
        dashboards_lotes = self.app_controller.get_ultimo_dashboards_lotes()
        if df_consolidado is None:
            messagebox.showwarning(
                "Aviso", "Calcule a alocação primeiro antes de exportar."
            )
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")],
            title="Salvar Relatório de Alocação",
            initialfile=f"Relatorio_Alocacao_{datetime.datetime.now().strftime('%Y-%m-%d')}.xlsx",
        )
        if not filepath:
            return
        try:
            with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
                if "Sheet" in writer.book.sheetnames:
                    writer.book.remove(writer.book["Sheet"])
                self._write_styled_df_to_excel(writer, "Consolidado", df_consolidado)
                if dashboards_lotes:
                    for nome_lote, df_lote in dashboards_lotes.items():
                        self._write_styled_df_to_excel(
                            writer, f"Lote {nome_lote}", df_lote
                        )
                        self._create_chart_for_sheet(
                            writer, f"Lote {nome_lote}", df_lote
                        )
            messagebox.showinfo(
                "Sucesso", f"Relatório exportado com sucesso para:\n{filepath}"
            )
        except Exception as e:
            messagebox.showerror(
                "Erro na Exportação", f"Ocorreu um erro ao salvar o arquivo:\n{e}"
            )

    def _write_styled_df_to_excel(self, writer, sheet_name, df):
        df_para_exportar = df.drop(columns=["Status"], errors="ignore")
        if df_para_exportar.empty:
            df_para_exportar.to_excel(writer, sheet_name=sheet_name, index=False)
            return
        df_sorted = df_para_exportar.sort_values(by=["Disciplina", "Funcionário"])

        ws = writer.book.create_sheet(title=sheet_name)

        ws.sheet_view.showGridLines = False
        header_font = Font(name="Arial", bold=True, color="FFFFFF")
        header_fill = PatternFill(
            start_color="44546A", end_color="44546A", fill_type="solid"
        )
        center_align = Alignment(horizontal="center", vertical="center")
        category_font = Font(name="Arial", bold=True)
        category_fill = PatternFill(
            start_color="E7E6E6", end_color="E7E6E6", fill_type="solid"
        )
        data_font = Font(name="Arial")
        thin_side = Side(border_style="thin", color="BFBFBF")
        full_border = Border(
            left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
        )
        headers = list(df_sorted.columns)
        ws.append(headers)
        ws.row_dimensions[1].height = 25
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
        last_disc = None
        for _, row_data in df_sorted.iterrows():
            if row_data["Disciplina"] != last_disc:
                last_disc = row_data["Disciplina"]
                ws.append([last_disc.upper()])
                cat_row_idx = ws.max_row
                ws.merge_cells(
                    start_row=cat_row_idx,
                    start_column=1,
                    end_row=cat_row_idx,
                    end_column=len(headers),
                )
            ws.append(list(row_data))
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(headers)):
            for cell in row:
                cell.border = full_border
                is_category_row = (
                    cell.row > 1
                    and ws.cell(row=cell.row, column=2).value is None
                    and len(ws[cell.row]) > 1
                )
                if cell.font.bold:
                    if is_category_row:
                        cell.fill = category_fill
                        cell.font = category_font
                else:
                    cell.font = data_font
        for i, col_name in enumerate(df_sorted.columns, 1):
            column_letter = chr(64 + i)
            try:
                max_len = max(
                    df_sorted[col_name].astype(str).map(len).max(), len(col_name)
                )
            except (TypeError, ValueError):
                max_len = len(col_name)
            ws.column_dimensions[column_letter].width = max_len + 4

    def _create_chart_for_sheet(self, writer, sheet_name, df):
        if df.empty:
            return
        ws = writer.sheets[sheet_name]
        df_numeric = df.copy()
        month_cols_formatted = [
            col for col in df.columns if "/" in col and col != "H.mês"
        ]
        if not month_cols_formatted:
            return
        for col in month_cols_formatted:
            df_numeric[col] = pd.to_numeric(
                df_numeric[col].str.replace(",", "."), errors="coerce"
            )
        monthly_decimal_totals = df_numeric[month_cols_formatted].sum()

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.legend = None

        if chart.title:
            chart.title.layout = Layout(manualLayout=ManualLayout(y=0, yMode="edge"))

        data_start_col = len(df.columns) + 3
        chart_data_start_row = 2

        ws.cell(
            row=chart_data_start_row - 1,
            column=data_start_col,
            value="Fonte de Dados para o Gráfico",
        )
        data_rows = [["Mês", "Total Decimal"]] + list(
            zip(monthly_decimal_totals.index, monthly_decimal_totals.values)
        )
        for r_idx, row_data in enumerate(data_rows, start=chart_data_start_row):
            for c_idx, value in enumerate(row_data, start=data_start_col):
                ws.cell(row=r_idx, column=c_idx, value=value)

        values = Reference(
            ws,
            min_col=data_start_col + 1,
            min_row=chart_data_start_row + 1,
            max_row=chart_data_start_row + len(monthly_decimal_totals),
        )
        cats = Reference(
            ws,
            min_col=data_start_col,
            min_row=chart_data_start_row + 1,
            max_row=chart_data_start_row + len(monthly_decimal_totals),
        )

        series = Series(
            values,
            title=ws.cell(row=chart_data_start_row, column=data_start_col + 1).value,
        )
        chart.append(series)
        chart.set_categories(cats)

        chart.data_labels = DataLabelList()
        chart.data_labels.showVal = True
        chart.data_labels.position = "outEnd"
        chart.data_labels.numFmt = "0.00"

        ws.add_chart(chart, f"A{ws.max_row + 5}")

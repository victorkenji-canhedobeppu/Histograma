# -*- coding: utf-8 -*-

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime
from collections import defaultdict


class CustomDialog(tk.Toplevel):
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
    def __init__(self, parent, app_controller):
        super().__init__(parent)
        self.transient(parent)
        self.app_controller = app_controller
        self.title("Gerenciar Equipe")
        self.geometry("400x500")

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
            self.app_controller.ui.remover_alocacoes_do_funcionario(display_string)
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

    def create_control_panel_widgets(self, parent):
        # Frame superior para agrupar inputs e ações
        top_container = ttk.Frame(parent)
        top_container.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        # Botão para gerenciar equipe
        ttk.Button(
            top_container, text="Gerenciar Equipe", command=self.open_team_manager
        ).pack(fill=tk.X, padx=5, pady=5, ipady=4)

        # Frame de Configuração
        setup_frame = ttk.LabelFrame(top_container, text="Configuração", padding=10)
        setup_frame.pack(fill=tk.X, padx=5, pady=5)

        # Widgets dentro do frame de Configuração
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

        # Frame para os botões de ação principais, posicionado logo abaixo da configuração
        action_frame = ttk.Frame(top_container)
        action_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=10)
        action_frame.columnconfigure(0, weight=1)

        # Botões de ação com mesmo estilo e tamanho
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

    def open_team_manager(self):
        if self.team_window and self.team_window.winfo_exists():
            self.team_window.focus()
        else:
            self.team_window = TeamManager(self, self.app_controller)

    def setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        BG_COLOR, PRIMARY_COLOR, TEXT_COLOR = "#F0F2F5", "#0078D7", "#212121"
        LIGHT_TEXT_COLOR, BORDER_COLOR = "#616161", "#CCCCCC"
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
            current_layout = style.layout("TNotebook.Tab")
            style.layout(
                "TNotebook.Tab",
                [elem for elem in current_layout if elem[0] != "TNotebook.focus"],
            )
        except tk.TclError:
            style.configure("TNotebook.Tab", focusthickness=0)
        style.map(
            "TNotebook.Tab",
            background=[("selected", "white"), ("", "#DCE4EC")],
            foreground=[("selected", PRIMARY_COLOR), ("", LIGHT_TEXT_COLOR)],
        )
        style.configure(
            "Modern.Treeview", background="white", fieldbackground="white", rowheight=25
        )
        style.configure(
            "Modern.Treeview.Heading", font=("Segoe UI", 10, "bold"), padding=5
        )
        style.map("Modern.Treeview.Heading", background=[("active", "#E5E5E5")])
        style.configure("Error.TEntry", fieldbackground="#FFC0CB")

    def remover_alocacoes_do_funcionario(self, display_string):
        for lote_data in self.lote_widgets.values():
            for disc_widgets in lote_data["widgets"].values():
                for aloc_widget in disc_widgets["alocacoes_widgets"][:]:
                    if (
                        aloc_widget["frame"].winfo_exists()
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
            for disc_widgets in lote_data.get("widgets", {}).values():
                for aloc_widget in disc_widgets.get("alocacoes_widgets", []):
                    if aloc_widget.get("combo") and aloc_widget["combo"].winfo_exists():
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
            lote_tab_frame = ttk.Frame(self.notebook, padding=0)
            self.notebook.add(lote_tab_frame, text=f"Lote {lote_nome}")
            lote_notebook = ttk.Notebook(lote_tab_frame)
            lote_notebook.pack(fill=tk.BOTH, expand=True)
            entrada_tab = ttk.Frame(lote_notebook, padding=10)
            lote_notebook.add(entrada_tab, text="Entrada de Dados")
            dashboard_lote_tab = ttk.Frame(lote_notebook, padding=5)
            lote_notebook.add(dashboard_lote_tab, text="Dashboard do Lote")
            self.lote_widgets[lote_nome] = {
                "widgets": self.criar_conteudo_aba_lote(entrada_tab, lote_nome),
                "dash_tree": self._criar_treeview(dashboard_lote_tab),
            }

    def criar_conteudo_aba_lote(self, tab, lote_nome):
        canvas = tk.Canvas(tab, highlightthickness=0, bg="#FFFFFF")
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Modern.TFrame")
        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _on_mousewheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        widgets_disciplina = {}
        for disc in self.app_controller.get_disciplinas():
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
            widgets_disciplina[disc] = {
                "start_date": start_date_entry,
                "end_date": end_date_entry,
                "alocacoes_frame": aloc_frame,
                "alocacoes_widgets": [],
            }
            ttk.Button(
                disc_frame,
                text="+ Alocar Pessoa",
                command=lambda d=disc: self.adicionar_alocacao(lote_nome, d),
            ).pack(anchor=tk.W, pady=5)
        return widgets_disciplina

    def adicionar_alocacao(self, lote_nome, disciplina):
        frame = self.lote_widgets[lote_nome]["widgets"][disciplina]["alocacoes_frame"]
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
        action_buttons_frame = ttk.Frame(row_frame)
        action_buttons_frame.pack(side=tk.RIGHT)
        aloc_widget_dict = {
            "combo": combo,
            "entry": entry,
            "frame": row_frame,
            "monthly_hours_frame": monthly_hours_frame,
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
        self.lote_widgets[lote_nome]["widgets"][disciplina]["alocacoes_widgets"].append(
            aloc_widget_dict
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
        self.adicionar_alocacao(lote_nome, disciplina)
        nova_aloc_widget = self.lote_widgets[lote_nome]["widgets"][disciplina][
            "alocacoes_widgets"
        ][-1]
        horas_divididas = horas_originais / 2.0
        aloc_origem["entry"].delete(0, tk.END)
        aloc_origem["entry"].insert(0, f"{horas_divididas:.2f}".replace(".", ","))
        nova_aloc_widget["combo"].set(novo_display_string)
        nova_aloc_widget["entry"].delete(0, tk.END)
        nova_aloc_widget["entry"].insert(0, f"{horas_divididas:.2f}".replace(".", ","))

    def processar_calculo(self):
        for lote_nome, lote_data in self.lote_widgets.items():
            for disc_nome, disc_widgets in lote_data["widgets"].items():
                for aloc_widget in disc_widgets["alocacoes_widgets"]:
                    if aloc_widget["frame"].winfo_exists():
                        aloc_widget["entry"].config(style="TEntry")
                        horas_str = aloc_widget["entry"].get().strip().replace(",", ".")
                        if horas_str:
                            try:
                                float(horas_str)
                            except ValueError:
                                aloc_widget["entry"].config(style="Error.TEntry")
                                aloc_widget["entry"].focus_set()
                                messagebox.showerror(
                                    "Erro de Entrada",
                                    f"Valor de horas inválido: '{aloc_widget['entry'].get()}'\n\nLote: {lote_nome}, Disciplina: {disc_nome}",
                                )
                                return
        try:
            horas_mes = int(self.horas_mes_entry.get() or 160)
            lotes_data = []
            for lote_nome, lote_data in self.lote_widgets.items():
                cronograma, alocacao = {}, {}
                for disc, disc_widgets in lote_data["widgets"].items():
                    cronograma[disc] = {
                        "inicio": disc_widgets["start_date"].get(),
                        "fim": disc_widgets["end_date"].get(),
                    }
                    alocacao[disc] = []
                    for aloc_widget in disc_widgets["alocacoes_widgets"]:
                        if (
                            aloc_widget["frame"].winfo_exists()
                            and aloc_widget["combo"].get()
                        ):
                            horas_str = (
                                aloc_widget["entry"].get().strip().replace(",", ".")
                            )
                            nome, resto = aloc_widget["combo"].get().rsplit(" (", 1)
                            alocacao[disc].append(
                                {
                                    "funcionario": (nome, resto.rstrip(")")),
                                    "horas_totais": float(horas_str or 0.0),
                                }
                            )
                lotes_data.append(
                    {"nome": lote_nome, "cronograma": cronograma, "alocacao": alocacao}
                )
            self.app_controller.processar_portfolio(lotes_data, horas_mes)
        except Exception as e:
            messagebox.showerror(
                "Erro Inesperado", f"Ocorreu um erro no processamento: {e}"
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
                self._write_styled_df_to_excel(writer, "Consolidado", df_consolidado)
                if dashboards_lotes:
                    for nome_lote, df_lote in dashboards_lotes.items():
                        self._write_styled_df_to_excel(
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
        df_sorted = df_para_exportar.sort_values(
            by=["Cargo", "Funcionário", "Disciplina"]
        )
        try:
            del writer.book["Sheet"]
        except KeyError:
            pass
        ws = writer.book.create_sheet(title=sheet_name)
        writer.sheets = {ws.title: ws for ws in writer.book.worksheets}
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
        last_cargo = None
        for _, row_data in df_sorted.iterrows():
            if row_data["Cargo"] != last_cargo:
                last_cargo = row_data["Cargo"]
                ws.append([last_cargo.upper()])
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
                    cell.row > 1 and ws.cell(row=cell.row, column=2).value is None
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

    def _preencher_treeview(self, tree, df_relatorio):
        tree.delete(*tree.get_children())
        if df_relatorio.empty:
            return
        colunas_ui = [col for col in df_relatorio.columns if col != "Status"]
        tree["columns"] = colunas_ui
        tree["show"] = "headings"
        tree.tag_configure("excedido", background="#ffdddd", foreground="red")
        tree.tag_configure(
            "header", background="#e0e0e0", font=("Segoe UI", 10, "bold")
        )
        for col in colunas_ui:
            tree.heading(col, text=col)
            width = 200 if col == "Funcionário" else 120 if col == "Disciplina" else 100
            tree.column(col, width=width, anchor=tk.W)
        df_sorted = df_relatorio.sort_values(by=["Cargo", "Funcionário", "Disciplina"])
        last_cargo = None
        for _, row in df_sorted.iterrows():
            if row["Cargo"] != last_cargo:
                last_cargo = row["Cargo"]
                tree.insert(
                    "",
                    "end",
                    values=[last_cargo.upper()] + [""] * (len(colunas_ui) - 1),
                    tags=("header",),
                )
            tags = (
                ("excedido",)
                if "Status" in row and row["Status"] == "Alocação Excedida"
                else ()
            )
            valores_linha = [row[col] for col in colunas_ui]
            tree.insert("", "end", values=valores_linha, tags=tags)

    def atualizar_dashboards(self, df_consolidado, dashboards_lotes, detalhes_tarefas):
        self.notebook.select(0)
        self._preencher_treeview(self.dash_tree, df_consolidado)
        for nome_lote, df_lote in dashboards_lotes.items():
            if nome_lote in self.lote_widgets:
                self._preencher_treeview(
                    self.lote_widgets[nome_lote]["dash_tree"], df_lote
                )
        self._atualizar_horas_detalhadas(detalhes_tarefas)

    def _atualizar_horas_detalhadas(self, detalhes_tarefas):
        disciplinas = self.app_controller.get_disciplinas()
        for lote_nome, lote_data in self.lote_widgets.items():
            contador_tarefa_disc = {disc: 0 for disc in disciplinas}
            for disc, disc_widgets in lote_data.get("widgets", {}).items():
                for aloc_widget in disc_widgets["alocacoes_widgets"]:
                    if not aloc_widget["frame"].winfo_exists():
                        continue
                    for child in aloc_widget["monthly_hours_frame"].winfo_children():
                        child.destroy()
                    if aloc_widget["combo"].get():
                        try:
                            nome, resto = aloc_widget["combo"].get().rsplit(" (", 1)
                            cargo = resto.rstrip(")")
                            pessoa_tuplo = (nome, cargo)
                            tarefa_key = (
                                lote_nome,
                                disc,
                                pessoa_tuplo,
                                contador_tarefa_disc[disc],
                            )
                            contador_tarefa_disc[disc] += 1
                            if dados_mensais := detalhes_tarefas.get(tarefa_key):
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

# -*- coding: utf-8 -*-

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.colors import SchemeColor

# from openpyxl.drawing.fill import SolidColorFill
from openpyxl.chart.shapes import GraphicalProperties
import datetime
import openpyxl
from functools import partial
import json
import os

ESTADO_APP_FILE = "estado_app.json"


# MODIFICADO: A janela de diálogo agora inclui a seleção de disciplina.
class CustomDialog(tk.Toplevel):
    def __init__(self, parent, disciplinas):
        super().__init__(parent)
        self.transient(parent)
        self.title("Novo Funcionário")
        self.parent = parent
        self.result = None
        self.configure(bg="#F0F2F5")
        body = ttk.Frame(self, padding="10")
        self.initial_focus = self.body(body, disciplinas)
        body.pack(padx=10, pady=10)
        self.buttonbox()
        self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.geometry(f"+{parent.winfo_rootx()+50}+{parent.winfo_rooty()+50}")
        self.initial_focus.focus_set()
        self.wait_window(self)

    def body(self, master, disciplinas):
        ttk.Label(master, text="Nome:").grid(
            row=0, column=0, sticky=tk.W, padx=5, pady=5
        )
        self.nome_entry = ttk.Entry(master, width=30)
        self.nome_entry.grid(row=0, column=1, padx=5, pady=5)

        # O CAMPO DE CARGO FOI REMOVIDO DAQUI

        ttk.Label(master, text="Disciplina Principal:").grid(
            row=1, column=0, sticky=tk.W  # A linha agora é 1
        )
        self.disciplina_combo = ttk.Combobox(
            master, values=disciplinas, state="readonly", width=28
        )
        self.disciplina_combo.grid(row=1, column=1, padx=5, pady=5)  # A linha agora é 1
        if disciplinas:
            self.disciplina_combo.current(0)

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
        disciplina = self.disciplina_combo.get()

        if not nome or not disciplina:
            messagebox.showwarning(
                "Entrada Inválida",
                "Nome e disciplina são obrigatórios.",
                parent=self,
            )
            return

        self.result = (nome, disciplina)  # Retorna a tupla com 2 elementos
        self.withdraw()
        self.update_idletasks()
        self.cancel()

    def cancel(self, event=None):
        self.parent.focus_set()
        self.destroy()


class TeamManager(tk.Toplevel):
    # ... (código inalterado)
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
        # Chama o novo método para obter a lista dinâmica de TODAS as tarefas ativas
        lista_de_tarefas = self.app_controller.get_tarefas_ativas()
        dialog = CustomDialog(
            self,
            lista_de_tarefas,  # Passa a nova lista dinâmica para o diálogo
        )
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


class SelecionarTarefaDialog(simpledialog.Dialog):
    """Diálogo para selecionar uma tarefa (disciplina/subcontrato) de uma lista."""

    def __init__(self, parent, title, tarefas_disponiveis):
        self.tarefas = tarefas_disponiveis
        self.resultado = None
        super().__init__(parent, title)

    def body(self, master):
        ttk.Label(master, text="Selecione a tarefa para adicionar:").pack(pady=5)
        self.combo = ttk.Combobox(
            master, values=self.tarefas, state="readonly", width=30
        )
        self.combo.pack(padx=10, pady=5)
        if self.tarefas:
            self.combo.current(0)
        return self.combo

    def apply(self):
        self.resultado = self.combo.get()


class GuiApp(tk.Tk):
    # ... (init e outros métodos iniciais sem alteração)
    def __init__(self, app_controller):
        super().__init__()
        self.app_controller = app_controller
        self.title("Planejador de Recursos por Lotes")
        self.geometry("1600x900")
        self.configure(bg="#F0F2F5")
        self.setup_styles()
        self.team_window = None

        ### NOVO: Flag para evitar recursão na formatação de data ###
        self._is_formatting_date = False
        self._is_loading = False

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

        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _abrir_dialogo_adicionar_subcontrato(self, lote_nome, parent_frame):
        """Abre um diálogo para o usuário selecionar e adicionar um novo SUBCONTRATO ao lote."""
        # Pega a lista de TODOS os subcontratos possíveis
        todos_subcontratos = self.app_controller.get_subcontratos_disponiveis()

        # Pega a lista de TODAS as tarefas que já estão no lote (sejam disciplinas ou subcontratos)
        tarefas_ja_adicionadas = self.app_controller.get_tarefas_ja_adicionadas(
            lote_nome
        )

        # Filtra a lista de subcontratos, mostrando apenas os que ainda não foram adicionados
        subcontratos_disponiveis = [
            s for s in todos_subcontratos if s not in tarefas_ja_adicionadas
        ]

        if not subcontratos_disponiveis:
            messagebox.showinfo(
                "Sem Subcontratos",
                "Todos os subcontratos disponíveis já foram adicionados a este lote.",
                parent=self,
            )
            return

        dialog = SelecionarTarefaDialog(
            self, "Adicionar Subcontrato", subcontratos_disponiveis
        )
        novo_subcontrato = dialog.resultado

        if novo_subcontrato:
            # A função _criar_frame_tarefa é genérica e funciona perfeitamente aqui
            self._criar_frame_tarefa(
                parent_frame, lote_nome, novo_subcontrato, popular_defaults=True
            )

    def _on_data_keyrelease(self, event):
        """
        Chamado quando uma tecla é solta no campo de data.
        Formata o texto para DD/MM/AAAA e reposiciona o cursor corretamente.
        """
        entry = event.widget

        # Pega a posição do cursor e o texto ANTES da formatação
        cursor_pos = entry.index(tk.INSERT)
        texto_antes = entry.get()

        # Limpa o texto, mantendo apenas os dígitos
        digitos = "".join(filter(str.isdigit, texto_antes))
        digitos = digitos[:8]  # Limita a 8 dígitos (DDMMAAAA)

        # Formata o texto com as barras
        texto_formatado = ""
        if len(digitos) > 4:
            texto_formatado = f"{digitos[:2]}/{digitos[2:4]}/{digitos[4:]}"
        elif len(digitos) > 2:
            texto_formatado = f"{digitos[:2]}/{digitos[2:]}"
        else:
            texto_formatado = digitos

        # Calcula o ajuste do cursor
        # Se uma barra foi adicionada ANTES da posição original do cursor,
        # o cursor precisa ser movido para a direita.
        ajuste_cursor = 0
        if len(texto_formatado) > len(texto_antes):
            # Verifica se uma barra foi inserida nos primeiros 2 caracteres
            if (
                cursor_pos >= 2
                and texto_antes.count("/") == 0
                and texto_formatado.count("/") >= 1
            ):
                ajuste_cursor = 1
            # Verifica se uma barra foi inserida nos primeiros 5 caracteres
            if (
                cursor_pos >= 5
                and texto_antes.count("/") == 1
                and texto_formatado.count("/") >= 2
            ):
                ajuste_cursor = 1

        nova_pos_cursor = cursor_pos + ajuste_cursor

        # Atualiza o texto no Entry e reposiciona o cursor
        entry.delete(0, tk.END)
        entry.insert(0, texto_formatado)
        entry.icursor(nova_pos_cursor)

        return "break"  # Impede que outros bindings de tecla sejam executados

    def _on_combobox_click(self, event, aloc_widget):
        """Chamado quando uma combobox de alocação é clicada."""
        combobox_alvo = event.widget
        # Encontra o lote ao qual a combobox pertence
        for lote_nome, lote_data in self.lote_widgets.items():
            if "disciplinas" in lote_data:
                for disc_data in lote_data["disciplinas"].values():
                    if aloc_widget in disc_data["alocacoes_widgets"]:
                        self._atualizar_opcoes_funcionario(
                            lote_nome, combobox_alvo, aloc_widget
                        )
                        return

    def _atualizar_opcoes_funcionario(
        self, lote_nome, combobox_alvo, aloc_widget_clicado
    ):
        """
        Atualiza a lista de valores de uma combobox, removendo nomes de funcionários
        que já foram selecionados em outras comboboxes do mesmo lote.
        """
        nomes_usados = set()
        lote_data = self.lote_widgets[lote_nome]
        # 1. Coleta todos os nomes já usados no lote
        if "disciplinas" in lote_data:
            for disc_data in lote_data["disciplinas"].values():
                for aloc_widget in disc_data["alocacoes_widgets"]:
                    combo_atual = aloc_widget.get("combo")
                    if combo_atual is not combobox_alvo and combo_atual.winfo_exists():
                        selecao = combo_atual.get()
                        if selecao:
                            nomes_usados.add(selecao)

        # 2. Obtém a lista de nomes base para a disciplina
        disciplina_alvo = aloc_widget_clicado.get("disciplina")
        nomes_base = self.app_controller.get_nomes_funcionarios_por_disciplina(
            disciplina_alvo
        )

        # 3. Filtra os nomes disponíveis
        opcoes_disponiveis = [nome for nome in nomes_base if nome not in nomes_usados]

        # 4. Garante que a seleção atual (se houver) permaneça na lista
        selecao_propria = combobox_alvo.get()
        if selecao_propria and selecao_propria not in opcoes_disponiveis:
            opcoes_disponiveis.insert(0, selecao_propria)

        # 5. Atualiza os valores da combobox
        combobox_alvo.config(values=sorted(opcoes_disponiveis))

    def limpar_alocacoes_de_funcionario_removido(self, display_string):
        """Limpa as seleções da combobox que continham o funcionário removido."""
        try:
            # Extrai apenas o nome da string completa "Nome (Cargo) [Disciplina]"
            nome_removido = display_string.split(" (")[0]
            for lote_data in self.lote_widgets.values():
                if "disciplinas" in lote_data:
                    for disc_widgets in lote_data["disciplinas"].values():
                        for aloc_widget in disc_widgets["alocacoes_widgets"]:
                            if (
                                aloc_widget["combo"].winfo_exists()
                                and aloc_widget["combo"].get() == nome_removido
                            ):
                                aloc_widget["combo"].set("")  # Limpa a seleção
            self.processar_calculo()  # Recalcula para atualizar os dashboards
        except Exception as e:
            print(f"Erro ao limpar alocação de funcionário removido: {e}")

    def _on_horas_keyrelease(self, event):
        entry = event.widget
        texto_atual = entry.get()
        texto_limpo = ""
        virgula_encontrada = False
        for char in texto_atual:
            if char.isdigit():
                texto_limpo += char
            elif char in (",", ".") and not virgula_encontrada:
                texto_limpo += ","
                virgula_encontrada = True

        if texto_limpo != texto_atual:
            cursor_pos = entry.index(tk.INSERT)
            entry.delete(0, tk.END)
            entry.insert(0, texto_limpo)
            if cursor_pos > 0:
                entry.icursor(cursor_pos - 1)

    def _on_closing(self):
        # Se o MAC não for autorizado, apenas fecha a janela sem salvar
        if not self.app_controller.mac_autorizado:
            self.destroy()
            return

        # Se for autorizado, salva o estado e fecha
        self.salvar_estado()
        self.destroy()

    def create_control_panel_widgets(self, parent):
        # Frame da Equipe
        team_frame = ttk.LabelFrame(parent, text="Equipe", padding=10)
        team_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))
        team_frame.columnconfigure(0, weight=1)

        self.btn_team = ttk.Button(
            team_frame,
            text="Gerenciar Equipe",
            command=self.open_team_manager,
        )
        self.btn_team.grid(row=0, column=0, sticky="ew", ipady=4)

        # Frame de Configuração
        setup_frame = ttk.LabelFrame(parent, text="Configuração", padding=10)
        setup_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

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
        self.btn_lotes = ttk.Button(
            setup_frame,
            text="Gerar Abas de Lotes",
            command=self.gerar_abas_lotes,
        )

        self.btn_lotes.pack(fill=tk.X)

        # Frame de Ações
        action_frame = ttk.LabelFrame(parent, text="Ações", padding=10)
        action_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)
        action_frame.columnconfigure(0, weight=1)

        self.btn_calcular = ttk.Button(
            action_frame,
            text="Calcular Alocação",
            command=self.processar_calculo,
            style="Accent.TButton",
        )
        self.btn_calcular.grid(row=0, column=0, sticky="ew", ipady=6, pady=2)

        self.btn_exportar_cronograma = ttk.Button(
            action_frame,
            text="Exportar Relatório de Cronograma",
            command=self.exportar_para_excel,
            style="Accent.TButton",
        )
        self.btn_exportar_cronograma.grid(row=1, column=0, sticky="ew", ipady=6, pady=2)

        self.btn_exportar_quantidades = ttk.Button(
            action_frame,
            text="Exportar Planilha de Quantidades",
            command=self.exportar_resumo_equipe,
            style="Accent.TButton",
        )
        self.btn_exportar_quantidades.grid(
            row=2, column=0, sticky="ew", ipady=6, pady=2
        )

    def set_modo_restringido(self, restringido):
        """Habilita ou desabilita as funcionalidades principais da aplicação."""
        if restringido:
            novo_estado = tk.DISABLED
            # Mostra um aviso claro para o usuário
            messagebox.showerror(
                "Licença Inválida",
                "Este aplicativo não está licenciado para esta máquina.\nAs funcionalidades de cálculo, exportação e salvamento foram desabilitadas.",
            )
        else:
            novo_estado = tk.NORMAL

        # Altera o estado dos botões
        if hasattr(self, "btn_calcular"):  # Verifica se os botões já existem
            self.btn_team.config(state=novo_estado)
            self.btn_lotes.config(state=novo_estado)
            self.btn_calcular.config(state=novo_estado)
            self.btn_exportar_cronograma.config(state=novo_estado)
            self.btn_exportar_quantidades.config(state=novo_estado)

    # NOVO: Função para exportar o resumo da equipe.
    def exportar_resumo_equipe(self):
        try:
            df_resumo = self.app_controller.get_resumo_equipe_df()

            if df_resumo.empty:
                messagebox.showwarning(
                    "Aviso",
                    "Não há dados de alocação da equipe para exportar.",
                    parent=self,
                )
                return

            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Arquivo Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")],
                title="Salvar Resumo da Equipe",
                initialfile=f"Resumo_Equipe_{datetime.datetime.now().strftime('%Y-%m-%d')}.xlsx",
            )
            if not filepath:
                return

            with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
                cargos = self.app_controller.get_cargos_disponiveis()
                self._write_styled_resumo_to_excel(
                    writer, "Resumo por Funcionário", df_resumo, cargos
                )

                # Remove a planilha padrão APENAS DEPOIS de criar a nova
                if "Sheet" in writer.book.sheetnames:
                    writer.book.remove(writer.book["Sheet"])

            messagebox.showinfo(
                "Sucesso",
                f"Resumo da equipe exportado com sucesso para:\n{filepath}",
                parent=self,
            )

        except Exception as e:
            messagebox.showerror(
                "Erro na Exportação",
                f"Ocorreu um erro ao exportar o resumo da equipe:\n{e}",
                parent=self,
            )

    # NOVO: Método para escrever o resumo da equipe com formatação avançada.
    def _write_styled_resumo_to_excel(self, writer, sheet_name, df, cargos):
        book = writer.book
        ws = book.create_sheet(title=sheet_name)
        writer.sheets[sheet_name] = ws
        ws.sheet_view.showGridLines = False

        # --- ESTILOS PADRONIZADOS ---
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(
            start_color="44546A", end_color="44546A", fill_type="solid"
        )
        category_font = Font(name="Arial", bold=True, size=12)
        category_fill = PatternFill(
            start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"
        )
        data_font = Font(name="Arial", size=10)
        thin_side = Side(border_style="thin", color="BFBFBF")
        full_border = Border(
            left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
        )
        horas_fill = PatternFill(
            start_color="FFFFE0", end_color="FFFFE0", fill_type="solid"
        )  # Amarelo claro

        # --- 1. TABELA DE PREÇOS ---
        ws.cell(row=2, column=2, value="Tabela de Preços por Cargo").font = Font(
            bold=True, size=14
        )
        price_header_font = Font(name="Arial", bold=True, size=11)
        cell_cargo_header = ws.cell(row=3, column=2, value="Cargo")
        cell_cargo_header.font = price_header_font
        cell_cargo_header.border = full_border
        cell_preco_header = ws.cell(row=3, column=3, value="Preço/Hora (R$)")
        cell_preco_header.font = price_header_font
        cell_preco_header.border = full_border  # Adiciona borda
        # ws.cell(row=3, column=2, value="Cargo").font = price_header_font
        # ws.cell(row=3, column=3, value="Preço/Hora (R$)").font = price_header_font

        input_style = NamedStyle(name="input_style_resumo")
        input_style.fill = PatternFill(
            start_color="FFFFCC", end_color="FFFFCC", fill_type="solid"
        )
        input_style.number_format = '"R$" #,##0.00'
        input_style.font = data_font
        if input_style.name not in book.style_names:
            book.add_named_style(input_style)

        start_row_price_table = 4
        for i, cargo_nome in enumerate(cargos, start=start_row_price_table):
            ws.cell(row=i, column=2, value=cargo_nome).font = data_font
            price_cell = ws.cell(row=i, column=3)
            price_cell.style = input_style

        end_row_price_table = start_row_price_table + len(cargos) - 1
        price_table_range_str = f"'{sheet_name}'!${openpyxl.utils.get_column_letter(2)}${start_row_price_table}:${openpyxl.utils.get_column_letter(3)}${end_row_price_table}"

        named_range = openpyxl.workbook.defined_name.DefinedName(
            "TabelaPrecos", attr_text=price_table_range_str
        )

        # CORRECTED FIX: Assign the named range like a dictionary entry
        book.defined_names["TabelaPrecos"] = named_range

        # --- 2. TABELA PRINCIPAL DE DADOS ---
        df["Cargo"] = pd.Categorical(df["Cargo"], categories=cargos, ordered=True)
        df_sorted = df.sort_values(by=["Disciplina", "Cargo", "Funcionário"])
        start_row_main_table = end_row_price_table + 3

        headers = [
            "Funcionário",
            "Cargo",
            "Disciplina",
            "Horas Totais Alocadas",
            "Valor Total (R$)",
        ]
        for c_idx, header_text in enumerate(headers, 1):
            cell = ws.cell(row=start_row_main_table, column=c_idx, value=header_text)
            cell.font = header_font
            cell.fill = header_fill

        rows = dataframe_to_rows(df_sorted, index=False, header=False)
        current_row_excel = start_row_main_table + 1
        last_discipline = None

        for row_data_tuple in rows:
            disciplina_atual = row_data_tuple[2]
            if disciplina_atual != last_discipline:
                cat_cell = ws.cell(
                    row=current_row_excel, column=1, value=disciplina_atual
                )
                cat_cell.font = category_font
                cat_cell.fill = category_fill
                ws.merge_cells(
                    start_row=current_row_excel,
                    start_column=1,
                    end_row=current_row_excel,
                    end_column=len(headers),
                )
                last_discipline = disciplina_atual
                current_row_excel += 1

            for c_idx, value in enumerate(row_data_tuple, 1):
                cell = ws.cell(row=current_row_excel, column=c_idx, value=value)
                if c_idx == 4:  # Coluna 4 é "Horas Totais Alocadas"
                    cell.fill = horas_fill
                    cell.font = data_font

            cargo_cell_coord = (
                f"{openpyxl.utils.get_column_letter(2)}{current_row_excel}"
            )
            horas_cell_coord = (
                f"{openpyxl.utils.get_column_letter(4)}{current_row_excel}"
            )
            formula = f'=IFERROR(VLOOKUP({cargo_cell_coord},TabelaPrecos,2,FALSE)*{horas_cell_coord},"-")'

            valor_cell = ws.cell(row=current_row_excel, column=5, value=formula)
            valor_cell.number_format = '"R$" #,##0.00'
            current_row_excel += 1
        end_row_main_table = current_row_excel - 1

        # --- 3. SEÇÃO DE TOTAIS E RESUMO FINANCEIRO ---
        total_geral_row = end_row_main_table + 1
        valor_col_letter = openpyxl.utils.get_column_letter(5)
        total_geral_label_cell = ws.cell(
            row=total_geral_row, column=4, value="Total Geral"
        )
        total_geral_label_cell.font = Font(name="Arial", bold=True, size=11)
        total_geral_label_cell.alignment = Alignment(horizontal="right")

        total_geral_valor_cell = ws.cell(
            row=total_geral_row,
            column=5,
            value=f"=SUM({valor_col_letter}{start_row_main_table+1}:{valor_col_letter}{end_row_main_table})",
        )
        total_geral_valor_cell.font = Font(name="Arial", bold=True, size=11)
        total_geral_valor_cell.number_format = '"R$" #,##0.00'

        start_row_summary = total_geral_row + 3
        ws.cell(
            row=start_row_summary, column=2, value="Resumo Financeiro por Disciplina"
        ).font = Font(bold=True, size=14)
        summary_headers = [
            "Disciplina",
            "Total HH (R$)",
            "Total por Desenho (R$)",
            "Subtotal (R$)",
        ]
        for c_idx, text in enumerate(summary_headers, 2):
            cell = ws.cell(row=start_row_summary + 1, column=c_idx, value=text)
            cell.font = header_font
            cell.fill = header_fill

        disciplina_col_main_table = openpyxl.utils.get_column_letter(3)
        valor_col_main_table = openpyxl.utils.get_column_letter(5)
        disciplinas_unicas = df["Disciplina"].unique()

        start_row_summary_data = start_row_summary + 2
        for i, disciplina_nome in enumerate(disciplinas_unicas):
            current_summary_row = start_row_summary_data + i

            ws.cell(row=current_summary_row, column=2, value=disciplina_nome)

            criteria_cell = ws.cell(row=current_summary_row, column=2).coordinate
            formula_sumif = f"=SUMIF({disciplina_col_main_table}${start_row_main_table+1}:{disciplina_col_main_table}${end_row_main_table}, {criteria_cell}, {valor_col_main_table}${start_row_main_table+1}:{valor_col_main_table}${end_row_main_table})"
            total_cell = ws.cell(row=current_summary_row, column=3, value=formula_sumif)
            total_cell.number_format = '"R$" #,##0.00'

            desconto_cell = ws.cell(row=current_summary_row, column=4)
            desconto_cell.style = input_style

            formula_subtotal = f"={ws.cell(row=current_summary_row, column=3).coordinate}-{ws.cell(row=current_summary_row, column=4).coordinate}"
            subtotal_cell = ws.cell(
                row=current_summary_row, column=5, value=formula_subtotal
            )
            subtotal_cell.number_format = '"R$" #,##0.00'

        # --- CORREÇÃO DA POSIÇÃO DO TOTAL FINAL ---
        last_data_row_summary = start_row_summary_data + len(disciplinas_unicas) - 1
        total_final_row = last_data_row_summary + 1

        total_final_label = ws.cell(row=total_final_row, column=4, value="Total Final")
        total_final_label.font = Font(name="Arial", bold=True, size=12)
        total_final_label.alignment = Alignment(horizontal="right")

        subtotal_col_letter = openpyxl.utils.get_column_letter(5)
        # CORREÇÃO DA FÓRMULA DE SOMA
        formula_total_final = f"=SUM({subtotal_col_letter}{start_row_summary_data}:{subtotal_col_letter}{last_data_row_summary})"
        total_final_valor = ws.cell(
            row=total_final_row, column=5, value=formula_total_final
        )
        total_final_valor.font = Font(name="Arial", bold=True, size=12)
        total_final_valor.number_format = '"R$" #,##0.00'

        # --- 4. FORMATAÇÃO FINAL ---
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                if cell.value is not None or cell.style == "input_style_resumo":
                    cell.border = full_border

        for i in range(1, len(headers) + 2):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 25

    # ... (O resto da classe GuiApp permanece inalterado)

    def open_team_manager(self):
        if self.team_window and self.team_window.winfo_exists():
            self.team_window.focus()
        else:
            self.team_window = TeamManager(self, self.app_controller)

    # ... (setup_styles inalterado)
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
        # ... (código inalterado)
        for lote_data in self.lote_widgets.values():
            if "disciplinas" in lote_data:
                for disc_widgets in lote_data["disciplinas"].values():
                    for aloc_widget in disc_widgets["alocacoes_widgets"][:]:
                        if (
                            aloc_widget["frame"].winfo_exists()
                            and aloc_widget.get("type") == "equipe"
                            # A comparação agora deve reconstruir a string completa
                            and f"{aloc_widget['combo'].get()} [{disc_widgets['frame'].cget('text')}]"
                            == display_string
                        ):
                            aloc_widget["frame"].destroy()
                            disc_widgets["alocacoes_widgets"].remove(aloc_widget)
        self.processar_calculo()

    def _criar_treeview(self, parent_frame):
        # ... (código inalterado)
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

    # MODIFICADO: Atualiza todas as listas de forma mais inteligente.
    def atualizar_lista_funcionarios(self):
        # ... (código inalterado, mas verificado para compatibilidade) ...
        if self.team_window and self.team_window.winfo_exists():
            self.team_window.populate_listbox()
        for lote_data in self.lote_widgets.values():
            if "disciplinas" in lote_data:
                for disc_nome, disc_widgets in lote_data["disciplinas"].items():
                    funcionarios_filtrados = (
                        self.app_controller.get_nomes_funcionarios_por_disciplina(
                            disc_nome
                        )
                    )
                    for aloc_widget in disc_widgets.get("alocacoes_widgets", []):
                        if (
                            aloc_widget.get("combo_nome")
                            and aloc_widget["combo_nome"].winfo_exists()
                        ):
                            selecao_atual = aloc_widget["combo_nome"].get()
                            aloc_widget["combo_nome"].config(
                                values=[""] + funcionarios_filtrados
                            )
                            if selecao_atual in funcionarios_filtrados:
                                aloc_widget["combo_nome"].set(selecao_atual)
                            else:
                                aloc_widget["combo_nome"].set("")

    def _atualizar_opcoes_funcionarios(
        self, lote_nome, disciplina_alvo, combobox_clicada
    ):
        """
        Filtra a lista de um Combobox para mostrar apenas funcionários não utilizados no lote.
        """
        nomes_ja_selecionados = set()
        lote_data = self.lote_widgets.get(lote_nome, {})

        # 1. Coleta todos os nomes já selecionados em OUTROS comboboxes do mesmo lote
        if "disciplinas" in lote_data:
            for disc_widgets in lote_data["disciplinas"].values():
                for aloc_widget in disc_widgets.get("alocacoes_widgets", []):
                    combo_atual = aloc_widget.get("combo_nome")

                    # Verifica se o widget ainda existe e se NÃO é o que foi clicado
                    if (
                        combo_atual
                        and combo_atual.winfo_exists()
                        and combo_atual is not combobox_clicada
                    ):
                        selecao = combo_atual.get()
                        if selecao:  # Garante que não adicione strings vazias
                            nomes_ja_selecionados.add(selecao)

        # 2. Pega a lista base de funcionários para a disciplina
        nomes_base = self.app_controller.get_nomes_funcionarios_por_disciplina(
            disciplina_alvo
        )

        # 3. Filtra a lista, removendo os nomes já selecionados
        opcoes_disponiveis = [
            nome for nome in nomes_base if nome not in nomes_ja_selecionados
        ]

        # 4. Garante que a seleção atual do combobox clicado esteja na lista (o "edge case")
        selecao_atual = combobox_clicada.get()
        if selecao_atual and selecao_atual not in opcoes_disponiveis:
            opcoes_disponiveis.insert(0, selecao_atual)

        # 5. Atualiza os valores do combobox, sempre incluindo a opção vazia
        combobox_clicada["values"] = [""] + sorted(opcoes_disponiveis)

    # ### INÍCIO DA MODIFICAÇÃO: Helper para remover linha de alocação ###
    def _remover_linha_alocacao(self, lote_nome, disc_nome, widget_dict_a_remover):
        """Remove uma linha de alocação. As horas dos outros não são recalculadas."""
        if widget_dict_a_remover["frame"].winfo_exists():
            widget_dict_a_remover["frame"].destroy()
        try:
            lista_widgets = self.lote_widgets[lote_nome]["disciplinas"][disc_nome][
                "alocacoes_widgets"
            ]
            if widget_dict_a_remover in lista_widgets:
                lista_widgets.remove(widget_dict_a_remover)
        except (KeyError, ValueError):
            pass

    # ### FIM DA MODIFICAÇÃO ###

    # MODIFICADO: A combobox de alocação agora usa a lista filtrada.
    def adicionar_alocacao_equipe(
        self, lote_nome, disciplina, cargo_a_definir=None, horas_a_definir=None
    ):
        frame = self.lote_widgets[lote_nome]["disciplinas"][disciplina][
            "alocacoes_frame"
        ]
        row_frame = ttk.Frame(frame)
        row_frame.pack(fill=tk.X, pady=2)
        ttk.Label(row_frame, text="Funcionário:").pack(side=tk.LEFT, padx=(0, 5))
        nomes_funcionarios = self.app_controller.get_funcionarios_para_tarefa(
            disciplina
        )
        combo_nome = ttk.Combobox(
            row_frame,
            values=[""] + sorted(nomes_funcionarios),
            state="readonly",
            width=25,
        )
        combo_nome.pack(side=tk.LEFT, padx=(0, 10))
        combo_nome.set("")

        # NOVO: Associa o clique do mouse à função de atualização da lista
        ttk.Label(row_frame, text="Cargo:").pack(side=tk.LEFT, padx=(0, 5))
        cargos_disponiveis = self.app_controller.get_cargos_disponiveis()
        combo_cargo = ttk.Combobox(
            row_frame, values=cargos_disponiveis, state="readonly", width=20
        )
        combo_cargo.pack(side=tk.LEFT, padx=(0, 10))
        if cargo_a_definir:
            combo_cargo.set(cargo_a_definir)

        ttk.Label(row_frame, text="Horas Totais:").pack(side=tk.LEFT, padx=(0, 5))
        entry = ttk.Entry(row_frame, width=10)
        entry.pack(side=tk.LEFT, padx=5)
        entry.bind("<KeyRelease>", self._on_horas_keyrelease)
        if horas_a_definir is not None:
            entry.insert(0, f"{horas_a_definir:.2f}".replace(".", ","))

        action_buttons_frame = ttk.Frame(row_frame)
        action_buttons_frame.pack(side=tk.RIGHT, padx=(10, 0))

        aloc_widget_dict = {
            "type": "equipe",
            "combo_nome": combo_nome,
            "combo_cargo": combo_cargo,
            "entry": entry,
            "frame": row_frame,
            "disciplina": disciplina,
        }

        dividir_cmd = partial(
            self.dividir_tarefa, lote_nome, disciplina, aloc_widget_dict
        )
        dividir_btn = ttk.Button(
            action_buttons_frame, text="Dividir", width=7, command=dividir_cmd
        )
        dividir_btn.pack(side=tk.LEFT, padx=(0, 8))

        remove_cmd = partial(
            self._remover_linha_alocacao, lote_nome, disciplina, aloc_widget_dict
        )
        remove_btn = ttk.Button(
            action_buttons_frame, text="X", width=3, command=remove_cmd
        )
        remove_btn.pack(side=tk.LEFT, padx=(0, 0))

        self.lote_widgets[lote_nome]["disciplinas"][disciplina][
            "alocacoes_widgets"
        ].append(aloc_widget_dict)

    def _criar_botao_adicionar_subcontrato(self, parent_frame, lote_nome):
        """Cria e posiciona o botão '+ Adicionar Subcontrato' no final do frame pai."""
        add_button_frame = ttk.Frame(parent_frame)
        add_button_frame.pack(fill=tk.X, padx=10, pady=20)

        cmd = partial(
            self._abrir_dialogo_adicionar_subcontrato, lote_nome, parent_frame
        )
        ttk.Button(
            add_button_frame,
            text="+ Adicionar Subcontrato",
            command=cmd,
            style="Accent.TButton",
        ).pack()

    def _calcular_horas_periodo(self, start_str, end_str):
        """Calcula o total de horas para um período e retorna como float, ou None se falhar."""
        try:
            horas_mes_str = self.horas_mes_entry.get()

            if not all([start_str, end_str, horas_mes_str]):
                return None

            inicio = datetime.datetime.strptime(start_str, "%d/%m/%Y")
            fim = datetime.datetime.strptime(end_str, "%d/%m/%Y")
            horas_mes = int(horas_mes_str)

            if fim < inicio:
                return None

            num_meses = (fim.year - inicio.year) * 12 + (fim.month - inicio.month) + 1
            total_horas_periodo = num_meses * horas_mes
            return float(total_horas_periodo)

        except (ValueError, TypeError):
            return None

    def _on_adicionar_alocacao_click(self, lote_nome, disc_nome):
        """
        Handler para o clique do botão '+ Alocar Equipe'.
        Pega as datas, calcula as horas e chama a função para criar a nova linha
        já com o valor das horas preenchido.
        """
        # Acessa os widgets da disciplina correta para pegar as datas
        disc_widgets = self.lote_widgets[lote_nome]["disciplinas"][disc_nome]
        start_str = disc_widgets["start_date"].get()
        end_str = disc_widgets["end_date"].get()

        # Usa a nova função para calcular as horas
        horas_calculadas = self._calcular_horas_periodo(start_str, end_str)

        # Chama a função que realmente cria os widgets na tela,
        # passando o valor calculado para o parâmetro 'horas_a_definir'.
        self.adicionar_alocacao_equipe(
            lote_nome, disc_nome, horas_a_definir=horas_calculadas
        )

    def _atualizar_horas_alocadas_por_disciplina(
        self, lote_nome, tarefa_nome, event=None
    ):
        if self._is_loading:
            return

        try:
            disc_widgets = self.lote_widgets[lote_nome]["disciplinas"][tarefa_nome]

            start_str = disc_widgets["start_date"].get()
            end_str = disc_widgets["end_date"].get()
            horas_mes_str = self.horas_mes_entry.get()

            if not all([start_str, end_str, horas_mes_str]):
                return  # Sai silenciosamente se as datas não estiverem prontas

            inicio = datetime.datetime.strptime(start_str, "%d/%m/%Y")
            fim = datetime.datetime.strptime(end_str, "%d/%m/%Y")
            horas_mes = int(horas_mes_str)

            if fim < inicio:
                return

            num_meses = (fim.year - inicio.year) * 12 + (fim.month - inicio.month) + 1
            total_horas_periodo = num_meses * horas_mes
            horas_formatadas = f"{total_horas_periodo:.2f}".replace(".", ",")

            # Itera sobre todas as alocações da tarefa
            for aloc in disc_widgets["alocacoes_widgets"]:
                if aloc["frame"].winfo_exists():
                    entry = aloc["entry"]
                    valor_atual_str = entry.get().strip()

                    # --- INÍCIO DA CORREÇÃO LÓGICA ---
                    deve_atualizar = False
                    if not valor_atual_str:
                        # Condição 1: O campo está literalmente vazio.
                        deve_atualizar = True
                    else:
                        try:
                            # Condição 2: O campo contém um número que é igual a zero.
                            valor_numerico = float(valor_atual_str.replace(",", "."))
                            if valor_numerico == 0:
                                deve_atualizar = True
                        except ValueError:
                            # Se o campo tiver um texto não numérico, não faz nada.
                            pass

                    if deve_atualizar:
                        entry.delete(0, tk.END)
                        entry.insert(0, horas_formatadas)
                    # --- FIM DA CORREÇÃO LÓGICA ---

        except (ValueError, TypeError, KeyError):
            # Captura qualquer erro de conversão de data, tipo ou busca de chave e ignora.
            pass

            # Em gui_app.py, na classe GuiApp

    def _recalcular_todas_as_horas_da_tarefa(self, lote_nome, tarefa_nome, event=None):
        """
        Calcula as horas do período e SOBRESCREVE o valor em todas as alocações da tarefa.
        Esta função é acionada pela mudança das datas.
        """
        if self._is_loading:
            return

        try:
            disc_widgets = self.lote_widgets[lote_nome]["disciplinas"][tarefa_nome]
            start_str = disc_widgets["start_date"].get()
            end_str = disc_widgets["end_date"].get()

            # Usa a função auxiliar que já temos para pegar o valor numérico das horas
            total_horas = self._calcular_horas_periodo(start_str, end_str)

            if total_horas is None:
                # Se o cálculo falhar (datas inválidas, etc.), não faz nada
                return

            horas_formatadas = f"{total_horas:.2f}".replace(".", ",")

            # Loop que sobrescreve TUDO
            for aloc in disc_widgets["alocacoes_widgets"]:
                if aloc["frame"].winfo_exists():
                    entry = aloc["entry"]
                    entry.delete(0, tk.END)
                    entry.insert(0, horas_formatadas)

        except (KeyError, ValueError):
            # Ignora erros se os widgets não forem encontrados ou as datas forem inválidas
            pass

    def _coletar_dados_da_ui(self):
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
                lote_dict = {"nome": lote_nome, "disciplinas": {}}

                # A chave "disciplinas" agora contém todas as tarefas (disciplinas e subcontratos)
                if "disciplinas" in lote_data:
                    for tarefa_nome, tarefa_widgets in lote_data["disciplinas"].items():
                        tarefa_dict = {
                            "cronograma": {
                                "inicio": tarefa_widgets["start_date"].get(),
                                "fim": tarefa_widgets["end_date"].get(),
                            },
                            "alocacoes": [],
                        }

                        for aloc_widget in tarefa_widgets["alocacoes_widgets"]:
                            if aloc_widget["frame"].winfo_exists():
                                horas_str = (
                                    aloc_widget["entry"].get().strip().replace(",", ".")
                                )
                                cargo_func = aloc_widget["combo_cargo"].get()
                                nome_func = aloc_widget["combo_nome"].get().strip()

                                # Mantém a lógica de salvamento
                                nome_a_salvar = nome_func or cargo_func
                                if nome_a_salvar:
                                    try:
                                        horas_float = float(horas_str or 0.0)
                                    except ValueError:
                                        horas_float = 0.0

                                    tarefa_dict["alocacoes"].append(
                                        {
                                            "funcionario": (nome_a_salvar, cargo_func),
                                            "horas_totais": horas_float,
                                        }
                                    )
                        lote_dict["disciplinas"][tarefa_nome] = tarefa_dict
                estado_geral["lotes"].append(lote_dict)
            return estado_geral
        except Exception as e:
            messagebox.showerror(
                "Erro de Coleta", f"Ocorreu um erro ao ler os dados para salvar: {e}"
            )
            return None

    # MODIFICADO: Coleta dados considerando a nova estrutura de funcionário.
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

    # MODIFICADO: Carrega o estado e recria as alocações corretamente.

    def carregar_ou_iniciar_ui(self):
        if not os.path.exists(ESTADO_APP_FILE):
            print(
                "Nenhum estado salvo encontrado. Aguardando geração de lotes pelo usuário."
            )
            return

        self._is_loading = True
        try:
            with open(ESTADO_APP_FILE, "r", encoding="utf-8") as f:
                estado = json.load(f)

            # Carrega configs e funcionários
            config = estado.get("config", {})
            self.horas_mes_entry.delete(0, tk.END)
            self.horas_mes_entry.insert(0, config.get("horas_mes", 160))
            self.num_lotes_entry.delete(0, tk.END)
            self.num_lotes_entry.insert(0, config.get("num_lotes", 1))

            for func_data in estado.get("funcionarios", []):
                if isinstance(func_data, list) and len(func_data) == 2:
                    self.app_controller.adicionar_funcionario(*func_data)

            # Gera as abas de lotes, mas SEM popular o conteúdo ainda
            self.gerar_abas_lotes(popular_defaults=False)

            # Reconstrói o estado de CADA lote
            for lote_data in estado.get("lotes", []):
                lote_nome = lote_data.get("nome")
                if lote_nome in self.lote_widgets:
                    # CORREÇÃO: Pega o scrollable_frame que foi salvo no dicionário de widgets
                    parent_frame = self.lote_widgets[lote_nome].get("scrollable_frame")
                    if not parent_frame:
                        continue

                    # Itera sobre TODAS as tarefas salvas para recriá-las
                    for tarefa_nome, tarefa_data in lote_data.get(
                        "disciplinas", {}
                    ).items():
                        self._criar_frame_tarefa(
                            parent_frame, lote_nome, tarefa_nome, popular_defaults=False
                        )

                        tarefa_widgets = self.lote_widgets[lote_nome]["disciplinas"][
                            tarefa_nome
                        ]
                        cronograma = tarefa_data.get("cronograma", {})
                        tarefa_widgets["start_date"].insert(
                            0, cronograma.get("inicio", "")
                        )
                        tarefa_widgets["end_date"].insert(0, cronograma.get("fim", ""))

                        for aloc_data in tarefa_data.get("alocacoes", []):
                            self.adicionar_alocacao_equipe(lote_nome, tarefa_nome)
                            nova_aloc = tarefa_widgets["alocacoes_widgets"][-1]
                            nome_func, cargo_func = aloc_data.get(
                                "funcionario", (None, None)
                            )
                            horas = aloc_data.get("horas_totais", "")
                            nova_aloc["combo_cargo"].set(cargo_func)
                            if nome_func != cargo_func:
                                nova_aloc["combo_nome"].set(nome_func)
                            nova_aloc["entry"].insert(0, str(horas).replace(".", ","))

                    # NOVO: Adiciona o botão APÓS recriar todas as tarefas do lote
                    self._criar_botao_adicionar_subcontrato(parent_frame, lote_nome)

            print("Estado carregado com sucesso.")
            self.processar_calculo()

        except Exception as e:
            messagebox.showerror(
                "Erro ao Carregar",
                f"Não foi possível carregar o estado salvo: {e}\n\nIniciando com uma nova sessão.",
                parent=self,
            )
            self.lote_widgets.clear()
            for i in range(len(self.notebook.tabs()) - 1, 0, -1):
                self.notebook.forget(i)

        finally:
            self._is_loading = False

    # MODIFICADO: Lógica de parsing no `processar_calculo`
    def processar_calculo(self):
        dados_coletados = self._coletar_dados_da_ui()
        if not dados_coletados:
            return  # A mensagem de erro já foi exibida durante a coleta

        try:
            horas_mes = dados_coletados["config"]["horas_mes"]
            self.app_controller.processar_portfolio(dados_coletados["lotes"], horas_mes)
        except Exception as e:
            messagebox.showerror(
                "Erro Inesperado", f"Ocorreu um erro no processamento: {e}", parent=self
            )

    # ... O restante do arquivo (dividir_tarefa, exportar_para_excel, etc.) pode permanecer o mesmo por enquanto
    # ... ou ser ajustado se a função dividir tarefa precisar de lógica de categoria
    def dividir_tarefa(self, lote_nome, disciplina, aloc_origem):
        # A lógica de dividir tarefa pode se tornar mais complexa.
        # Por agora, ela vai adicionar um novo funcionário (que precisa de uma categoria)
        # e alocá-lo à mesma tarefa. A filtragem cuidará do resto.
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

        # O novo funcionário será criado na mesma disciplina da tarefa atual.
        dialog = CustomDialog(
            self, self.app_controller.get_cargos_disponiveis(), [disciplina]
        )
        if not dialog.result:
            return

        novo_nome, novo_cargo, nova_disciplina = dialog.result
        self.app_controller.adicionar_funcionario(
            novo_nome, novo_cargo, nova_disciplina
        )

        # O novo funcionário agora aparecerá na lista filtrada.
        self.adicionar_alocacao_equipe(lote_nome, disciplina)

        nova_aloc_widget = self.lote_widgets[lote_nome]["disciplinas"][disciplina][
            "alocacoes_widgets"
        ][-1]

        horas_divididas = horas_originais / 2.0
        aloc_origem["entry"].delete(0, tk.END)
        aloc_origem["entry"].insert(0, f"{horas_divididas:.2f}".replace(".", ","))

        novo_display_string = f"{novo_nome} ({novo_cargo})"
        nova_aloc_widget["combo"].set(novo_display_string)
        nova_aloc_widget["entry"].delete(0, tk.END)
        nova_aloc_widget["entry"].insert(0, f"{horas_divididas:.2f}".replace(".", ","))

    # ...
    # O restante dos métodos (gerar_abas_lotes, _preencher_treeview, atualizar_dashboards, exportar_para_excel, etc)
    # permanecem inalterados por enquanto, pois a lógica de processamento final no core já agrupa por disciplina.
    # ... (cole o restante dos métodos inalterados aqui)
    def gerar_abas_lotes(self, popular_defaults=True):
        # Limpa abas e widgets antigos
        for i in range(len(self.notebook.tabs()) - 1, 0, -1):
            self.notebook.forget(i)
        self.lote_widgets.clear()

        try:
            num_lotes = int(self.num_lotes_entry.get())
            if num_lotes <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Erro", "Insira um número válido e positivo de lotes.")
            return

        # Cria as novas abas
        for i in range(num_lotes):
            lote_nome = str(i + 1)
            lote_tab_frame = ttk.Frame(self.notebook)
            self.notebook.add(lote_tab_frame, text=f"Lote {lote_nome}")
            lote_notebook = ttk.Notebook(lote_tab_frame)
            lote_notebook.pack(fill=tk.BOTH, expand=True)
            entrada_tab = ttk.Frame(lote_notebook, padding=10)
            dashboard_lote_tab = ttk.Frame(lote_notebook, padding=5)
            lote_notebook.add(entrada_tab, text="Entrada de Dados")
            lote_notebook.add(dashboard_lote_tab, text="Dashboard do Lote")

            # --- INÍCIO DA CORREÇÃO ---
            # 1. CRIA a estrutura de dados do lote PRIMEIRO.
            self.lote_widgets[lote_nome] = {
                "nome": lote_nome,
                "disciplinas": {},  # A chave "disciplinas" é criada vazia aqui.
                "dash_tree": self._criar_treeview(dashboard_lote_tab),
                "scrollable_frame": None,  # Deixa um espaço reservado.
            }

            # 2. CHAMA a função que vai preencher a UI e a estrutura de dados já criada.
            scrollable_frame = self.criar_conteudo_aba_entrada(
                entrada_tab, lote_nome, popular_defaults=popular_defaults
            )

            # 3. ATUALIZA o dicionário com a referência correta para o frame de rolagem.
            self.lote_widgets[lote_nome]["scrollable_frame"] = scrollable_frame
            # --- FIM DA CORREÇÃO ---

        self.atualizar_lista_funcionarios()

    def _popular_lotes_com_defaults(self):
        cargos_padrao = ["Sênior", "Pleno", "Júnior", "Estagiário"]
        if not self.lote_widgets:
            return
        for lote_nome in self.lote_widgets.keys():
            for disc in self.app_controller.get_disciplinas():
                for cargo in cargos_padrao:
                    self.adicionar_alocacao_equipe(
                        lote_nome, disc, cargo_a_definir=cargo
                    )

    ### MODIFICADO: Cria 4 funcionários padrão e corrige o botão de remover. ###
    def criar_conteudo_aba_entrada(self, parent_tab, lote_nome, popular_defaults=True):
        # Esta parte inicial que cria a área de rolagem está correta.
        canvas = tk.Canvas(parent_tab, highlightthickness=0, bg="#FFFFFF")
        scrollbar = ttk.Scrollbar(parent_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Modern.TFrame")

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

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

        # Se for um lote novo, cria as disciplinas padrão e o botão no final
        if popular_defaults:
            disciplinas_iniciais = self.app_controller.get_disciplinas()
            for disc in disciplinas_iniciais:
                self._criar_frame_tarefa(
                    scrollable_frame, lote_nome, disc, popular_defaults=True
                )

            # Adiciona o botão APÓS as disciplinas padrão
            self._criar_botao_adicionar_subcontrato(scrollable_frame, lote_nome)

        # Adicionamos um retorno para que a função de carregamento possa obter a referência ao frame
        return scrollable_frame

    # A função não precisa mais retornar um dicionário

    def _criar_frame_tarefa(
        self, parent_frame, lote_nome, nome_tarefa, popular_defaults=True
    ):
        # ... (todo o código de criação de widgets permanece o mesmo) ...
        # ... (a criação de disc_widgets_dict permanece a mesma) ...

        disc_frame = ttk.LabelFrame(parent_frame, text=nome_tarefa, padding=10)
        disc_frame.pack(fill=tk.X, expand=True, padx=10, pady=(10, 5))

        disc_frame.columnconfigure(0, weight=1)
        date_frame = ttk.Frame(disc_frame)
        date_frame.grid(row=0, column=0, sticky="w", pady=(0, 5))

        ttk.Label(date_frame, text="Início (DD/MM/AAAA):").pack(side=tk.LEFT)
        start_date_entry = ttk.Entry(date_frame, width=15)
        start_date_entry.pack(side=tk.LEFT, padx=5)
        start_date_entry.bind("<KeyRelease>", self._on_data_keyrelease)

        ttk.Label(date_frame, text="Fim (DD/MM/AAAA):").pack(side=tk.LEFT, padx=10)
        end_date_entry = ttk.Entry(date_frame, width=15)
        end_date_entry.pack(side=tk.LEFT, padx=5)
        end_date_entry.bind("<KeyRelease>", self._on_data_keyrelease)

        aloc_frame = ttk.Frame(disc_frame)
        aloc_frame.grid(row=1, column=0, sticky="ew", pady=5)

        disc_widgets_dict = {
            "start_date": start_date_entry,
            "end_date": end_date_entry,
            "alocacoes_frame": aloc_frame,
            "alocacoes_widgets": [],
            "frame": disc_frame,
        }

        self.lote_widgets[lote_nome]["disciplinas"][nome_tarefa] = disc_widgets_dict

        # --- INÍCIO DA CORREÇÃO ---
        # O comando agora "congela" os identificadores, não o dicionário inteiro.
        calculo_cmd = partial(
            self._recalcular_todas_as_horas_da_tarefa, lote_nome, nome_tarefa
        )
        # --- FIM DA CORREÇÃO ---

        start_date_entry.bind("<FocusOut>", calculo_cmd)
        end_date_entry.bind("<FocusOut>", calculo_cmd)

        if popular_defaults:
            cargos_padrao = [
                "Eng. Sênior",
                "Eng. Pleno",
                "Eng. Júnior",
                "Estagiário/Projetista",
            ]
            cargos_disponiveis = self.app_controller.get_cargos_disponiveis()
            for cargo_padrao in cargos_disponiveis:
                if cargo_padrao in cargos_disponiveis:
                    self.adicionar_alocacao_equipe(
                        lote_nome, nome_tarefa, cargo_a_definir=cargo_padrao
                    )

        cmd_add_func = partial(
            self._on_adicionar_alocacao_click, lote_nome, nome_tarefa
        )
        ttk.Button(disc_frame, text="+ Alocar Equipe", command=cmd_add_func).grid(
            row=2, column=0, sticky="w", pady=5
        )

    def dividir_tarefa(self, lote_nome, disciplina, aloc_origem_widget):
        try:
            horas_str = aloc_origem_widget["entry"].get().replace(",", ".")
            horas_originais = float(horas_str or 0.0)
            if horas_originais <= 0:
                messagebox.showwarning(
                    "Aviso",
                    "A tarefa precisa ter horas positivas para ser dividida.",
                    parent=self,
                )
                return

            horas_divididas = horas_originais / 2.0

            aloc_origem_widget["entry"].delete(0, tk.END)
            aloc_origem_widget["entry"].insert(
                0, f"{horas_divididas:.2f}".replace(".", ",")
            )

            self.adicionar_alocacao_equipe(
                lote_nome, disciplina, horas_a_definir=horas_divididas
            )

            nova_aloc_widget = self.lote_widgets[lote_nome]["disciplinas"][disciplina][
                "alocacoes_widgets"
            ][-1]

            cargo_original = aloc_origem_widget["combo_cargo"].get()
            nova_aloc_widget["combo_cargo"].set(cargo_original)

        except (ValueError, KeyError) as e:
            messagebox.showerror(
                "Erro ao Dividir",
                f"Não foi possível dividir a tarefa: {e}",
                parent=self,
            )

    ### MODIFICADO: Usa StringVar para os campos de data e adiciona a trace ###
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

        # --- Início da Modificação ---
        ttk.Label(row_frame, text="Início:").pack(side=tk.LEFT, padx=(10, 0))
        start_entry = ttk.Entry(row_frame, width=12)
        start_entry.pack(side=tk.LEFT, padx=5)
        start_entry.bind("<KeyRelease>", self._on_data_keyrelease)  # ADICIONADO

        ttk.Label(row_frame, text="Fim:").pack(side=tk.LEFT, padx=(10, 0))
        end_entry = ttk.Entry(row_frame, width=12)
        end_entry.pack(side=tk.LEFT, padx=5)
        end_entry.bind("<KeyRelease>", self._on_data_keyrelease)  # ADICIONADO
        # --- Fim da Modificação ---

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
        df_sorted = df_relatorio.sort_values(by=["Disciplina", "Cargo", "Funcionário"])
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
        self,
        df_consolidado,
        dashboards_lotes,
        detalhes_tarefas=None,
        alerta_composicao=None,
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

        if alerta_composicao:
            porcentagem_global = alerta_composicao.get("global", 0.0)
            detalhes_lotes = alerta_composicao.get("lotes", {})

            # Verifica se algum lote individual excedeu o limite
            lotes_excedidos = [
                lote for lote, perc in detalhes_lotes.items() if perc >= 50.0
            ]

            # CONDIÇÃO DE ALERTA: global >= 50% OU algum lote individual >= 50%
            if porcentagem_global >= 50.0 or lotes_excedidos:
                # Monta a mensagem principal
                if porcentagem_global >= 50.0:
                    mensagem_alerta = (
                        f"Atenção: A participação global de Estagiários/Projetistas é de "
                        f"{porcentagem_global:.1f}%, excedendo o limite de 50%."
                    ).replace(".", ",")
                else:  # Se chegou aqui, é porque apenas lotes individuais excederam
                    mensagem_alerta = (
                        f"Atenção: Um ou mais lotes excederam o limite individual de 50% de horas "
                        f"para Estagiários/Projetistas (Global: {porcentagem_global:.1f}%)."
                    ).replace(".", ",")

                # Monta o detalhamento
                mensagem_alerta += "\n\nDetalhes da Participação por Lote:"

                lotes_ordenados = sorted(
                    detalhes_lotes.items(), key=lambda item: item[1], reverse=True
                )

                for nome_lote, percent_lote in lotes_ordenados:
                    # Adiciona um marcador se o lote específico excedeu o limite
                    marcador = "  <<-- LIMITE EXCEDIDO!" if percent_lote >= 50.0 else ""
                    mensagem_alerta += (
                        f"\n- Lote {nome_lote}: {percent_lote:.1f}%{marcador}".replace(
                            ".", ","
                        )
                    )

                messagebox.showwarning(
                    "Alerta de Composição da Equipe", mensagem_alerta, parent=self
                )

    def exportar_para_excel(self):
        df_consolidado = self.app_controller.get_ultimo_df_consolidado()
        dashboards_lotes = self.app_controller.get_ultimo_dashboards_lotes()
        if df_consolidado is None or df_consolidado.empty:
            messagebox.showwarning(
                "Aviso",
                "Não há dados consolidados para exportar. Calcule a alocação primeiro.",
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
                # 1. Escreve a planilha consolidada e cria o gráfico para ela
                self._write_styled_df_to_excel(writer, "Consolidado", df_consolidado)
                self._create_chart_for_sheet(writer, "Consolidado", df_consolidado)

                # 2. Escreve as planilhas de lote e seus gráficos
                if dashboards_lotes:
                    for nome_lote, df_lote in dashboards_lotes.items():
                        if not df_lote.empty:
                            sheet_name = f"Lote {nome_lote}"
                            self._write_styled_df_to_excel(writer, sheet_name, df_lote)
                            self._create_chart_for_sheet(writer, sheet_name, df_lote)

                # 3. Remove a planilha padrão APENAS DEPOIS de criar as novas
                if "Sheet" in writer.book.sheetnames:
                    writer.book.remove(writer.book["Sheet"])

            messagebox.showinfo(
                "Sucesso", f"Relatório exportado com sucesso para:\n{filepath}"
            )
        except Exception as e:
            messagebox.showerror(
                "Erro na Exportação", f"Ocorreu um erro ao salvar o arquivo:\n{e}"
            )

    def _write_styled_df_to_excel(self, writer, sheet_name, df):
        # Remove a coluna de status que não é necessária no Excel
        df_para_exportar = df.drop(columns=["Status"], errors="ignore")
        if df_para_exportar.empty:
            return

        # Ordena por Disciplina para o agrupamento
        df_sorted = df_para_exportar.sort_values(
            by=["Disciplina", "Cargo", "Funcionário"]
        )

        # Escreve os dados na planilha
        df_sorted.to_excel(
            writer, sheet_name=sheet_name, index=False, header=False, startrow=1
        )
        ws = writer.sheets[sheet_name]
        ws.sheet_view.showGridLines = False

        # --- ESTILOS PADRONIZADOS ---
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(
            start_color="44546A", end_color="44546A", fill_type="solid"
        )
        category_font = Font(name="Arial", bold=True, size=12)
        category_fill = PatternFill(
            start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"
        )  # Azul claro
        data_font = Font(name="Arial", size=10)
        thin_side = Side(border_style="thin", color="BFBFBF")
        full_border = Border(
            left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
        )

        # Escreve e formata o cabeçalho
        headers = list(df_sorted.columns)
        for c_idx, header_text in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c_idx, value=header_text)
            cell.font = header_font
            cell.fill = header_fill

        # Agrupamento por Disciplina
        last_discipline = None
        current_row_offset = 0
        # Itera sobre as linhas de dados para inserir os cabeçalhos de categoria
        for r_idx, row in enumerate(
            dataframe_to_rows(df_sorted, index=False, header=False), 2
        ):
            if row[0] != last_discipline:
                ws.insert_rows(r_idx + current_row_offset)

                cat_cell = ws.cell(
                    row=r_idx + current_row_offset, column=1, value=row[0]
                )
                cat_cell.font = category_font
                cat_cell.fill = category_fill
                ws.merge_cells(
                    start_row=r_idx + current_row_offset,
                    start_column=1,
                    end_row=r_idx + current_row_offset,
                    end_column=len(headers),
                )

                last_discipline = row[0]
                current_row_offset += 1

        # Aplica bordas e fontes a todas as células com conteúdo
        for row_cells in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row_cells:
                if cell.value is not None:
                    cell.border = full_border
                    if not cell.font.bold:
                        cell.font = data_font

        # Ajusta a largura das colunas
        for i, col_name in enumerate(headers, 1):
            column_letter = openpyxl.utils.get_column_letter(i)
            max_len = max(df_sorted[col_name].astype(str).map(len).max(), len(col_name))
            ws.column_dimensions[column_letter].width = max(max_len + 4, 12)

    def _create_chart_for_sheet(self, writer, sheet_name, df):
        if df.empty:
            return

        ws = writer.book[sheet_name]

        df_numeric = df.copy()
        month_cols_formatted = [
            col
            for col in df.columns
            if "/" in col and col != "H.mês" and "Cargo" not in col
        ]

        if not month_cols_formatted:
            return

        for col in month_cols_formatted:
            df_numeric[col] = pd.to_numeric(
                df_numeric[col].str.replace(",", "."), errors="coerce"
            )

        monthly_decimal_totals = df_numeric[month_cols_formatted].sum()

        if monthly_decimal_totals.sum() == 0:
            return

        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = f"Histograma de pessoal - {sheet_name}\n(Homen/mês)"
        chart.title.layout = Layout(manualLayout=ManualLayout(y=0, yMode="edge"))
        chart.legend = None

        chart.y_axis.majorGridlines = None
        axis_line_props = LineProperties(solidFill="000000")
        chart.y_axis.spPr = GraphicalProperties(ln=axis_line_props)
        chart.x_axis.spPr = GraphicalProperties(ln=axis_line_props)
        chart.x_axis.delete = False
        chart.y_axis.delete = False

        # --- INÍCIO DAS MODIFICAÇÕES ---

        # 2. Mover o título para cima ajustando o layout da área de plotagem
        # Isso move a área do gráfico (as barras) para baixo, dando mais espaço para o título.
        chart.layout = Layout(
            manualLayout=ManualLayout(
                y=0.15,  # Posição Y da área de plotagem (15% a partir do topo)
                h=0.75,  # Altura da área de plotagem (75% da altura total do gráfico)
            )
        )

        # --- FIM DAS MODIFICAÇÕES ---

        data_start_col = len(df.columns) + 2
        chart_data_start_row = 2

        data_rows = [["Mês", "Total Decimal"]] + list(
            zip(monthly_decimal_totals.index, monthly_decimal_totals.values)
        )
        for r_idx, row_data in enumerate(data_rows, start=chart_data_start_row):
            for c_idx, value in enumerate(row_data, start=data_start_col):
                ws.cell(row=r_idx, column=c_idx, value=value)

        values = Reference(
            ws,
            min_col=data_start_col + 1,
            min_row=chart_data_start_row,
            max_row=chart_data_start_row + len(monthly_decimal_totals),
        )
        cats = Reference(
            ws,
            min_col=data_start_col,
            min_row=chart_data_start_row + 1,
            max_row=chart_data_start_row + len(monthly_decimal_totals),
        )

        series = Series(values, title_from_data=True)

        series.graphicalProperties.solidFill = "44546A"
        series.graphicalProperties.line.solidFill = "44546A"

        chart.append(series)
        chart.set_categories(cats)

        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = True
        chart.dLbls.showSerName = False
        chart.dLbls.showCatName = False
        chart.dLbls.showLegendKey = False
        chart.dLbls.position = "outEnd"
        chart.dLbls.numFmt = "0.00"

        ws.add_chart(chart, f"A{ws.max_row + 5}")

# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext


class GuiApp(tk.Tk):
    """
    Interface gráfica com abas para gerenciar múltiplos Lotes.
    """

    def __init__(self, app_controller):
        super().__init__()
        self.app_controller = app_controller

        self.title("Planejador de Portfólio de Projetos")
        self.geometry("1200x800")

        # --- Frames ---
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        setup_frame = ttk.LabelFrame(
            main_frame, text="Configuração Inicial", padding="10"
        )
        setup_frame.pack(fill=tk.X, pady=5)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=5)

        dashboard_frame = ttk.LabelFrame(
            main_frame, text="Dashboard Consolidado", padding="10"
        )
        dashboard_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # --- Configuração Inicial ---
        ttk.Label(setup_frame, text="Número de Lotes:").grid(row=0, column=0, padx=5)
        self.num_lotes_entry = ttk.Entry(setup_frame, width=10)
        self.num_lotes_entry.grid(row=0, column=1, padx=5)
        ttk.Button(setup_frame, text="Criar Lotes", command=self.criar_abas_lotes).grid(
            row=0, column=2, padx=5
        )

        ttk.Label(setup_frame, text="Horas/Mês por Funcionário:").grid(
            row=0, column=3, padx=5
        )
        self.horas_mes_entry = ttk.Entry(setup_frame, width=10)
        self.horas_mes_entry.insert(0, "160")  # Valor Padrão
        self.horas_mes_entry.grid(row=0, column=4, padx=5)

        ttk.Button(
            setup_frame,
            text="CALCULAR PORTFÓLIO",
            command=self.processar_calculo,
            style="Accent.TButton",
        ).grid(row=0, column=5, padx=10)
        self.style = ttk.Style(self)
        self.style.configure("Accent.TButton", font=("Helvetica", 10, "bold"))

        # --- Dashboard ---
        self.alertas_listbox = scrolledtext.ScrolledText(
            dashboard_frame, height=5, wrap=tk.WORD, state=tk.DISABLED, bg="#f0f0f0"
        )
        self.alertas_listbox.pack(fill=tk.X, pady=5)

        self.tree = ttk.Treeview(dashboard_frame)
        self.tree.pack(fill=tk.BOTH, expand=True, pady=5)

        self.entries = {"lotes": []}

    def criar_abas_lotes(self):
        # Limpa abas existentes
        for i in range(self.notebook.index("end")):
            self.notebook.forget(0)
        self.entries["lotes"] = []

        try:
            num_lotes = int(self.num_lotes_entry.get())
            for i in range(num_lotes):
                lote_nome = f"Lote {i+1}"
                tab = ttk.Frame(self.notebook, padding="10")
                self.notebook.add(tab, text=lote_nome)
                self.criar_conteudo_aba(tab, i)
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira um número válido de lotes.")

    def criar_conteudo_aba(self, tab, lote_index):
        lote_entries = {"cronograma": {}, "alocacao": {}}

        headers = [
            "Disciplina",
            "Data Início (DD/MM/AAAA)",
            "Data Fim (DD/MM/AAAA)",
        ] + self.app_controller.classes_func
        for i, header in enumerate(headers):
            ttk.Label(tab, text=header, font=("Helvetica", 9, "bold")).grid(
                row=0, column=i, padx=5, pady=2, sticky=tk.W
            )

        for i, disc in enumerate(self.app_controller.disciplinas):
            ttk.Label(tab, text=disc).grid(
                row=i + 1, column=0, sticky=tk.W, padx=5, pady=2
            )

            lote_entries["cronograma"][disc] = {}
            start_date_entry = ttk.Entry(tab, width=20)
            start_date_entry.grid(row=i + 1, column=1, padx=5, pady=2)
            lote_entries["cronograma"][disc]["inicio"] = start_date_entry

            end_date_entry = ttk.Entry(tab, width=20)
            end_date_entry.grid(row=i + 1, column=2, padx=5, pady=2)
            lote_entries["cronograma"][disc]["fim"] = end_date_entry

            lote_entries["alocacao"][disc] = {}
            for j, classe in enumerate(self.app_controller.classes_func):
                entry = ttk.Entry(tab, width=15)
                entry.grid(row=i + 1, column=j + 3, padx=5, pady=2)
                lote_entries["alocacao"][disc][classe] = entry

        self.entries["lotes"].append(lote_entries)

    def processar_calculo(self):
        try:
            horas_mes = int(self.horas_mes_entry.get() or 160)

            lotes_data = []
            for i, lote_entries in enumerate(self.entries["lotes"]):
                cronograma = {}
                for disc, dates in lote_entries["cronograma"].items():
                    cronograma[disc] = {
                        "inicio": dates["inicio"].get(),
                        "fim": dates["fim"].get(),
                    }

                alocacao = {}
                for disc, classes in lote_entries["alocacao"].items():
                    alocacao[disc] = {}
                    for classe, entry in classes.items():
                        alocacao[disc][classe] = float(
                            entry.get().replace(",", ".") or 0.0
                        )

                lotes_data.append(
                    {
                        "nome": f"Lote {i+1}",
                        "cronograma": cronograma,
                        "alocacao": alocacao,
                    }
                )

            self.app_controller.processar_portfolio(lotes_data, horas_mes)

        except ValueError:
            messagebox.showerror(
                "Erro de Entrada",
                "Alocação deve ser um número decimal (ex: 0.5 ou 0,5) e Horas/Mês deve ser um número inteiro.",
            )
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro: {e}")

    def atualizar_dashboard(self, df_resultado, alertas):
        # Atualiza alertas
        self.alertas_listbox.config(state=tk.NORMAL)
        self.alertas_listbox.delete("1.0", tk.END)
        if alertas:
            self.alertas_listbox.insert(
                "1.0", "ALERTAS DE ALOCAÇÃO:\n" + "\n".join(alertas)
            )
            self.alertas_listbox.config(fg="red")
        else:
            self.alertas_listbox.insert(
                "1.0", "Nenhum conflito de alocação encontrado."
            )
            self.alertas_listbox.config(fg="green")
        self.alertas_listbox.config(state=tk.DISABLED)

        # Atualiza a tabela
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df_resultado.columns)
        self.tree["show"] = "headings"

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)

        for _, row in df_resultado.iterrows():
            self.tree.insert("", "end", values=list(row))

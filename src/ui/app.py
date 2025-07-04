# src/ui/app.py
# Este módulo contém a classe da aplicação da interface gráfica (GUI).

import customtkinter as ctk
from tkinter import filedialog, messagebox
from core.scheduler import ProjectScheduler


class App(ctk.CTk):
    """
    Classe principal da aplicação que cria e gerencia a interface do usuário.
    """

    def __init__(self):
        super().__init__()

        self.title("Gerador de Cronograma de Projeto")
        self.geometry("1000x800")

        # --- Estrutura Principal ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Frame rolável para conter todos os inputs
        self.scrollable_frame = ctk.CTkScrollableFrame(
            self, label_text="Parâmetros de Entrada"
        )
        self.scrollable_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        self.input_widgets = {}
        self.categories = [
            "Geometria",
            "Terraplenagem",
            "Drenagem",
            "Pavimento",
            "Sinalização",
            "Geologia",
            "Geotecnia",
            "Estruturas",
        ]
        self.employee_classes = [
            "Projetista/Estagiário",
            "Eng. Júnior",
            "Eng. Pleno",
            "Eng. Sênior",
        ]

        # --- Seção de Inputs Globais ---
        self._create_global_inputs_section()

        # --- Seção de Inputs por Categoria ---
        self._create_categories_section()

        # --- Botão de Ação ---
        self.generate_button = ctk.CTkButton(
            self, text="Gerar Cronograma", command=self.generate_schedule_callback
        )
        self.generate_button.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

    def _create_global_inputs_section(self):
        """Cria os widgets para os parâmetros globais."""
        globals_frame = ctk.CTkFrame(self.scrollable_frame)
        globals_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        globals_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkLabel(
            globals_frame, text="Configurações Gerais", font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=4, pady=5)

        self.input_widgets["globals"] = {}

        # Horas por mês
        ctk.CTkLabel(globals_frame, text="Horas/Mês/Pessoa:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        self.input_widgets["globals"]["total_hours"] = ctk.CTkEntry(
            globals_frame, placeholder_text="Ex: 160"
        )
        self.input_widgets["globals"]["total_hours"].grid(
            row=1, column=1, padx=5, pady=5, sticky="ew"
        )

        # Número de Lotes
        ctk.CTkLabel(globals_frame, text="Quantidade de Lotes:").grid(
            row=1, column=2, padx=5, pady=5, sticky="w"
        )
        self.input_widgets["globals"]["num_lots"] = ctk.CTkEntry(
            globals_frame, placeholder_text="Ex: 3"
        )
        self.input_widgets["globals"]["num_lots"].grid(
            row=1, column=3, padx=5, pady=5, sticky="ew"
        )

        # Percentuais máximos
        ctk.CTkLabel(
            globals_frame,
            text="Percentual Máximo de Horas por Classe (%)",
            font=ctk.CTkFont(weight="bold"),
        ).grid(row=2, column=0, columnspan=4, pady=(15, 5))
        self.input_widgets["globals"]["max_percentages"] = {}
        for i, emp_class in enumerate(self.employee_classes):
            ctk.CTkLabel(globals_frame, text=f"{emp_class}:").grid(
                row=3 + i, column=0, padx=5, pady=5, sticky="w"
            )
            entry = ctk.CTkEntry(globals_frame, placeholder_text="Ex: 50")
            entry.grid(row=3 + i, column=1, padx=5, pady=5, sticky="ew")
            self.input_widgets["globals"]["max_percentages"][emp_class] = entry

    def _create_categories_section(self):
        """Cria os frames e widgets de input para cada categoria."""
        self.input_widgets["categories"] = {}

        for i, cat_name in enumerate(self.categories):
            cat_frame = ctk.CTkFrame(self.scrollable_frame)
            cat_frame.grid(row=i + 1, column=0, padx=10, pady=10, sticky="ew")
            cat_frame.grid_columnconfigure((1, 3), weight=1)

            ctk.CTkLabel(
                cat_frame, text=cat_name, font=ctk.CTkFont(size=14, weight="bold")
            ).grid(row=0, column=0, columnspan=4, pady=5)

            self.input_widgets["categories"][cat_name] = {"total_hours": {}}

            # Inputs de Data
            ctk.CTkLabel(cat_frame, text="Data Início:").grid(
                row=1, column=0, padx=5, pady=5, sticky="w"
            )
            self.input_widgets["categories"][cat_name]["start_date"] = ctk.CTkEntry(
                cat_frame, placeholder_text="AAAA-MM-DD"
            )
            self.input_widgets["categories"][cat_name]["start_date"].grid(
                row=1, column=1, padx=5, pady=5, sticky="ew"
            )

            ctk.CTkLabel(cat_frame, text="Data Fim:").grid(
                row=1, column=2, padx=5, pady=5, sticky="w"
            )
            self.input_widgets["categories"][cat_name]["end_date"] = ctk.CTkEntry(
                cat_frame, placeholder_text="AAAA-MM-DD"
            )
            self.input_widgets["categories"][cat_name]["end_date"].grid(
                row=1, column=3, padx=5, pady=5, sticky="ew"
            )

            # Lotes Aplicáveis
            ctk.CTkLabel(cat_frame, text="Lotes Aplicáveis:").grid(
                row=2, column=0, padx=5, pady=5, sticky="w"
            )
            self.input_widgets["categories"][cat_name]["applicable_lots"] = (
                ctk.CTkEntry(cat_frame, placeholder_text="Todos (ou ex: 1, 3)")
            )
            self.input_widgets["categories"][cat_name]["applicable_lots"].grid(
                row=2, column=1, columnspan=3, padx=5, pady=5, sticky="ew"
            )

            # MUDANÇA: Inputs de Horas Totais em vez de Nº de Funcionários
            ctk.CTkLabel(
                cat_frame,
                text="Total de Horas/Mês por Classe:",
                font=ctk.CTkFont(weight="bold"),
            ).grid(row=3, column=0, columnspan=4, pady=(10, 0))
            for j, emp_class in enumerate(self.employee_classes):
                row, col = (4 + j // 2, (j % 2) * 2)
                ctk.CTkLabel(cat_frame, text=f"{emp_class}:").grid(
                    row=row, column=col, padx=5, pady=5, sticky="w"
                )
                entry = ctk.CTkEntry(cat_frame, placeholder_text="Ex: 320")
                entry.grid(row=row, column=col + 1, padx=5, pady=5, sticky="ew")
                self.input_widgets["categories"][cat_name]["total_hours"][
                    emp_class
                ] = entry

    def _collect_input_data(self):
        """Coleta todos os dados dos widgets e os estrutura em um dicionário."""
        data = {"globals": {}, "categories": {}}

        # Coleta dados globais
        global_inputs = self.input_widgets["globals"]
        data["globals"]["total_hours"] = global_inputs["total_hours"].get()
        data["globals"]["num_lots"] = global_inputs["num_lots"].get()
        data["globals"]["max_percentages"] = {
            cls: entry.get() or "0"
            for cls, entry in global_inputs["max_percentages"].items()
        }

        # Coleta dados das categorias
        for cat_name, widgets in self.input_widgets["categories"].items():
            data["categories"][cat_name] = {
                "start_date": widgets["start_date"].get(),
                "end_date": widgets["end_date"].get(),
                "applicable_lots": widgets["applicable_lots"].get(),
                "total_hours": {
                    emp_class: float(entry.get() or 0)
                    for emp_class, entry in widgets["total_hours"].items()
                },
            }
        return data

    def generate_schedule_callback(self):
        """Função chamada quando o botão 'Gerar Cronograma' é pressionado."""
        try:
            input_data = self._collect_input_data()

            # Instancia e executa a lógica de negócio
            scheduler = ProjectScheduler(input_data)
            schedule_df, warnings = scheduler.generate_schedule()

            if schedule_df.empty:
                messagebox.showwarning(
                    "Nenhum Dado",
                    "O cronograma não pôde ser gerado. Verifique se as datas, lotes e horas inseridas estão corretos.",
                )
                return

            # Pede ao usuário para escolher onde salvar o arquivo
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Salvar Cronograma Como...",
            )

            if not file_path:
                return  # Usuário cancelou a operação de salvar

            # Salva o DataFrame em um arquivo Excel
            schedule_df.to_excel(file_path, index=False, engine="openpyxl")

            # Monta a mensagem final
            final_message = "Cronograma gerado e salvo com sucesso!"
            if warnings:
                final_message += (
                    "\n\nAtenção para os seguintes avisos:\n- " + "\n- ".join(warnings)
                )
                messagebox.showwarning("Sucesso com Avisos", final_message)
            else:
                messagebox.showinfo("Sucesso", final_message)

        except (ValueError, TypeError) as e:
            messagebox.showerror(
                "Erro nos Dados de Entrada",
                f"Por favor, verifique os dados inseridos.\nDetalhe: {e}",
            )
        except Exception as e:
            messagebox.showerror(
                "Erro Inesperado",
                f"Ocorreu um erro inesperado durante a geração do cronograma:\n{e}",
            )

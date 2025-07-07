# -*- coding: utf-8 -*-

from core.project import Portfolio
from .gui_app import GuiApp


class App:
    """
    Controlador que conecta a UI (com gestão de pessoas) com a lógica de negócio.
    """

    def __init__(self):
        self.disciplinas = [
            "Geometria",
            "Terraplenagem",
            "Drenagem",
            "Pavimento",
            "Sinalização",
            "Geologia",
            "Geotecnia",
            "Estruturas",
            "Coordenação",
        ]
        self.cargos_disponiveis = [
            "Projetista",
            "Eng. Civil",
            "Eng. Pleno",
            "Eng. Sênior",
            "Coordenador",
        ]

        self.funcionarios = []
        self.portfolio = Portfolio(self.disciplinas)
        self.ui = GuiApp(app_controller=self)

    def run(self):
        """Inicia o loop principal da interface gráfica."""
        self.ui.mainloop()

    def get_funcionarios_para_display(self):
        """Formata a lista de funcionários para exibição na UI (ex: 'Nome (Cargo)')."""
        return sorted([f"{nome} ({cargo})" for nome, cargo in self.funcionarios])

    def adicionar_funcionario(self, nome, cargo):
        """Adiciona um funcionário (nome, cargo) à lista mestra e atualiza a UI."""
        novo_funcionario = (nome, cargo)
        if novo_funcionario not in self.funcionarios:
            self.funcionarios.append(novo_funcionario)
            self.funcionarios.sort(key=lambda x: x[0])
            self.ui.atualizar_lista_funcionarios(self.get_funcionarios_para_display())

    def remover_funcionario(self, display_string):
        """Remove um funcionário da lista mestra com base na string de exibição."""
        try:
            nome, resto = display_string.rsplit(" (", 1)
            cargo = resto.rstrip(")")
            funcionario_a_remover = (nome, cargo)
            if funcionario_a_remover in self.funcionarios:
                self.funcionarios.remove(funcionario_a_remover)
                self.ui.atualizar_lista_funcionarios(
                    self.get_funcionarios_para_display()
                )
        except ValueError:
            print(
                f"Erro ao tentar remover: formato de string inválido '{display_string}'"
            )

    def processar_portfolio(self, lotes_data, horas_mes):
        """
        Recebe os dados de todos os lotes, processa-os e atualiza a UI.
        """
        self.portfolio.definir_configuracoes_gerais(horas_mes, self.funcionarios)
        self.portfolio.definir_dados_lotes(lotes_data)

        # Gera os dois relatórios
        df_dashboard = self.portfolio.gerar_relatorio_alocacao_decimal()
        detalhes_tarefas = self.portfolio.gerar_relatorio_detalhado_por_tarefa()

        # Passa ambos para a UI
        self.ui.atualizar_dashboard(df_dashboard, detalhes_tarefas)

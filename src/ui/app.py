# -*- coding: utf-8 -*-

from config.settings import DISCIPLINAS
from core.project import Portfolio
from .gui_app import GuiApp


class App:
    """
    Controlador que conecta a UI (com gestão de pessoas) com a lógica de negócio.
    """

    def __init__(self):

        self.funcionarios = []
        self.portfolio = Portfolio(DISCIPLINAS)
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

        # 1. Gerar o dashboard consolidado
        df_dashboard_consolidado = self.portfolio.gerar_relatorio_alocacao_decimal()

        # 2. Gerar o detalhamento de horas por tarefa
        detalhes_tarefas = self.portfolio.gerar_relatorio_detalhado_por_tarefa()

        # 3. Gerar um dashboard para cada lote
        dashboards_lotes = {}
        for lote in lotes_data:
            nome_lote = lote["nome"]
            df_lote = self.portfolio.gerar_relatorio_alocacao_decimal(
                nome_lote=nome_lote
            )
            dashboards_lotes[nome_lote] = df_lote

        # 4. Passar tudo para a UI atualizar
        self.ui.atualizar_dashboards(
            df_dashboard_consolidado, dashboards_lotes, detalhes_tarefas
        )

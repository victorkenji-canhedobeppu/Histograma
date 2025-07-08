# -*- coding: utf-8 -*-

from config.settings import CARGOS, DISCIPLINAS
from core.project import Portfolio
from .gui_app import GuiApp


class App:
    """
    Controlador que conecta a UI com a lógica de negócio.
    """

    def __init__(self):
        self.funcionarios = []
        self.portfolio = Portfolio(list(DISCIPLINAS.values()))
        self.ui = GuiApp(app_controller=self)
        self.ultimo_df_consolidado = None
        self.ultimo_dashboards_lotes = None

    def run(self):
        """Inicia o loop principal da interface gráfica."""
        self.ui.mainloop()

    def get_disciplinas(self):
        return list(DISCIPLINAS.values())

    def get_cargos_disponiveis(self):
        return sorted(list(CARGOS.values()), reverse=True)

    def get_ultimo_df_consolidado(self):
        return self.ultimo_df_consolidado

    def get_ultimo_dashboards_lotes(self):
        return self.ultimo_dashboards_lotes

    def get_funcionarios_para_display(self):
        """Formata a lista de funcionários para exibição na UI."""
        return sorted([f"{nome} ({cargo})" for nome, cargo in self.funcionarios])

    def adicionar_funcionario(self, nome, cargo):
        """Adiciona um funcionário à lista mestra e atualiza a UI."""
        novo_funcionario = (nome, cargo)
        if novo_funcionario not in self.funcionarios:
            self.funcionarios.append(novo_funcionario)
            self.funcionarios.sort(key=lambda x: x[0])
            self.ui.atualizar_lista_funcionarios(self.get_funcionarios_para_display())

    def remover_funcionario(self, display_string):
        """Remove um funcionário da lista mestra."""
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
        """Recebe os dados, processa-os e solicita a atualização da UI."""
        self.portfolio.definir_configuracoes_gerais(horas_mes, self.funcionarios)
        self.portfolio.definir_dados_lotes(lotes_data)

        self.ultimo_df_consolidado = self.portfolio.gerar_relatorio_alocacao_decimal()

        detalhes_tarefas = self.portfolio.gerar_relatorio_detalhado_por_tarefa()

        self.ultimo_dashboards_lotes = {}
        for lote in lotes_data:
            nome_lote = lote["nome"]
            df_lote = self.portfolio.gerar_relatorio_alocacao_decimal(
                nome_lote=nome_lote
            )
            self.ultimo_dashboards_lotes[nome_lote] = df_lote

        self.ui.atualizar_dashboards(
            self.ultimo_df_consolidado, self.ultimo_dashboards_lotes, detalhes_tarefas
        )

# -*- coding: utf-8 -*-

from config.settings import CARGOS, DISCIPLINAS, SUBCONTRATOS
from core.project import Portfolio
from ui.gui_app import GuiApp

# --- DICIONÁRIOS GLOBAIS ---


class App:
    def __init__(self):
        self.funcionarios = []
        self.portfolio = Portfolio(list(DISCIPLINAS.values()))
        # Passamos a instância de App para a GuiApp poder chamar seus métodos
        self.ui = GuiApp(app_controller=self)
        self.ultimo_df_consolidado = None
        self.ultimo_dashboards_lotes = None

    def run(self):
        # MODIFICADO: Carrega o estado anterior antes de iniciar o loop principal
        self.ui.carregar_estado()
        self.ui.mainloop()

    def get_disciplinas(self):
        return sorted(list(DISCIPLINAS.values()))

    def get_cargos_disponiveis(self):
        return sorted(list(CARGOS.values()), reverse=True)

    def get_subcontratos_disponiveis(self):
        return sorted(list(SUBCONTRATOS.values()))

    def get_ultimo_df_consolidado(self):
        return self.ultimo_df_consolidado

    def get_ultimo_dashboards_lotes(self):
        return self.ultimo_dashboards_lotes

    def get_funcionarios_para_display(self):
        return sorted([f"{nome} ({cargo})" for nome, cargo in self.funcionarios])

    def adicionar_funcionario(self, nome, cargo):
        novo_funcionario = (nome, cargo)
        if novo_funcionario not in self.funcionarios:
            self.funcionarios.append(novo_funcionario)
            self.funcionarios.sort(key=lambda x: x[0])
            self.ui.atualizar_lista_funcionarios(self.get_funcionarios_para_display())

    def remover_funcionario(self, display_string):
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
        self.portfolio.definir_configuracoes_gerais(horas_mes, self.funcionarios)
        self.portfolio.definir_dados_lotes(lotes_data)

        self.ultimo_df_consolidado = self.portfolio.gerar_relatorio_alocacao_decimal()

        # A função detalhada não é mais usada ativamente na UI, mas pode ser mantida para depuração
        detalhes_tarefas = self.portfolio.gerar_relatorio_detalhado_por_tarefa()

        self.ultimo_dashboards_lotes = {}
        for lote in lotes_data:
            nome_lote = lote["nome"]
            df_lote = self.portfolio.gerar_relatorio_alocacao_decimal(
                nome_lote=lote["nome"]
            )
            self.ultimo_dashboards_lotes[nome_lote] = df_lote

        self.ui.atualizar_dashboards(
            self.ultimo_df_consolidado, self.ultimo_dashboards_lotes, detalhes_tarefas
        )

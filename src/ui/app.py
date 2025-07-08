# -*- coding: utf-8 -*-

from config.settings import CARGOS, DISCIPLINAS, SUBCONTRATOS
from core.project import Portfolio
from ui.gui_app import GuiApp


class App:
    def __init__(self):
        self.funcionarios = []
        self.portfolio = Portfolio(list(DISCIPLINAS.values()))
        self.ui = GuiApp(app_controller=self)
        self.ultimo_df_consolidado = None
        self.ultimo_dashboards_lotes = None

    def run(self):
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
        return sorted(
            [
                f"{nome} ({cargo}) [{disciplina}]"
                for nome, cargo, disciplina in self.funcionarios
            ]
        )

    def get_funcionarios_para_display_por_disciplina(self, disciplina_filtro):
        funcionarios_filtrados = []
        for nome, cargo, disciplina in self.funcionarios:
            if disciplina == disciplina_filtro:
                funcionarios_filtrados.append(f"{nome} ({cargo})")
        return sorted(funcionarios_filtrados)

    def adicionar_funcionario(self, nome, cargo, disciplina):
        novo_funcionario = (nome, cargo, disciplina)
        if novo_funcionario not in self.funcionarios:
            self.funcionarios.append(novo_funcionario)
            self.funcionarios.sort(key=lambda x: x[0])
            self.ui.atualizar_lista_funcionarios()

    def remover_funcionario(self, display_string):
        try:
            resto, disciplina = display_string.rsplit(" [", 1)
            disciplina = disciplina.rstrip("]")
            nome, cargo = resto.rsplit(" (", 1)
            cargo = cargo.rstrip(")")

            funcionario_a_remover = (nome, cargo, disciplina)
            if funcionario_a_remover in self.funcionarios:
                self.funcionarios.remove(funcionario_a_remover)
                self.ui.atualizar_lista_funcionarios()
        except ValueError:
            print(
                f"Erro ao tentar remover: formato de string inválido '{display_string}'"
            )

    # MODIFICADO: Desempacota os resultados e passa para a UI.
    def processar_portfolio(self, lotes_data, horas_mes):
        self.portfolio.definir_configuracoes_gerais(horas_mes, self.funcionarios)
        self.portfolio.definir_dados_lotes(lotes_data)

        self.ultimo_df_consolidado = self.portfolio.gerar_relatorio_alocacao_decimal()
        detalhes_tarefas = self.portfolio.gerar_relatorio_detalhado_por_tarefa()

        self.ultimo_dashboards_lotes = {}
        for lote in lotes_data:
            nome_lote = lote["nome"]
            df_lote = self.portfolio.gerar_relatorio_alocacao_decimal(
                nome_lote=lote["nome"]
            )
            self.ultimo_dashboards_lotes[nome_lote] = df_lote

        # A verificação agora retorna uma tupla (global, detalhes_por_lote)
        porcentagem_global, detalhes_lotes = (
            self.portfolio.verificar_porcentagem_horas_cargo(cargo_alvo=CARGOS["ESTAG"])
        )

        # Monta um dicionário para passar para a UI
        alerta_composicao = {"global": porcentagem_global, "lotes": detalhes_lotes}

        self.ui.atualizar_dashboards(
            self.ultimo_df_consolidado,
            self.ultimo_dashboards_lotes,
            detalhes_tarefas,
            alerta_composicao=alerta_composicao,
        )

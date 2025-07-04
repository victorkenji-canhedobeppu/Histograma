# -*- coding: utf-8 -*-

from core.project import Portfolio
from ui.gui_app import GuiApp


class App:
    """
    Controlador que conecta a UI (com abas) com a lógica de negócio (Portfolio).
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
        ]
        self.classes_func = [
            "Projetista/Estagiário",
            "Eng. Júnior",
            "Eng. Pleno",
            "Eng. Sênior",
        ]

        self.portfolio = Portfolio(self.disciplinas, self.classes_func)
        self.ui = GuiApp(app_controller=self)

    def run(self):
        """Inicia o loop principal da interface gráfica."""
        self.ui.mainloop()

    def processar_portfolio(self, lotes_data, horas_mes):
        """
        Recebe os dados de todos os lotes, processa e atualiza a UI.
        """
        # 1. Configurar o portfólio com os dados coletados
        self.portfolio.definir_configuracoes_gerais(horas_mes)
        self.portfolio.definir_dados_lotes(lotes_data)

        # 2. Realizar os cálculos e a validação
        df_resultado, _ = self.portfolio.calcular_resumo_consolidado()
        alertas = self.portfolio.validar_alocacao_mensal()

        # 3. Chamar o método da UI para exibir os resultados
        self.ui.atualizar_dashboard(df_resultado, alertas)

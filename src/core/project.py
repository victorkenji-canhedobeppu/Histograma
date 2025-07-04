# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta


class Portfolio:
    """
    Gerencia um portfólio de múltiplos projetos (Lotes), calcula a alocação
    de horas e valida a capacidade dos recursos entre os projetos.
    """

    def __init__(self, disciplinas, classes_func):
        self.disciplinas = disciplinas
        self.classes_func = classes_func
        self.lotes_data = []
        self.horas_por_funcionario_mes = 160  # Valor Padrão

    def definir_configuracoes_gerais(self, horas_por_funcionario):
        """Define as horas base para um funcionário em um mês, vindo da UI."""
        self.horas_por_funcionario_mes = horas_por_funcionario

    def definir_dados_lotes(self, lotes_data):
        """Armazena os dados de todos os lotes (cronograma e alocação)."""
        self.lotes_data = lotes_data

    def calcular_resumo_consolidado(self):
        """
        Processa os dados de todos os lotes e retorna um DataFrame consolidado.
        """
        dados_consolidados = []
        total_horas_portfolio = 0

        for lote in self.lotes_data:
            for disc in self.disciplinas:
                if disc not in lote["alocacao"]:
                    continue

                alocacao_disciplina = lote["alocacao"][disc]
                total_alocacao_percentual = sum(alocacao_disciplina.values())
                horas_disciplina = (
                    total_alocacao_percentual * self.horas_por_funcionario_mes
                )
                total_horas_portfolio += horas_disciplina

                linha = {
                    "Lote": lote["nome"],
                    "Disciplina": disc,
                    "Data Início": lote["cronograma"][disc]["inicio"],
                    "Data Fim": lote["cronograma"][disc]["fim"],
                }
                linha.update(alocacao_disciplina)
                linha["Alocação Total (%)"] = round(total_alocacao_percentual * 100, 2)
                linha["Horas na Disciplina"] = round(horas_disciplina, 2)
                dados_consolidados.append(linha)

        df_resumo = pd.DataFrame(dados_consolidados)

        colunas_ordem = (
            ["Lote", "Disciplina", "Data Início", "Data Fim"]
            + self.classes_func
            + ["Alocação Total (%)", "Horas na Disciplina"]
        )

        # Garante que todas as colunas existam antes de reordenar
        df_resumo = df_resumo.reindex(columns=colunas_ordem, fill_value=0)

        return df_resumo, total_horas_portfolio

    def validar_alocacao_mensal(self):
        """
        Valida a alocação de cada classe de funcionário em todos os lotes, mês a mês.
        Retorna uma lista de alertas se a alocação exceder 1.0 (100%).
        """
        alertas = []
        if not self.lotes_data:
            return alertas

        # Encontrar o intervalo de datas de todo o portfólio
        try:
            datas_inicio = [
                datetime.strptime(lote["cronograma"][disc]["inicio"], "%d/%m/%Y")
                for lote in self.lotes_data
                for disc in self.disciplinas
                if lote["cronograma"][disc]["inicio"]
            ]
            datas_fim = [
                datetime.strptime(lote["cronograma"][disc]["fim"], "%d/%m/%Y")
                for lote in self.lotes_data
                for disc in self.disciplinas
                if lote["cronograma"][disc]["fim"]
            ]
        except ValueError:
            return ["Erro: Formato de data inválido. Use DD/MM/AAAA."]

        if not datas_inicio or not datas_fim:
            return []  # Sem datas para validar

        data_corrente = min(datas_inicio).replace(day=1)
        data_final = max(datas_fim)

        # Iterar mês a mês
        while data_corrente <= data_final:
            mes_ano_str = data_corrente.strftime("%B/%Y")
            alocacao_mensal_por_classe = {classe: 0.0 for classe in self.classes_func}

            # Verificar cada lote
            for lote in self.lotes_data:
                for disc in self.disciplinas:
                    try:
                        inicio_disc = datetime.strptime(
                            lote["cronograma"][disc]["inicio"], "%d/%m/%Y"
                        )
                        fim_disc = datetime.strptime(
                            lote["cronograma"][disc]["fim"], "%d/%m/%Y"
                        )

                        # Checa se a disciplina está ativa neste mês
                        if inicio_disc <= data_corrente <= fim_disc or (
                            inicio_disc.year == data_corrente.year
                            and inicio_disc.month == data_corrente.month
                        ):
                            for classe, valor in lote["alocacao"][disc].items():
                                alocacao_mensal_por_classe[classe] += valor
                    except (ValueError, KeyError):
                        continue  # Ignora se a data ou alocação não estiver preenchida

            # Validar os totais do mês
            for classe, total_alocado in alocacao_mensal_por_classe.items():
                if total_alocado > 1.0:
                    alerta = (
                        f"Alerta em {mes_ano_str}: '{classe}' com alocação "
                        f"de {total_alocado*100:.0f}% (Limite é 100%)."
                    )
                    alertas.append(alerta)

            data_corrente += relativedelta(months=1)

        return alertas

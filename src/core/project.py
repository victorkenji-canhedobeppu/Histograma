# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from collections import defaultdict


class Portfolio:
    def __init__(self, disciplinas):
        self.disciplinas = disciplinas
        self.lotes_data = []
        self.funcionarios = []
        self.horas_trabalhaveis_mes = 160

    def definir_configuracoes_gerais(self, horas_por_funcionario, funcionarios):
        self.horas_trabalhaveis_mes = horas_por_funcionario
        self.funcionarios = funcionarios

    def definir_dados_lotes(self, lotes_data):
        self.lotes_data = lotes_data

    def _processar_alocacao(self, inicio_str, fim_str, horas):
        if not inicio_str or not fim_str or not horas > 0:
            return None
        inicio = datetime.strptime(inicio_str, "%d/%m/%Y")
        fim = datetime.strptime(fim_str, "%d/%m/%Y")
        num_meses = (fim.year - inicio.year) * 12 + (fim.month - inicio.month) + 1
        if num_meses <= 0 or self.horas_trabalhaveis_mes == 0:
            return None
        decimal_mensal = (horas / self.horas_trabalhaveis_mes) / num_meses
        decimais = {}
        data_corrente = inicio.replace(day=1)
        while data_corrente <= fim:
            mes_key = data_corrente.strftime("%Y-%m")
            decimais[mes_key] = decimal_mensal
            data_corrente += relativedelta(months=1)
        return decimais

    def gerar_relatorio_alocacao_decimal(self, nome_lote=None):
        if not self.lotes_data:
            return pd.DataFrame()

        decimal_bruto_geral = defaultdict(lambda: defaultdict(float))

        lotes_a_processar = self.lotes_data
        if nome_lote:
            lotes_a_processar = [
                lote for lote in self.lotes_data if lote["nome"] == nome_lote
            ]

        for lote in lotes_a_processar:
            for disc, dados_disc in lote.get("disciplinas", {}).items():
                cronograma = dados_disc.get("cronograma", {})
                for alocacao in dados_disc.get("alocacoes", []):
                    try:
                        pessoa_tuplo = alocacao["funcionario"]
                        chave_agregacao = (disc, pessoa_tuplo[1], pessoa_tuplo[0])
                        decimais = self._processar_alocacao(
                            cronograma.get("inicio"),
                            cronograma.get("fim"),
                            alocacao["horas_totais"],
                        )
                        if decimais:
                            for mes, valor in decimais.items():
                                decimal_bruto_geral[chave_agregacao][mes] += valor
                    except (ValueError, KeyError, ZeroDivisionError):
                        continue

            for sub_aloc in lote.get("subcontratos", []):
                try:
                    nome_sub = sub_aloc["nome"]
                    chave_agregacao = ("Subcontrato", nome_sub, "")
                    decimais = self._processar_alocacao(
                        sub_aloc.get("inicio"),
                        sub_aloc.get("fim"),
                        sub_aloc["horas_totais"],
                    )
                    if decimais:
                        for mes, valor in decimais.items():
                            decimal_bruto_geral[chave_agregacao][mes] += valor
                except (ValueError, KeyError, ZeroDivisionError):
                    continue

        total_decimal_por_pessoa = defaultdict(lambda: defaultdict(float))
        for (disc, cargo, nome), meses in decimal_bruto_geral.items():
            if nome:
                for mes, valor in meses.items():
                    total_decimal_por_pessoa[(nome, cargo)][mes] += valor

        todos_meses_keys = sorted(
            list(
                set(
                    mes
                    for meses in decimal_bruto_geral.values()
                    for mes in meses.keys()
                )
            )
        )

        relatorio_final = []
        for (disc, cargo_ou_sub, nome), meses_data in decimal_bruto_geral.items():
            linha = {
                "Disciplina": disc,
                "Cargo/Subcontrato": cargo_ou_sub,
                "Funcionário": nome,
            }
            is_excedido = any(
                total_decimal_por_pessoa.get((nome, cargo_ou_sub), {}).get(m, 0.0) > 1.0
                for m in todos_meses_keys
            )

            for mes in todos_meses_keys:
                linha[mes] = meses_data.get(mes, 0.0)

            linha["Status"] = "Alocação Excedida" if is_excedido else "OK"
            relatorio_final.append(linha)

        df = pd.DataFrame(relatorio_final)
        if df.empty:
            return pd.DataFrame()

        colunas_meses_numericas = [key for key in todos_meses_keys if key in df.columns]
        if colunas_meses_numericas:
            df["H.mês"] = df[colunas_meses_numericas].sum(axis=1)

        meses_display_map = {
            key: datetime.strptime(key, "%Y-%m").strftime("%b/%y")
            for key in todos_meses_keys
        }
        df.rename(columns=meses_display_map, inplace=True)

        colunas_para_formatar = list(meses_display_map.values())
        if "H.mês" in df.columns:
            colunas_para_formatar.append("H.mês")

        for col in colunas_para_formatar:
            if col in df.columns:
                df[col] = df[col].apply(
                    lambda x: f"{x:.2f}".replace(".", ",") if pd.notna(x) else "0,00"
                )

        colunas_finais = ["Disciplina", "Cargo/Subcontrato", "Funcionário"] + list(
            meses_display_map.values()
        )
        if "H.mês" in df.columns:
            colunas_finais.append("H.mês")
        colunas_finais.append("Status")

        return df.reindex(columns=colunas_finais).fillna("")

    def gerar_relatorio_detalhado_por_tarefa(self):
        return {}

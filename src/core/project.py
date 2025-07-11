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

    def definir_configuracoes_gerais(self, horas_por_funcionario, funcionarios, cargos):
        self.horas_trabalhaveis_mes = horas_por_funcionario
        self.funcionarios = funcionarios
        self.cargos_disponiveis = cargos

    def definir_dados_lotes(self, lotes_data):
        self.lotes_data = lotes_data

    def gerar_relatorio_horas_por_funcionario(self):
        # ... (código inalterado) ...
        if not self.lotes_data:
            return pd.DataFrame()
        horas_agregadas = defaultdict(float)
        for lote in self.lotes_data:
            for disciplina_nome, dados_disc in lote.get("disciplinas", {}).items():
                for alocacao in dados_disc.get("alocacoes", []):
                    try:
                        nome_func, cargo_func = alocacao["funcionario"]
                        horas = alocacao.get("horas_totais", 0)
                        chave = (nome_func, cargo_func, disciplina_nome)
                        horas_agregadas[chave] += horas
                    except (KeyError, ValueError):
                        continue
        if not horas_agregadas:
            return pd.DataFrame()
        lista_relatorio = [
            {
                "Funcionário": nome,
                "Cargo": cargo,
                "Disciplina": disciplina,
                "Horas Totais Alocadas": horas,
            }
            for (nome, cargo, disciplina), horas in horas_agregadas.items()
        ]
        return pd.DataFrame(lista_relatorio).sort_values(
            by=["Funcionário", "Disciplina"]
        )

    def _processar_alocacao(self, inicio_str, fim_str, horas):
        # ... (código inalterado) ...
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
                        # --- MÉTODO REVISADO: Validação do funcionário ---
                        nome_func, cargo_func = alocacao["funcionario"]

                        # Um funcionário é válido se for um cargo (alocação default) ou se estiver na lista de funcionários
                        if nome_func in self.cargos_disponiveis:
                            employee_exists = True
                        else:
                            employee_exists = any(
                                f[0] == nome_func for f in self.funcionarios
                            )

                        if not employee_exists:
                            continue

                        chave_agregacao = (disc, cargo_func, nome_func)
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

        # ... (restante do método inalterado) ...
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
            linha = {"Disciplina": disc, "Cargo": cargo_ou_sub, "Funcionário": nome}
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
        colunas_finais = ["Disciplina", "Cargo", "Funcionário"] + list(
            meses_display_map.values()
        )
        if "H.mês" in df.columns:
            colunas_finais.append("H.mês")
        colunas_finais.append("Status")
        return df.reindex(columns=colunas_finais).fillna("")

    def gerar_relatorio_detalhado_por_tarefa(self):
        return {}

    def verificar_porcentagem_horas_cargo(self, cargo_alvo):
        if not self.lotes_data:
            return 0.0, {}

        horas_cargo_alvo_global = 0.0
        horas_totais_projeto_global = 0.0
        detalhes_por_lote = {}

        for lote in self.lotes_data:
            nome_lote = lote.get("nome", "Desconhecido")
            horas_cargo_alvo_lote = 0.0
            horas_totais_lote = 0.0

            for disc, dados_disc in lote.get("disciplinas", {}).items():
                for alocacao in dados_disc.get("alocacoes", []):
                    try:
                        horas = alocacao.get("horas_totais", 0)
                        horas_totais_lote += horas

                        _, cargo_func = alocacao["funcionario"]
                        if cargo_func == cargo_alvo:
                            horas_cargo_alvo_lote += horas
                    except (KeyError, ValueError):
                        continue

            for sub_aloc in lote.get("subcontratos", []):
                try:
                    horas = sub_aloc.get("horas_totais", 0)
                    horas_totais_lote += horas
                except (KeyError, ValueError):
                    continue

            if horas_totais_lote > 0:
                percent_lote = (horas_cargo_alvo_lote / horas_totais_lote) * 100
                detalhes_por_lote[nome_lote] = percent_lote
            else:
                detalhes_por_lote[nome_lote] = 0.0

            horas_cargo_alvo_global += horas_cargo_alvo_lote
            horas_totais_projeto_global += horas_totais_lote

        if horas_totais_projeto_global == 0:
            porcentagem_global = 0.0
        else:
            porcentagem_global = (
                horas_cargo_alvo_global / horas_totais_projeto_global
            ) * 100

        return porcentagem_global, detalhes_por_lote

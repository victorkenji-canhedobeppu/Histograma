# -*- coding: utf-8 -*-


class TerminalUI:
    """
    Responsável por toda a interação com o usuário no terminal (entradas e saídas).
    """

    @staticmethod
    def exibir_cabecalho():
        """Exibe o cabeçalho do programa."""
        print("=" * 60)
        print("   PAINEL DE PLANEJAMENTO DE PROJETO (Python)   ")
        print("=" * 60)
        print("Insira os dados para calcular a alocação de horas.\n")

    @staticmethod
    def obter_dados_gerais():
        """Coleta os dados gerais do projeto do usuário."""
        try:
            limite_horas = int(input("Digite o limite de horas totais por mês: "))
            qtd_lotes = int(input("Digite a quantidade de lotes do projeto: "))
            return limite_horas, qtd_lotes
        except ValueError:
            print("\nErro: Por favor, insira um número inteiro válido.")
            return TerminalUI.obter_dados_gerais()

    @staticmethod
    def obter_alocacao_equipe(disciplinas, classes_func):
        """Coleta a quantidade de funcionários por classe para cada disciplina."""
        print("\n--- Alocação de Equipe por Disciplina ---")
        alocacao = {}
        for disc in disciplinas:
            print(f"\nDisciplina: {disc}")
            alocacao[disc] = {}
            for classe in classes_func:
                while True:
                    try:
                        qtd = int(input(f"  - Qtd. {classe}: "))
                        if qtd < 0:
                            print("  Erro: A quantidade não pode ser negativa.")
                            continue
                        alocacao[disc][classe] = qtd
                        break
                    except ValueError:
                        print("  Erro: Por favor, insira um número inteiro.")
        return alocacao

    @staticmethod
    def exibir_dashboard(df_resumo, total_horas, limite_horas):
        """Exibe o dashboard de resumo final com os resultados."""
        print("\n" + "=" * 80)
        print(" " * 30 + "DASHBOARD DE RESUMO")
        print("=" * 80)

        print("\n--- Tabela de Alocação e Horas Calculadas ---\n")
        print(df_resumo.to_string(index=False))

        print("\n" + "-" * 80)

        print(f"\nTotal de Horas Calculadas: {total_horas} h")
        print(f"Limite de Horas Mensais:   {limite_horas} h")

        if total_horas > limite_horas:
            print("\nStatus: \033[91mHORAS EXCEDIDAS\033[0m")  # Vermelho
        else:
            print("\nStatus: \033[92mDENTRO DO LIMITE\033[0m")  # Verde

        print("\n" + "=" * 80)

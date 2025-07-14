# -*- coding: utf-8 -*-

from config.settings import (
    CARGOS,
    DISCIPLINAS,
    MAPEAMENTO_TAREFA_DISCIPLINA,
    OUTROS,
    SUBCONTRATOS,
)
from core.project import Portfolio
from ui.gui_app import GuiApp
import psutil
from cryptography.fernet import Fernet
import base64


class App:
    def __init__(self):
        self.mac_address = self.get_mac_address()
        try:
            # Tenta ler e descriptografar a chave do arquivo
            self.chave_atual = self.descriptografar_chave()
        except Exception as e:
            # Se o arquivo não existir ou der erro, usa um valor que garantirá a falha
            print(
                f"AVISO: Não foi possível ler ou descriptografar a chave: {e}. Usando fallback."
            )
            self.chave_atual = "ERRO-NA-CHAVE"

        print(f"Endereço MAC da Máquina: {self.mac_address}")
        print(f"Chave da Licença: {self.chave_atual}")

        # Compara o MAC da máquina com a chave da licença
        self.mac_autorizado = self.mac_address == self.chave_atual
        if not self.mac_autorizado:
            print(
                "AVISO: Endereço MAC não autorizado. Funcionalidades serão desabilitadas."
            )
        else:
            print("INFO: Licença validada com sucesso.")
        self.funcionarios = []
        self.portfolio = Portfolio(list(DISCIPLINAS.values()))
        self.ui = GuiApp(app_controller=self)
        self.ultimo_df_consolidado = None
        self.ultimo_dashboards_lotes = None

        self.ui.set_modo_restringido(not self.mac_autorizado)

    def get_mac_address(self):
        """Obtém o endereço MAC da primeira interface de rede física."""
        try:
            # itera sobre todas as interfaces de rede
            for interface, snics in psutil.net_if_addrs().items():
                for snic in snics:
                    # AF_LINK é o endereço MAC
                    if snic.family == psutil.AF_LINK:
                        # Retorna o primeiro endereço MAC encontrado que não seja de loopback
                        if snic.address and snic.address != "00:00:00:00:00:00":
                            # Formata para o padrão com hífens e maiúsculas
                            return snic.address.replace(":", "-").upper()
            return "Endereço MAC não encontrado"
        except Exception as e:
            return f"Erro ao obter endereço MAC: {e}"

    def descriptografar_chave(self):
        """Lê o arquivo de licença e descriptografa o endereço MAC contido nele."""
        with open("chave.cbel", "rb") as file:
            conteudo = base64.urlsafe_b64decode(file.read())

        # Assume que o formato é chave_secreta||mensagem_criptografada
        chave_secreta, mensagem_criptografada = conteudo.split(b"||")

        fernet = Fernet(chave_secreta)
        resultado_bytes = fernet.decrypt(mensagem_criptografada)

        # Retorna o resultado decodificado e formatado
        return resultado_bytes.decode("utf-8").strip().upper()

    # Em app.py, dentro da classe App

    def get_funcionarios_para_tarefa(self, nome_da_tarefa):
        """
        Retorna uma lista de funcionários qualificados para uma tarefa, usando
        um mapeamento BIDIRECIONAL.
        """
        # 1. Começa com um conjunto (set) que inclui a própria tarefa como uma disciplina válida.
        disciplinas_relacionadas = {nome_da_tarefa}

        # 2. Itera sobre o mapa para encontrar todas as relações (nos dois sentidos).
        for tarefa_mapeada, disciplinas_fonte in MAPEAMENTO_TAREFA_DISCIPLINA.items():
            # Normaliza os nomes para comparação insensível a maiúsculas/minúsculas
            tarefa_mapeada_lower = tarefa_mapeada.lower()
            disciplinas_fonte_lower = [d.lower() for d in disciplinas_fonte]
            nome_da_tarefa_lower = nome_da_tarefa.lower()

            # LÓGICA BIDIRECIONAL:
            # a) Se a tarefa atual é a CHAVE do mapa, adicione as disciplinas da DIREITA.
            if tarefa_mapeada_lower == nome_da_tarefa_lower:
                disciplinas_relacionadas.update(disciplinas_fonte)

            # b) Se a tarefa atual está na LISTA da direita, adicione a CHAVE da esquerda.
            if nome_da_tarefa_lower in disciplinas_fonte_lower:
                disciplinas_relacionadas.add(tarefa_mapeada)

        # 3. Agora que temos todas as disciplinas relacionadas, coletamos os funcionários.
        funcionarios_qualificados = set()
        for disciplina in disciplinas_relacionadas:
            funcionarios_encontrados = self.get_nomes_funcionarios_por_disciplina(
                disciplina
            )
            for func in funcionarios_encontrados:
                funcionarios_qualificados.add(func)

        return sorted(list(funcionarios_qualificados))

    def get_resumo_equipe_df(self):
        """Garante que os dados mais recentes da UI estão no portfólio e retorna o resumo."""
        dados_ui = self.ui._coletar_dados_da_ui()
        if not dados_ui:
            return None
        self.portfolio.definir_dados_lotes(dados_ui["lotes"])
        return self.portfolio.gerar_relatorio_horas_por_funcionario()

    def run(self):
        self.ui.carregar_ou_iniciar_ui()
        self.ui.mainloop()

    def get_disciplinas(self):
        return sorted(list(DISCIPLINAS.values()))

    def get_cargos_disponiveis(self):
        return sorted(list(CARGOS.values()), reverse=True)

    def get_tarefas_adicionais_disponiveis(self):
        """Combina as listas de SUBCONTRATOS e OUTROS em uma única lista de tarefas adicionáveis."""
        tarefas = list(SUBCONTRATOS.values()) + list(OUTROS.values())
        return sorted(tarefas)

    def get_ultimo_df_consolidado(self):
        return self.ultimo_df_consolidado

    def get_ultimo_dashboards_lotes(self):
        return self.ultimo_dashboards_lotes

    def get_funcionarios_para_display(self):
        """Retorna a string formatada para a lista de gerenciamento de equipe."""
        return sorted(
            [f"{nome} [{disciplina}]" for nome, disciplina in self.funcionarios]
        )

    def get_nomes_funcionarios_por_disciplina(self, disciplina_filtro):
        """Retorna uma lista de NOMES de funcionários para uma dada disciplina."""
        return sorted(
            [
                nome
                for nome, disciplina in self.funcionarios
                if disciplina == disciplina_filtro
            ]
        )

    def get_tarefas_ativas(self):
        """
        Retorna uma lista unificada de todas as disciplinas padrão mais quaisquer
        subcontratos que já foram adicionados a algum lote na interface.
        """
        # Começa com um conjunto (set) para evitar duplicatas, já preenchido com as disciplinas padrão.
        tarefas_ativas = set(self.get_disciplinas())

        # Itera sobre os widgets da UI para encontrar tarefas (subcontratos) adicionadas
        if self.ui and self.ui.lote_widgets:
            for lote_data in self.ui.lote_widgets.values():
                # A chave "disciplinas" contém todas as tarefas do lote
                for nome_da_tarefa in lote_data.get("disciplinas", {}).keys():
                    tarefas_ativas.add(nome_da_tarefa)

        # Retorna como uma lista ordenada
        return sorted(list(tarefas_ativas))

    def get_todas_as_tarefas(self):
        """Retorna uma lista unificada de todas as disciplinas e subcontratos."""
        disciplinas = self.get_disciplinas()
        subcontratos = self.get_subcontratos_disponiveis()
        return sorted(disciplinas + subcontratos)

    def get_tarefas_ja_adicionadas(self, lote_nome):
        """Retorna as tarefas que já existem na UI para um determinado lote."""
        if lote_nome in self.ui.lote_widgets:
            return list(self.ui.lote_widgets[lote_nome]["disciplinas"].keys())
        return []

    def adicionar_funcionario(self, nome, disciplina):
        nome_tratado = nome.strip()
        if not nome_tratado:
            return

        novo_funcionario = (nome_tratado, disciplina)

        # Verifica se um funcionário com o mesmo nome já existe
        if not any(f[0].lower() == nome_tratado.lower() for f in self.funcionarios):
            self.funcionarios.append(novo_funcionario)
            self.funcionarios.sort(key=lambda x: x[0])
            self.ui.atualizar_lista_funcionarios()
        else:
            print(f"Aviso: Funcionário com nome '{nome_tratado}' já existe.")

    def remover_funcionario(self, display_string):
        """Remove um funcionário baseado na sua string de exibição 'Nome [Disciplina]'."""
        try:
            # Faz o parsing da string "Nome [Disciplina]"
            nome, disciplina = display_string.rsplit(" [", 1)
            disciplina = disciplina.rstrip("]")

            funcionario_a_remover = (nome, disciplina)
            if funcionario_a_remover in self.funcionarios:
                self.funcionarios.remove(funcionario_a_remover)
                self.ui.atualizar_lista_funcionarios()

        except ValueError:
            print(
                f"Erro ao tentar remover: formato de string inválido '{display_string}'"
            )

    def processar_portfolio(self, lotes_data, horas_mes):
        """Processa todos os dados e atualiza a UI com os resultados."""
        self.portfolio.definir_configuracoes_gerais(
            horas_mes, self.funcionarios, self.get_cargos_disponiveis()
        )
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

        porcentagem_global, detalhes_lotes = (
            self.portfolio.verificar_porcentagem_horas_cargo(cargo_alvo=CARGOS["ESTAG"])
        )

        alerta_composicao = {"global": porcentagem_global, "lotes": detalhes_lotes}

        self.ui.atualizar_dashboards(
            self.ultimo_df_consolidado,
            self.ultimo_dashboards_lotes,
            detalhes_tarefas,
            alerta_composicao=alerta_composicao,
        )

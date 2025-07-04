# -*- coding: utf-8 -*-

import sys
import os

# Adiciona o diretório 'src' ao path do Python para encontrar os módulos.
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "src")))

# Importa a classe principal da aplicação do módulo 'app'
from ui.app import App

if __name__ == "__main__":
    """
    Ponto de entrada principal da aplicação.

    Este script instancia a classe App e chama o método run() para iniciar.
    Para executar, navegue até a raiz do projeto no terminal e rode: python main.py
    """
    # Cria uma instância da aplicação
    application = App()

    # Inicia a execução
    application.run()

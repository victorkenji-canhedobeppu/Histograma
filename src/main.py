# -*- coding: utf-8 -*-

import sys
import os

from ui.app import App

# --- Configuração do Path ---
# Adiciona o diretório 'src' ao path do Python.
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "src")))

# --- Importação e Execução ---

if __name__ == "__main__":
    """
    Ponto de entrada principal da aplicação.
    """
    # Cria uma instância da aplicação
    application = App()

    # Inicia a execução
    application.run()

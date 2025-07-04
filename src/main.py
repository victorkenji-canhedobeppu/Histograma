# src/main.py
# Este é o ponto de entrada principal da aplicação.
# Sua única responsabilidade é instanciar e iniciar a interface do usuário.

from ui.app import App
import customtkinter as ctk


def main():
    """
    Função principal que inicializa e executa a aplicação.
    """
    # Define o tema da aplicação
    ctk.set_appearance_mode("System")  # Pode ser "System", "Dark", "Light"
    ctk.set_default_color_theme("blue")  # Tema de cor padrão

    # Cria a instância da janela principal
    app = App()

    # Inicia o loop de eventos da interface gráfica
    app.mainloop()


if __name__ == "__main__":
    main()

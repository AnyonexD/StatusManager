import flet as ft
import logging


class Status_Selection:
    @staticmethod
    def dropdown():
        return ft.Dropdown(   # <-- agora ele devolve o componente
            label="Status",
            border_color=ft.Colors.WHITE,
            label_style=ft.TextStyle(color=ft.Colors.WHITE, size=12),
            bgcolor="#1b1b1b",
            width=220,
            text_style=ft.TextStyle(color=ft.Colors.WHITE, size=12),
            options=[
                ft.dropdown.Option(
                    text="Disponível",
                    content=ft.Row(
                        controls=[
                            ft.Icon(ft.Icons.CHECK, color=ft.Colors.GREEN),
                            ft.Text("Disponível", color=ft.Colors.WHITE, size=12)
                        ],
                        alignment=ft.MainAxisAlignment.START
                    )
                ),
                ft.dropdown.Option(
                    text="Defeito/Inutilizado",
                    content=ft.Row(
                        controls=[
                            ft.Icon(ft.Icons.CLOSE, color=ft.Colors.RED),
                            ft.Text("Defeito/Inutilizado", color=ft.Colors.WHITE, size=12)
                        ],
                        alignment=ft.MainAxisAlignment.START
                    )
                ),
                ft.dropdown.Option(
                    text="Procurar",
                    content=ft.Row(
                        controls=[
                            ft.Icon(ft.Icons.SEARCH, color=ft.Colors.BLUE),
                            ft.Text("Procurar", color=ft.Colors.WHITE, size=12)
                        ],
                        alignment=ft.MainAxisAlignment.START
                    )
                )
            ]
        )


class Login_and_pass():
    @staticmethod
    def login():
        return ft.TextField(
        label='Login',
        border_color=ft.Colors.WHITE,
        label_style=ft.TextStyle(color=ft.Colors.WHITE, size=12),
        color=ft.Colors.WHITE,
        prefix_icon=ft.Icons.PERSON,
        width=280
    )
    
    def password():
        return ft.TextField(
            label='Senha',
            border_color=ft.Colors.WHITE,
            label_style=ft.TextStyle(color=ft.Colors.WHITE, size=12),
            color=ft.Colors.WHITE,
            password=True,
            can_reveal_password=True,
            prefix_icon=ft.Icons.LOCK,
            width=280
        )
    
class Logger_Manager():
    def __init__(self, log_file="meu_app.log"):
        # Cria logger
        self.logger = logging.getLogger("MeuLogger")
        self.logger.setLevel(logging.INFO)

        # Evita adicionar handlers duplicados
        if not self.logger.hasHandlers():
            # Formato do log
            formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

            # Log em arquivo
            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(formatter)

            # Log no console
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)

            # Adiciona os handlers
            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)

    def info(self, message):
        self.logger.info(message)

    def warning(self, message):
        self.logger.warning(message)

    def error(self, message):
        self.logger.error(message)

    def debug(self, message):
        self.logger.debug(message)
import flet as ft
import logging
import os


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
    
class Logger_Manager:
    def __init__(self, log_file="SM.log"):
        # Cria logger
        self.logger = logging.getLogger("MeuLogger")
        self.logger.setLevel(logging.INFO)
        
        # Evita adicionar handlers duplicados
        if not self.logger.hasHandlers():
            # Determina o caminho seguro para o arquivo de log
            arquivo_log_seguro = self._obter_caminho_seguro_log(log_file)
            
            # Formato do log
            formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
            
            try:
                # Log em arquivo
                file_handler = logging.FileHandler(arquivo_log_seguro, encoding='utf-8')
                file_handler.setFormatter(formatter)
                self.logger.addHandler(file_handler)
            except (PermissionError, OSError) as e:
                # Se falhar o log em arquivo, continua apenas com console
                console_handler = logging.StreamHandler()
                console_handler.setFormatter(formatter)
                self.logger.addHandler(console_handler)
                self.logger.error(f"Não foi possível criar arquivo de log {arquivo_log_seguro}: {e}")
                self.logger.info("Continuando apenas com log no console")
                return
            
            # Log no console
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)
            
            self.logger.info(f"Logger inicializado. Arquivo de log: {arquivo_log_seguro}")
    
    def _obter_caminho_seguro_log(self, nome_arquivo_log):
        """Encontra um diretório seguro para criar o arquivo de log"""
        userdir = os.path.expanduser("~")
        
        # Tenta diferentes localizações em ordem de preferência
        caminhos_possiveis = [
            # Diretório do usuário
            userdir,
            # Desktop
            os.path.join(userdir, "Desktop"),
            os.path.join(userdir, "Área de Trabalho"),
            # Localizações do OneDrive
            os.path.join(userdir, "OneDrive - Desktop SA", "Área de Trabalho", "Robos"),
            os.path.join(userdir, "OneDrive - Desktop Sigmanet", "Área de Trabalho", "Robos"),
            os.path.join(userdir, "OneDrive", "Área de Trabalho", "Robos"),
            # Diretório temporário como último recurso
            os.environ.get('TEMP', userdir)
        ]
        
        for caminho in caminhos_possiveis:
            try:
                # Testa se consegue escrever neste diretório
                if os.path.exists(caminho):
                    arquivo_teste = os.path.join(caminho, ".teste_escrita")
                    with open(arquivo_teste, 'w') as f:
                        f.write("teste")
                    os.remove(arquivo_teste)
                    return os.path.join(caminho, nome_arquivo_log)
                elif caminho != os.environ.get('TEMP', userdir):
                    # Tenta criar diretório (exceto para temp)
                    os.makedirs(caminho, exist_ok=True)
                    return os.path.join(caminho, nome_arquivo_log)
            except (OSError, PermissionError):
                continue
        
        # Último recurso - usa diretório temporário
        import tempfile
        return os.path.join(tempfile.gettempdir(), nome_arquivo_log)
    
    def info(self, message):
        self.logger.info(message)
    
    def warning(self, message):
        self.logger.warning(message)
    
    def error(self, message):
        self.logger.error(message)
    
    def debug(self, message):
        self.logger.debug(message)

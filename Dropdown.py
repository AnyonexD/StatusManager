import flet as ft


class Status_Selection:
        def dropdown():
            status_dropdown = ft.Dropdown(
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
                # ft.dropdown.Option(
                #     text="Baixado/Perdido",
                #     content=ft.Row(
                #         controls=[
                #             ft.Icon(ft.Icons.DELETE, color=ft.Colors.GREY),
                #             ft.Text("Baixado/Perdido", color=ft.Colors.WHITE, size=12)
                #         ],
                #         alignment=ft.MainAxisAlignment.START
                #     )
                # ),

                
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
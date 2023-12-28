import flet as ft
import coelhos

AVISO_PRINCIPAL = "Selecione uma planilha para começar!"
AVISO_CARREGANDO = "Carregando planilha..."
AVISO_SUCESSO = "Planilha carregada! Salve o resultado! :D"
AVISO_PAGINA = "Selecione uma página!"
AVISO_ERRO = "Erro ao carregar planilha! Verifique se o arquivo está correto."
AVISO_ERRO_CARREGAR = "Erro ao carregar planilha! Verifique se o arquivo está correto."
AVISO_SALVANDO = "Salvando planilha..."
AVISO_SALVO = "Planilha salva com sucesso!"

def main(page: ft.page):
    ###### Page Config ######
    page.title = "Grau de Parentesco de Coelhos"
    page.window_width = 600
    page.window_height = 800


    ###### File Input ######
    def pick_file_result(e: ft.FilePickerResultEvent):
        global filename

        set_aviso(AVISO_CARREGANDO)

        try:
            filename = e.files[0].name
            selected_filename.value = filename
            selected_filename.update()
            print('Filename:', filename)
        except:
            set_aviso(AVISO_ERRO_CARREGAR)

        clear_input_table()
        clear_output_table()

        if filename.endswith('.xlsx'):
            sheets = coelhos.get_sheets(filename)
            print('Sheets:', sheets)

            dropdown_sheets.options = [ft.dropdown.Option(sheet) for sheet in sheets]
            dropdown_sheets.disabled = False
            dropdown_sheets.update()

            set_aviso(AVISO_PAGINA)
        else:
            dropdown_sheets.options = []
            dropdown_sheets.disabled = True
            dropdown_sheets.update()
            set_aviso(AVISO_ERRO_CARREGAR)


    filepick_input = ft.FilePicker(on_result=pick_file_result)
    page.overlay.append(filepick_input)

    insert_button = ft.ElevatedButton(
        "Inserir Planilha", 
        icon=ft.icons.UPLOAD_FILE,
        on_click=lambda _: filepick_input.pick_files(allow_multiple=False)
    )
    selected_filename = ft.Text()

    dropdown_sheets = ft.Dropdown(
        width=130,
        options=[],
        disabled=True,
        on_change=lambda e: change_tables(filename, e.data)
    )

    def change_tables(filename, sheet):
        global gp

        clear_input_table()
        clear_output_table()

        try:
            gp = coelhos.GrauParentesco(filename, sheet)
            change_input_table(gp.get_data_input)
            change_output_table(gp.get_data_output)
        except Exception as e:
            set_aviso(AVISO_ERRO)
            raise e
    ###### Input Table ######

    def change_input_table(data):
        for row in data:
            input_table.rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(row['individuo'])),
                        ft.DataCell(ft.Text(row['pai'])),
                        ft.DataCell(ft.Text(row['mae'])),
                        ft.DataCell(ft.Text(row['sexo'])),
                    ]
                )
            )
        input_table.update()
    
    def clear_input_table():
        input_table.rows = []
        input_table.update()

    input_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("Indivíduo")),
            ft.DataColumn(ft.Text("Pai")),
            ft.DataColumn(ft.Text("Mãe")),
            ft.DataColumn(ft.Text("Sexo"))
        ],
        rows=[
        ],
        show_bottom_border=True,
    )
    lv_input = ft.ListView(height=250, spacing=10, padding=20, auto_scroll=False)
    lv_input.controls.append(input_table)

    text_input = ft.Text("COELHOS")


    ###### Output Table ######

    def change_output_table(data):
        for row in data:
            output_table.rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(row['origem'])),
                        ft.DataCell(ft.Text(row['destino'])),
                        ft.DataCell(ft.Text(row['parentesco'])),
                        ft.DataCell(ft.Text(row['coeficiente'])),
                    ]
                )
            )

        output_table.update()

        save_button.disabled = False
        save_button.update()

        set_aviso(AVISO_SUCESSO)

    def clear_output_table():
        output_table.rows = []
        output_table.update()

        save_button.disabled = True
        save_button.update()

    output_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("Origem")),
            ft.DataColumn(ft.Text("Destino")),
            ft.DataColumn(ft.Text("Parentesco")),
            ft.DataColumn(ft.Text("Coeficiente")),
        ],
        rows=[
        ],
        show_bottom_border=True,
    )
    lv_output = ft.ListView(height=250, spacing=10, padding=20, auto_scroll=False)
    lv_output.controls.append(output_table)

    text_output = ft.Text("PARENTESCOS")

    ###### Save Button ######

    def save_output(e: ft.FilePickerResultEvent):
        status_text.value = AVISO_SALVANDO
        status_text.update()

        output_filename = e.path

        gp.copy_sheet(output_filename)
        gp.salvar_coeficientes()
        gp.salvar_coeficientes_detalhados()

        set_aviso(AVISO_SALVO)

    filepicker_save = ft.FilePicker(on_result=save_output)
    page.overlay.append(filepicker_save)
    save_button = ft.ElevatedButton("Salvar", icon=ft.icons.SAVE, disabled=True, on_click=lambda _: filepicker_save.save_file(allowed_extensions=['xlsx']))

    ###### Aviso ######

    def set_aviso(text):
        status_text.value = text
        status_text.update()
        if text == AVISO_SALVO:
            status_container.bgcolor = ft.colors.GREEN_200
        elif 'erro' in text.lower():
            status_container.bgcolor = ft.colors.RED_200
        else:
            status_container.bgcolor = ft.colors.AMBER
        status_container.update()

    status_text = ft.Text(AVISO_PRINCIPAL, color=ft.colors.BLACK)
    status_container = ft.Container(
        content=status_text,
        bgcolor=ft.colors.AMBER,
        border_radius=8,
        padding=10,
    )

    ###### layout ######

    controls = ft.Column([
        ft.Row([insert_button, selected_filename, dropdown_sheets], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        ft.Divider(),
        ft.Column([text_input, lv_input], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
        ft.Divider(),
        ft.Column([text_output, lv_output], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
        ft.Row([save_button], alignment=ft.MainAxisAlignment.CENTER),
        ft.Row([status_container], alignment=ft.MainAxisAlignment.CENTER),
    ])

    page.add(controls)

ft.app(target=main, assets_dir='assets')
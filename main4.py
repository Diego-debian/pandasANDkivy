import os
import pandas as pd
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.dropdown import DropDown
from kivy.uix.scrollview import ScrollView
from kivy.core.window import Window


class ExcelApp(App):
    title = 'Programa que genera reporte de asistencias'

    def build(self):
        self.file_to_process = None
        self.action = None

        layout = BoxLayout(orientation='vertical', spacing=10, padding=10)

        self.dropdown = DropDown()
        merge_button = Button(text='Unificar hojas', size_hint_y=None, height=40)
        merge_button.bind(on_release=lambda btn: self.dropdown.select(btn.text))
        self.dropdown.add_widget(merge_button)
        check_button = Button(text='Verificar asistencia', size_hint_y=None, height=40)
        check_button.bind(on_release=lambda btn: self.dropdown.select(btn.text))
        self.dropdown.add_widget(check_button)

        action_label = Label(text='Seleccione una acción:', font_size='18sp', halign='center')
        layout.add_widget(action_label)
        main_button = Button(text='Seleccione una acción', size_hint=(1, None), width=200)
        main_button.bind(on_release=self.dropdown.open)
        layout.add_widget(main_button)


        file_label = Label(text='Seleccione un archivo Excel:', font_size='18sp', halign='left')
        layout.add_widget(file_label)

        scroll_view = ScrollView()

        self.file_chooser = FileChooserListView(path=os.getcwd(), size_hint=(1, None), height=400)
        scroll_view.add_widget(self.file_chooser)

        layout.add_widget(scroll_view)

        self.file_label = Label(text='No se ha seleccionado ningún archivo', font_size='14sp', halign='left')
        layout.add_widget(self.file_label)

        select_button = Button(text='Seleccionar archivo', on_release=self.select_file, size_hint=(1, None), height=40)
        layout.add_widget(select_button)

        




        process_button = Button(text='Procesar archivo', on_release=self.process_file, size_hint=(1, None), height=40)
        layout.add_widget(process_button)

        self.result_label = Label(text='', font_size='14sp', halign='left', valign='middle')
        layout.add_widget(self.result_label)

        self.dropdown.bind(on_select=lambda instance, x: self.set_action(x))

        return layout

    def set_action(self, action):
        self.action = action

    def select_file(self, instance):
        if self.file_chooser.selection:
            self.file_to_process = self.file_chooser.selection[0]
            self.file_label.text = f'Archivo Excel seleccionado:\n{os.path.basename(self.file_to_process)}'

    def process_file(self, instance):
        if self.action == 'Unificar hojas':
            self.merge_sheets()
        elif self.action == 'Verificar asistencia':
            self.check_attendance()

    def merge_sheets(self):
        if self.file_to_process:
            try:
                # Leer todas las hojas del archivo Excel
                xls = pd.ExcelFile(self.file_to_process)
                sheet_names = xls.sheet_names

                # Unificar hojas que contienen la palabra "Grupo" en su nombre en un DataFrame
                df_merged = pd.DataFrame()
                for sheet_name in sheet_names:
                    if 'Grupo' in sheet_name:
                        df_sheet = pd.read_excel(self.file_to_process, sheet_name=sheet_name)
                        # Eliminar "basico_apps_grupo_0" de la columna de grupos
                        df_sheet['Grupos'] = df_sheet['Grupos'].str.replace('basico_apps_grupo_0', '')
                        df_merged = pd.concat([df_merged, df_sheet], ignore_index=True)

                # Exportar los datos unificados a un nuevo archivo Excel
                output_file = 'datos_unificados.xlsx'
                df_merged.to_excel(output_file, index=False)

                self.result_label.text = f'Se han unificado los datos en el archivo:\n{output_file}'
            except Exception as e:
                self.result_label.text = f'Error al procesar el archivo:\n{str(e)}'
        else:
            self.result_label.text = 'Seleccione un archivo Excel'

    def check_attendance(self):
        if self.file_to_process:
            try:
                # Leer el archivo Excel y cargar los datos en un DataFrame
                df = pd.read_excel(self.file_to_process)

                # Verificar si las columnas requeridas están presentes
                required_columns = ['Marca temporal', 'Dirección de correo electrónico', 'Nombre completo',
                                    'No. Documento de identidad', 'Teléfono de contacto', 'Nivel', 'Grupo',
                                    'Nombre del formador']
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    self.result_label.text = f'El archivo no contiene las siguientes columnas requeridas:\n{", ".join(missing_columns)}'
                    return

                # Convertir todos los correos a minúsculas
                df['Dirección de correo electrónico'] = df['Dirección de correo electrónico'].str.lower()

                # Obtener el correo electrónico, grupo y fechas de asistencia
                df['Fecha de asistencia'] = df['Marca temporal'].dt.strftime('%d de %B de %Y')
                student_attendance = df.groupby(['Grupo', 'Dirección de correo electrónico', 'Nombre completo'])['Fecha de asistencia'].apply(list).reset_index()
                student_attendance.columns = ['Grupo', 'Correo electrónico', 'Nombre completo', 'Fechas de asistencia']

                # Ordenar el DataFrame por grupo y correo electrónico
                student_attendance.sort_values(['Grupo', 'Correo electrónico'], inplace=True)

                # Eliminar filas con el correo electrónico repetido
                student_attendance.drop_duplicates(subset='Correo electrónico', keep='first', inplace=True)

                # Realizar depuración adicional por nombre
                student_attendance.drop_duplicates(subset='Nombre completo', keep='first', inplace=True)

                # Exportar los resultados a un nuevo archivo Excel
                output_file = 'asistencia_estudiantes.xlsx'
                student_attendance.to_excel(output_file, index=False)

                self.result_label.text = f'Se ha generado el archivo de asistencia:\n{output_file}'
            except Exception as e:
                self.result_label.text = f'Error al procesar el archivo:\n{str(e)}'
        else:
            self.result_label.text = 'Seleccione un archivo Excel'


if __name__ == '__main__':
    Window.size = (600, 700)  # Ajusta el tamaño de la ventana principal
    ExcelApp().run()

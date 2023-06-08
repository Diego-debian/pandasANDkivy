import os
import pandas as pd
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView

class ColumnComparisonApp(App):
    title = 'Programa que genera lista de estudiantes que núnca han ingresado a clases'
    def build(self):
        
        self.file1 = None
        self.file2 = None

        layout = BoxLayout(orientation='vertical')

        self.file1_chooser = FileChooserListView(path=os.getcwd())
        layout.add_widget(Label(text='Seleccione el archivo datos_unificados.xlsx'))
        layout.add_widget(self.file1_chooser)

        self.file2_chooser = FileChooserListView(path=os.getcwd())
        layout.add_widget(Label(text='Seleccione el archivo asistencia_estudiantes.xlsx'))
        layout.add_widget(self.file2_chooser)

        self.result_label = Label(text='')
        layout.add_widget(self.result_label)

        compare_button = Button(text='Generar reporte', on_release=self.compare_columns)
        layout.add_widget(compare_button)

        return layout

    def compare_columns(self, instance):
        file1_selected_file = self.file1_chooser.selection and self.file1_chooser.selection[0]
        file2_selected_file = self.file2_chooser.selection and self.file2_chooser.selection[0]

        if file1_selected_file and file2_selected_file:
            try:
                # Leer los archivos XLSX y cargar los datos en DataFrames
                df1 = pd.read_excel(file1_selected_file, usecols=['Nombre', 'Apellido(s)', 'Dirección de correo', 'Grupos'])
                df2 = pd.read_excel(file2_selected_file, usecols=['Grupo', 'Correo electrónico', 'Nombre completo', 'Fechas de asistencia'])

                # Convertir los correos a minúsculas
                df1['Dirección de correo'] = df1['Dirección de correo'].str.lower()
                df2['Correo electrónico'] = df2['Correo electrónico'].str.lower()

                # Realizar la comparación de columnas
                merged_df = pd.merge(df1, df2, how='left', left_on='Dirección de correo', right_on='Correo electrónico')
                students_not_attended = merged_df[merged_df['Fechas de asistencia'].isna()][['Nombre', 'Dirección de correo', 'Grupos']]
                students_not_attended.columns = ['Nombre', 'Correo', 'Grupo']

                # Exportar los estudiantes sin asistencia a un nuevo archivo XLSX
                output_file = 'estudiantes_sin_asistencia.xlsx'
                students_not_attended.to_excel(output_file, index=False)

                self.result_label.text = f'Se ha generado el archivo de estudiantes sin asistencia: "{output_file}".'

            except Exception as e:
                self.result_label.text = f'Error al procesar los archivos XLSX: {str(e)}'
        else:
            self.result_label.text = 'Seleccione tanto el archivo XLSX de estudiantes como el archivo XLSX de asistencia'

if __name__ == '__main__':
    ColumnComparisonApp().run()
